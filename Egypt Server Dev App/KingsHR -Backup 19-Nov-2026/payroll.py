"""Payroll calculation engine — identical logic to the desktop app."""
import calendar as _cal
from datetime import date as _date, timedelta as _td
from db import qy

INACTIVE_STATUSES = ('Inactive', 'Resigned', 'Terminated')
ACTIVE_EMP_SQL = ("ISNULL(EmploymentStatus,'Active') NOT IN ("
                  + ",".join(f"'{s}'" for s in INACTIVE_STATUSES) + ")")

DEFAULT_PAYROLL_SETTINGS = {
    'DayDivisor': 30, 'HourDivisor': 8, 'SICap': 16700,
    'SIRate': 0.11, 'LatenessDivisor': 480,
    'PersonalExemptionAnnual': 20000,
    'AnnualExemption1': 40000, 'AnnualExemption2': 55000,
    'TaxRate1': 0.10, 'TaxRate2': 0.15, 'TaxRate3': 0.20,
    'TaxRate4': 0.225, 'TaxRate5': 0.25, 'TaxRate6': 0.275, 'TaxRate7': 0.30,
    'TaxBracket1': 40000, 'TaxBracket2': 55000,
}
SETTING_COLS = list(DEFAULT_PAYROLL_SETTINGS.keys())

def get_payroll_settings(year=None):
    year = year or _date.today().year
    rows = qy("SELECT * FROM PayrollSettings WHERE PayrollYear<=? ORDER BY PayrollYear DESC", (year,))
    if not rows:
        return DEFAULT_PAYROLL_SETTINGS.copy()
    cols_r = qy("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='PayrollSettings' ORDER BY ORDINAL_POSITION")
    col_names = [r[0] for r in cols_r]
    row = rows[0]
    d = DEFAULT_PAYROLL_SETTINGS.copy()
    for k in SETTING_COLS:
        if k in col_names:
            idx = col_names.index(k)
            if idx < len(row) and row[idx] is not None:
                d[k] = float(row[idx])
    return d

EARNING_COLS = ['BasicSalary','HousingAllowance','TransportAllowance','OtherFixedAllowances',
                'VariableAllowances','Overtime','Bonuses','Commissions',
                'OtherEarnings','FlightAllowance',
                'Reimbursement','ReallocationAllowance']

def calc_day_deductions(basic, absent_days, late_minutes, sick_days, sanction_days, other_days, s,
                        sanction_amount=None, other_amount=None):
    """Calculate day-based deductions.

    sanction_amount / other_amount: if provided as a direct EGP value (not None),
    those override the days-based calculation for that item.
    """
    div = float(s.get('DayDivisor', 30))
    day_rate = basic / div if div else 0
    late_div = float(s.get('LatenessDivisor', 480))
    late_rate = day_rate / late_div if late_div else 0

    # Sanctions: direct EGP if sanction_amount provided, else days × day_rate
    if sanction_amount is not None:
        sanctions_egp = float(sanction_amount or 0)
        sanction_days_eff = sanctions_egp / day_rate if day_rate else 0
    else:
        sanctions_egp = day_rate * float(sanction_days or 0)
        sanction_days_eff = float(sanction_days or 0)

    # Other deductions: direct EGP if other_amount provided, else days × day_rate
    if other_amount is not None:
        other_egp = float(other_amount or 0)
        other_days_eff = other_egp / day_rate if day_rate else 0
    else:
        other_egp = day_rate * float(other_days or 0)
        other_days_eff = float(other_days or 0)

    absences_egp = day_rate * float(absent_days or 0)
    lateness_egp = late_rate * float(late_minutes or 0)
    sick_egp     = day_rate * float(sick_days or 0)

    deductibles = absences_egp + lateness_egp + sick_egp + sanctions_egp + other_egp

    return {
        'day_rate':   day_rate,
        'absences':   absences_egp,
        'lateness':   lateness_egp,
        'sick':       sick_egp,
        'sanctions':  sanctions_egp,
        'other':      other_egp,
        'deductibles': deductibles,
    }

def calc_egypt_tax(gross, basic, deductibles, s):
    si_cap  = float(s.get('SICap', 16700))
    si_rate = float(s.get('SIRate', 0.11))
    si_base    = min(basic, si_cap)
    si_monthly = si_base * si_rate
    pe_annual  = float(s.get('PersonalExemptionAnnual', 20000))
    pe_monthly = pe_annual / 12
    net_before_tax  = max(0, gross - deductibles)          # gross minus day-deductibles
    taxable_monthly = max(0, net_before_tax - si_monthly - pe_monthly)
    taxable_annual  = taxable_monthly * 12

    # Egypt 2024 brackets (cumulative upper limits, not band sizes):
    #  0 – 40,000        → 0%  (exempt)
    #  40,001 – 55,000   → 10%
    #  55,001 – 70,000   → 15%
    #  70,001 – 200,000  → 20%
    #  200,001 – 400,000 → 22.5%
    #  400,001+          → 25%
    BRACKETS = [
        (40000,  0.00, '0 – 40,000'),
        (15000,  0.10, '40,000 – 55,000'),
        (15000,  0.15, '55,000 – 70,000'),
        (130000, 0.20, '70,000 – 200,000'),
        (200000, 0.225,'200,000 – 400,000'),
        (None,   0.25, '400,000+'),
    ]
    bracket_detail = []
    remaining = taxable_annual
    annual_tax = 0.0
    for band_size, rate, label in BRACKETS:
        if remaining <= 0:
            bracket_detail.append({'label': label, 'rate': rate, 'tax': 0.0, 'base': 0.0})
            continue
        chunk = min(remaining, band_size) if band_size else remaining
        tax_in_band = chunk * rate
        annual_tax += tax_in_band
        remaining  -= chunk
        bracket_detail.append({'label': label, 'rate': rate, 'tax': tax_in_band, 'base': chunk})

    monthly_tax = annual_tax / 12
    return {
        'si_monthly':      si_monthly,
        'pe_monthly':      pe_monthly,
        'net_before_tax':  net_before_tax,
        'taxable_monthly': taxable_monthly,
        'taxable_annual':  taxable_annual,
        'annual_tax':      annual_tax,
        'monthly_tax':     monthly_tax,
        'brackets':        bracket_detail,
    }

def run_payroll(earnings, absent_days=0, late_minutes=0, sick_days=0,
                sanction_days=0, other_days=0, loan=0, settings=None,
                sanction_amount=None, other_amount=None):
    """Run payroll calculation.

    sanction_amount / other_amount: pass as direct EGP amounts (not days).
    When provided they override the days-based calculation for that item.
    """
    s = settings or get_payroll_settings()
    basic = float(earnings.get('BasicSalary') or 0)
    gross = sum(float(earnings.get(k) or 0) for k in EARNING_COLS)
    d = calc_day_deductions(basic, absent_days, late_minutes, sick_days, sanction_days, other_days, s,
                            sanction_amount=sanction_amount, other_amount=other_amount)
    t = calc_egypt_tax(gross, basic, d['deductibles'], s)
    # SI is used to reduce taxable income (inside calc_egypt_tax) but is NOT
    # deducted again as a separate line — matches the Excel logic.
    # Total deductions = day-based deductions + income tax + loan only.
    total_ded = d['absences'] + d['lateness'] + d['sick'] + d['sanctions'] + d['other'] + t['monthly_tax'] + float(loan or 0)
    net = gross - total_ded
    return {
        'gross': gross, 'monthly_tax': t['monthly_tax'],
        'tax': t, 'loan': float(loan or 0),
        'total_ded': total_ded, 'net': net, 'ded': d,
        'inputs': {'absent_days': float(absent_days or 0),
                   'late_minutes': float(late_minutes or 0),
                   'sick_days': float(sick_days or 0),
                   'sanction_days': float(sanction_days or 0),
                   'sanction_amount': float(sanction_amount or 0) if sanction_amount is not None else None,
                   'other_days': float(other_days or 0),
                   'other_amount': float(other_amount or 0) if other_amount is not None else None},
        'settings': s,
    }

def load_emp_basic(eid):
    cols = ','.join(f'ISNULL({c},0)' for c in EARNING_COLS)
    r = qy(f"""SELECT EmpID, ISNULL(FirstName,'')+' '+ISNULL(FamilyName,''), {cols},
               ISNULL(WorkSchedule,''), DateOfJoining, DateOfBirth, ISNULL(EmployeeID,''),
               ISNULL(Department,''), ISNULL(JobTitle,''), ISNULL(PersonalEmail,'')
               FROM Employees WHERE EmpID=?""", (eid,))
    if not r: return None
    r = r[0]; n = len(EARNING_COLS)
    earn = {c: float(r[2 + i] or 0) for i, c in enumerate(EARNING_COLS)}
    return {
        'eid': r[0], 'name': r[1], 'earnings': earn,
        'schedule': r[2 + n], 'join_date': r[3 + n],
        'dob': r[4 + n], 'emp_id': r[5 + n],
        'dept': r[6 + n], 'job': r[7 + n], 'email': r[8 + n],
        'settings': get_payroll_settings(),
    }

def valid_month(s):
    from datetime import datetime
    try: datetime.strptime(s, '%Y-%m'); return True
    except Exception: return False
