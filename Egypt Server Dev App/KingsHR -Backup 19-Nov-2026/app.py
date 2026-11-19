"""Kings School Egypt — HR & Payroll Web App (Flask).
Multi-user, browser-based. Mirrors the desktop app logic.
"""
from flask import (Flask, render_template, request, redirect, url_for,
                   session, jsonify, send_file, flash)
from functools import wraps
from datetime import date, timedelta, datetime
import io, os, json

from db import qy, ex, has_col
from payroll import (run_payroll, get_payroll_settings, load_emp_basic,
                     ACTIVE_EMP_SQL, EARNING_COLS, valid_month)

import hashlib, hmac as _hmac

# Ensure openpyxl is available (Excel export)
try:
    import openpyxl as _openpyxl_check  # noqa: F401
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl', '--quiet'])
_PW_ITER = 120_000

def verify_pw(stored, plain):
    """Verify password — supports pbkdf2 hashed (pb$salt$hash) and legacy plain-text."""
    stored = stored or ''
    if stored.startswith('pb$'):
        try:
            _, salt, h = stored.split('$')
            calc = hashlib.pbkdf2_hmac('sha256', plain.encode(), salt.encode(), _PW_ITER).hex()
            return _hmac.compare_digest(calc, h)
        except ValueError:
            return False
    return _hmac.compare_digest(stored, plain)   # legacy plain-text

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'kings-school-egypt-2024-secret')

# ─── Startup DB migrations (idempotent) ───────────────────────────────────────
def _run_migrations():
    """Add any missing columns to existing tables on startup."""
    try:
        if not has_col('Employees', 'FlightAllowance'):
            ex("ALTER TABLE Employees ADD FlightAllowance DECIMAL(10,2) NULL")
    except Exception:
        pass
    # Monthly variable earnings columns on PayrollRuns
    for col, typ in [
        ('VariableAllowances',    'DECIMAL(10,2)'),
        ('Overtime',              'DECIMAL(10,2)'),
        ('Bonuses',               'DECIMAL(10,2)'),
        ('Commissions',           'DECIMAL(10,2)'),
        ('Reimbursement',         'DECIMAL(10,2)'),
        ('ReallocationAllowance', 'DECIMAL(10,2)'),
        ('FlightAllowance',       'DECIMAL(10,2)'),  # moved from Employees to monthly PayrollRuns
    ]:
        try:
            if not has_col('PayrollRuns', col):
                ex(f"ALTER TABLE PayrollRuns ADD {col} {typ} NULL")
        except Exception:
            pass
    # Issue 6: Expat / Egyptian employee type + GBP salary + Nationality columns
    for col, typ in [
        ('EmployeeType', "VARCHAR(20)"),    # 'Egyptian' or 'Expat'
        ('GBPSalary',    'DECIMAL(12,2)'),  # GBP salary amount for expats
        ('Nationality',  "VARCHAR(50)"),    # Employee nationality (determines Expat logic)
    ]:
        try:
            if not has_col('Employees', col):
                ex(f"ALTER TABLE Employees ADD {col} {typ} NULL")
        except Exception:
            pass
    # GBP/EGP exchange rate + GBP/EGP components stored per payroll run
    for col, typ in [
        ('GBPEGPRate',    'DECIMAL(12,4)'),  # exchange rate used for the month
        ('GBPComponent',  'DECIMAL(12,2)'),  # 75% portion paid in GBP
        ('EGPComponent',  'DECIMAL(12,2)'),  # 25% portion paid in EGP
    ]:
        try:
            if not has_col('PayrollRuns', col):
                ex(f"ALTER TABLE PayrollRuns ADD {col} {typ} NULL")
        except Exception:
            pass
    # Issues 1 & 2: GBP monthly variable earnings and deductions for Expat employees
    for col, typ in [
        ('GBPVariableAllowances',    'DECIMAL(12,2)'),  # GBP variable allowances
        ('GBPOvertime',              'DECIMAL(12,2)'),  # GBP overtime
        ('GBPBonuses',               'DECIMAL(12,2)'),  # GBP bonus earned this month
        ('GBPCommissions',           'DECIMAL(12,2)'),  # GBP commissions / other earnings
        ('GBPFlightAllowance',       'DECIMAL(12,2)'),  # GBP flight allowance
        ('GBPReimbursement',         'DECIMAL(12,2)'),  # GBP reimbursement
        ('GBPReallocationAllowance', 'DECIMAL(12,2)'),  # GBP reallocation allowance
        ('GBPOtherEarnings',         'DECIMAL(12,2)'),  # GBP other/misc earnings
        ('GBPLoanDeduction',         'DECIMAL(12,2)'),  # GBP loan repayment
        ('GBPOtherDeductions',       'DECIMAL(12,2)'),  # GBP other deductions
        ('GBPAbsentDays',            'DECIMAL(6,2)'),   # GBP absence days
        ('GBPLateMinutes',           'DECIMAL(8,2)'),   # GBP late/early minutes
        ('GBPSickDays',              'DECIMAL(6,2)'),   # GBP sick leave days
        ('GBPSanctionDays',          'DECIMAL(6,2)'),   # GBP sanction days
    ]:
        try:
            if not has_col('PayrollRuns', col):
                ex(f"ALTER TABLE PayrollRuns ADD {col} {typ} NULL")
        except Exception:
            pass
    # Request 4: IsManualEntry flag — distinguish manual entries from bulk-run entries
    try:
        if not has_col('PayrollRuns', 'IsManualEntry'):
            ex("ALTER TABLE PayrollRuns ADD IsManualEntry BIT NOT NULL DEFAULT 0")
    except Exception:
        pass
_run_migrations()

import logging
logging.basicConfig(
    filename=os.path.join(os.path.dirname(__file__), 'logs', 'app.log'),
    level=logging.WARNING,
    format='%(asctime)s %(levelname)s %(message)s'
)
os.makedirs(os.path.join(os.path.dirname(__file__), 'logs'), exist_ok=True)

def _detect_users_cols():
    """Auto-detect password and active column names in the Users table."""
    try:
        cols_r = qy("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='Users'")
        cols = [r[0] for r in cols_r] if cols_r else []
        cols_lower = [c.lower() for c in cols]
        app.config['_all_user_cols'] = cols  # store all cols for other queries
        # Password column
        for candidate in ('PasswordHash', 'Password', 'UserPassword', 'Pass', 'Pwd'):
            if candidate.lower() in cols_lower:
                app.config['_pwd_col'] = cols[cols_lower.index(candidate.lower())]
                break
        else:
            app.config['_pwd_col'] = 'Password'
        # IsActive column
        for candidate in ('IsActive', 'Active', 'Enabled', 'Status'):
            if candidate.lower() in cols_lower:
                app.config['_active_col'] = cols[cols_lower.index(candidate.lower())]
                break
        else:
            app.config['_active_col'] = None
    except Exception:
        app.config['_pwd_col']    = 'PasswordHash'
        app.config['_active_col'] = None

# ─── Auth helpers ────────────────────────────────────────────────────────────

ROLE_LABELS = {
    'Admin': 'Admin',
    'HR':    'HR Manager',
    'IT':    'IT Support',
    'Ops':   'Operations',
    'Emp':   'Employee',
}

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def role_required(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if session.get('role') not in roles:
                flash('Access denied.', 'error')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return decorated
    return decorator

# ─── Routes: Auth ─────────────────────────────────────────────────────────────

@app.route('/', methods=['GET', 'POST'])
@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'user' in session:
        return redirect(url_for('dashboard'))
    error = None
    if request.method == 'POST':
        uname = request.form.get('username', '').strip()
        pwd   = request.form.get('password', '').strip()
        # Detect actual column names once and cache on app config
        if not app.config.get('_users_cols_checked'):
            _detect_users_cols()
            app.config['_users_cols_checked'] = True
        pwd_col    = app.config.get('_pwd_col', 'PasswordHash')
        active_col = app.config.get('_active_col', None)
        # Check if EmpID column exists
        all_cols = app.config.get('_all_user_cols', [])
        emp_id_sel = "ISNULL(EmpID,0)" if 'empid' in [c.lower() for c in all_cols] else "0"
        # Fetch user by username only, then verify password in Python (supports hashed + plain-text)
        if active_col:
            sql = (f"SELECT UserID, Username, Role, ISNULL(FullName,Username), {emp_id_sel}, {pwd_col} "
                   f"FROM Users WHERE Username=? AND ISNULL({active_col},1)=1")
        else:
            sql = (f"SELECT UserID, Username, Role, ISNULL(FullName,Username), {emp_id_sel}, {pwd_col} "
                   f"FROM Users WHERE Username=?")
        rows = qy(sql, (uname,))
        app.logger.warning(f"LOGIN DEBUG: user={uname!r} pwd_col={pwd_col!r} active_col={active_col!r} rows={rows!r}")
        matched = next((r for r in rows if verify_pw(str(r[5] or ''), pwd)), None) if rows else None
        if rows:
            for r in rows:
                stored = str(r[5] or '')
                app.logger.warning(f"LOGIN DEBUG: stored_pw={stored!r} verify={verify_pw(stored, pwd)!r}")
        if matched:
            r = matched
            role_raw = str(r[2] or '').strip()
            # Normalise role to Title case (DB may store 'admin' or 'Admin')
            role_map = {k.lower(): k for k in ROLE_LABELS}
            role = role_map.get(role_raw.lower(), role_raw.title())
            session['user_id']  = r[0]
            session['user']     = r[1]
            session['role']     = role
            session['fullname'] = r[3]
            session['emp_id']   = r[4]
            return redirect(url_for('dashboard'))
        else:
            error = 'Invalid username or password.'
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ─── Routes: Dashboard ────────────────────────────────────────────────────────

@app.route('/dashboard')
@login_required
def dashboard():
    role = session.get('role', '')
    # Quick stats
    emp_count = (qy(f"SELECT COUNT(*) FROM Employees WHERE {ACTIVE_EMP_SQL}") or [(0,)])[0][0]
    this_month = date.today().strftime('%Y-%m')
    payroll_run = (qy("SELECT COUNT(DISTINCT EmpID) FROM PayrollRuns WHERE PayrollMonth=?",
                      (this_month,)) or [(0,)])[0][0]
    return render_template('dashboard.html',
                           emp_count=emp_count, payroll_run=payroll_run,
                           this_month=this_month)

# ─── Routes: Employees ────────────────────────────────────────────────────────

@app.route('/employees')
@login_required
@role_required('Admin', 'HR')
def employees():
    search              = request.args.get('q', '').strip()
    status_filter       = request.args.get('status_filter', '').strip()
    dept_filter         = request.args.get('dept_filter', '').strip()
    title_filter        = request.args.get('title_filter', '').strip()
    nationality_filter  = request.args.get('nationality_filter', '').strip()
    mobile_filter       = request.args.get('mobile_filter', '').strip()
    join_from           = request.args.get('join_from', '').strip()
    join_to             = request.args.get('join_to', '').strip()

    has_filter = any([search, status_filter, dept_filter, title_filter,
                      nationality_filter, mobile_filter, join_from, join_to])

    # Build WHERE clause
    conditions = []
    params = []
    if search:
        conditions.append("(FirstName LIKE ? OR FamilyName LIKE ? OR EmployeeID LIKE ?)")
        params += [f'%{search}%', f'%{search}%', f'%{search}%']
    if status_filter:
        conditions.append("ISNULL(EmploymentStatus,'Active') = ?")
        params.append(status_filter)
    # When no status filter, default to active only (original behaviour)
    if not has_filter:
        conditions.append(ACTIVE_EMP_SQL)
    if dept_filter:
        conditions.append("ISNULL(Department,'') = ?")
        params.append(dept_filter)
    if title_filter:
        conditions.append("ISNULL(JobTitle,'') = ?")
        params.append(title_filter)
    if nationality_filter:
        conditions.append("ISNULL(Nationality,'') = ?")
        params.append(nationality_filter)
    if mobile_filter:
        conditions.append("ISNULL(Phone,'') LIKE ?")
        params.append(f'%{mobile_filter}%')
    if join_from:
        conditions.append("DateOfJoining >= ?")
        params.append(join_from)
    if join_to:
        conditions.append("DateOfJoining <= ?")
        params.append(join_to)

    where = ('WHERE ' + ' AND '.join(conditions)) if conditions else ''
    # Total salary = sum of all pay components (use ISNULL to treat NULLs as 0)
    total_sal = """(
        ISNULL(BasicSalary,0)+ISNULL(HousingAllowance,0)+ISNULL(TransportAllowance,0)+
        ISNULL(OtherFixedAllowances,0)+ISNULL(VariableAllowances,0)+ISNULL(Overtime,0)+
        ISNULL(Bonuses,0)+ISNULL(Commissions,0)+
        ISNULL(CASE WHEN COL_LENGTH('Employees','Reimbursement') IS NOT NULL THEN Reimbursement ELSE 0 END,0)+
        ISNULL(CASE WHEN COL_LENGTH('Employees','ReallocationAllowance') IS NOT NULL THEN ReallocationAllowance ELSE 0 END,0)+
        ISNULL(CASE WHEN COL_LENGTH('Employees','OtherEarnings') IS NOT NULL THEN OtherEarnings ELSE 0 END,0)
        -- FlightAllowance excluded: entered monthly in payroll, not stored on employee
    )"""
    sql = f"""SELECT EmpID, ISNULL(EmployeeID,''), ISNULL(FirstName,'')+' '+ISNULL(FamilyName,''),
              ISNULL(Department,''), ISNULL(JobTitle,''),
              {total_sal},
              ISNULL(EmploymentStatus,'Active'), CONVERT(VARCHAR(10), DateOfJoining, 120),
              ISNULL(Nationality,''), ISNULL(Phone,'')
              FROM Employees {where} ORDER BY FirstName"""
    rows = qy(sql, params) if params else qy(sql)
    emps = [{'eid': r[0], 'emp_id': r[1], 'name': r[2],
              'dept': r[3], 'job': r[4], 'total_salary': r[5], 'status': r[6],
              'join_date': r[7] or '', 'nationality': r[8], 'phone': r[9]} for r in rows]

    # Populate filter dropdowns from distinct DB values (always all employees, unfiltered)
    def _distinct(col):
        r = qy(f"SELECT DISTINCT ISNULL({col},'') FROM Employees WHERE ISNULL({col},'')<>'' ORDER BY 1")
        return [x[0] for x in r] if r else []

    dept_list        = _distinct('Department')
    title_list       = _distinct('JobTitle')
    nationality_list = _distinct('Nationality')

    return render_template('employees.html', emps=emps, search=search,
                           status_filter=status_filter,
                           dept_filter=dept_filter, title_filter=title_filter,
                           nationality_filter=nationality_filter,
                           mobile_filter=mobile_filter,
                           join_from=join_from, join_to=join_to,
                           dept_list=dept_list, title_list=title_list,
                           nationality_list=nationality_list)

@app.route('/employees/new', methods=['GET', 'POST'])
@login_required
@role_required('Admin', 'HR')
def employee_new():
    if request.method == 'POST':
        return _save_employee(None)
    return render_template('employee_form.html', emp={}, title='New Employee')

@app.route('/employees/<int:eid>/edit', methods=['GET', 'POST'])
@login_required
@role_required('Admin', 'HR')
def employee_edit(eid):
    if request.method == 'POST':
        return _save_employee(eid)

    def _d(v):
        """Return date as dd/mm/yyyy for display in the form."""
        if not v: return ''
        s = str(v)[:10]  # 'yyyy-mm-dd' from SQL Server
        try:
            from datetime import datetime as _dt
            return _dt.strptime(s, '%Y-%m-%d').strftime('%d/%m/%Y')
        except Exception:
            return s
    def _hc(col): return has_col('Employees', col)

    # Build SELECT dynamically based on which new columns exist
    extra_select = []
    extra_map    = []

    for col, key in [
        ('SecondName','second_name'),('ThirdName','third_name'),
        ('FullNameArabic','full_name_arabic'),('MaritalStatus','marital_status'),
        ('Address','address'),('SchoolEmail','school_email'),
        ('EmploymentType','employment_type'),('ContractStartDate','contract_start_date'),
        ('ContractEndDate','contract_end_date'),('ConfirmationDate','confirmation_date'),
        ('ResignDate','resign_date'),('ReportingManager','reporting_manager'),
        ('BranchLocation','branch_location'),('ProbationPeriod','probation_period'),
        ('WorkingHoursPerDay','working_hours_per_day'),
        ('SalaryType','salary_type'),('VariableAllowances','variable_allowances'),
        ('Overtime','overtime'),('Bonuses','bonuses'),('Commissions','commissions'),
        ('BankName','bank_name'),('BankAccountNumber','bank_account_number'),('IBAN','iban'),
        ('TaxableEmployee','taxable_employee'),('TaxRegistrationID','tax_registration_id'),
        ('TaxTreatment','tax_treatment'),('TaxableBenefits','taxable_benefits'),
        ('OtherTaxExemptions','other_tax_exemptions'),
        ('AnnualTaxAdjustment','annual_tax_adjustment'),('TaxStatus','tax_status'),
        ('InsuranceNumber','insurance_number'),('InsuranceOffice','insurance_office'),
        ('InsuranceStartDate','insurance_start_date'),('InsuranceWage','insurance_wage'),
        ('InsuranceStatus','insurance_status'),('InsuranceCategory','insurance_category'),
        ('PreviousInsurancePeriod','previous_insurance_period'),
        ('PreviousEmployer','previous_employer'),('PreviousPensionInfo','previous_pension_info'),
        ('Reimbursement','reimbursement'),('ReallocationAllowance','reallocation_allowance'),
        ('OtherEarnings','other_earnings'),
        # FlightAllowance moved to payroll monthly earnings
        ('DocNationalID','doc_national_id'),('DocBirthCert','doc_birth_cert'),
        ('DocEmploymentContract','doc_employment_contract'),
        ('DocEducationalCert','doc_educational_cert'),('DocSyndicate','doc_syndicate'),
        ('DocMedicalForm111','doc_medical_form_111'),
        ('DocSocialInsurance','doc_social_insurance'),('DocBankDocument','doc_bank_document'),
        ('MedicalFitnessCert','medical_fitness_cert'),
        ('EmployeeType','employee_type'),  # Issue 6
        ('GBPSalary','gbp_salary'),        # Issue 6
    ]:
        if _hc(col):
            extra_select.append(f"ISNULL({col},'')" if col not in (
                'ContractStartDate','ContractEndDate','ConfirmationDate',
                'ResignDate','InsuranceStartDate','TaxableEmployee',
                'DocNationalID','DocBirthCert','DocEmploymentContract',
                'DocEducationalCert','DocSyndicate','DocMedicalForm111',
                'DocSocialInsurance','DocBankDocument','MedicalFitnessCert',
                'ProbationPeriod','WorkingHoursPerDay','VariableAllowances',
                'Overtime','Bonuses','Commissions','InsuranceWage',
                'TaxableBenefits','OtherTaxExemptions','AnnualTaxAdjustment',
                'PreviousInsurancePeriod','Reimbursement','ReallocationAllowance',
                'OtherEarnings','GBPSalary') else col)
            extra_map.append((key, col in (
                'ContractStartDate','ContractEndDate','ConfirmationDate',
                'ResignDate','InsuranceStartDate')))

    extra_sql = (', ' + ', '.join(extra_select)) if extra_select else ''
    row = qy(f"""SELECT EmpID, ISNULL(EmployeeID,''), ISNULL(FirstName,''), ISNULL(FamilyName,''),
                ISNULL(Department,''), ISNULL(JobTitle,''), ISNULL(BasicSalary,0),
                ISNULL(HousingAllowance,0), ISNULL(TransportAllowance,0),
                ISNULL(OtherFixedAllowances,0), ISNULL(PersonalEmail,''),
                ISNULL(WorkSchedule,''), DateOfJoining, DateOfBirth,
                ISNULL(EmploymentStatus,'Active'), ISNULL(NationalID,''),
                ISNULL(Phone,''), ISNULL(Gender,''), ISNULL(Nationality,'')
                {extra_sql}
                FROM Employees WHERE EmpID=?""", (eid,))
    if not row:
        flash('Employee not found.', 'error')
        return redirect(url_for('employees'))
    r = row[0]
    emp = {
        'eid': r[0], 'emp_id': r[1], 'first_name': r[2], 'family_name': r[3],
        'dept': r[4], 'job': r[5], 'basic': r[6], 'housing': r[7], 'transport': r[8],
        'other_fixed': r[9], 'email': r[10], 'schedule': r[11],
        'join_date': _d(r[12]), 'dob': _d(r[13]),
        'status': r[14], 'national_id': r[15], 'phone': r[16],
        'gender': r[17], 'nationality': r[18],
    }
    # Map extra columns
    for i, (key, is_date) in enumerate(extra_map):
        val = r[19 + i]
        emp[key] = _d(val) if is_date else (bool(val) if key.startswith('doc_') or key == 'medical_fitness_cert' or key == 'taxable_employee' else val)

    # Compute read-only tax summary fields using same Egypt formula as payroll
    from payroll import calc_egypt_tax, get_payroll_settings
    _s = get_payroll_settings()
    _basic = float(emp.get('basic') or 0)
    _gross = sum(float(emp.get(k) or 0) for k in [
        'basic','housing','transport','other_fixed',
        'variable_allowances','overtime','bonuses','commissions',
        'other_earnings','reimbursement','reallocation_allowance',
        # flight_allowance excluded: entered monthly in payroll earnings tab
    ])
    _t = calc_egypt_tax(_gross, _basic, 0, _s)  # deductibles=0 for base salary view
    emp['calc_taxable_salary'] = _t['taxable_monthly']
    emp['calc_monthly_tax']    = _t['monthly_tax']
    emp['calc_si_deduction']   = _t['si_monthly']
    emp['calc_annual_tax']     = _t['annual_tax']

    return render_template('employee_form.html', emp=emp, title='Edit Employee')

def _save_employee(eid):
    f = request.form
    def fv(k, default=''):
        v = f.get(k, default)
        return v.strip() if isinstance(v, str) else v
    def fnum(k):
        v = fv(k)
        try: return float(v) if v else 0
        except: return 0
    def fdate(k):
        """Parse dd/mm/yyyy from form → yyyy-mm-dd for SQL Server."""
        v = fv(k)
        if not v: return None
        # Accept dd/mm/yyyy or already yyyy-mm-dd
        for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y'):
            try:
                from datetime import datetime as _dt
                return _dt.strptime(v, fmt).strftime('%Y-%m-%d')
            except ValueError:
                continue
        return v  # pass through if unrecognised
    def fbit(k):
        return 1 if f.get(k) else 0

    # Core fields always present
    fields = {
        'EmployeeID': fv('emp_id'), 'FirstName': fv('first_name'),
        'FamilyName': fv('family_name'), 'Department': fv('dept'),
        'JobTitle': fv('job'), 'BasicSalary': fnum('basic'),
        'HousingAllowance': fnum('housing'), 'TransportAllowance': fnum('transport'),
        'OtherFixedAllowances': fnum('other_fixed'),
        'PersonalEmail': fv('email'), 'WorkSchedule': fv('schedule'),
        'NationalID': fv('national_id'), 'Phone': fv('phone'),
        'Gender': fv('gender'), 'Nationality': fv('nationality'),
        'EmploymentStatus': fv('status', 'Active'),
    }
    date_fields = {
        'DateOfJoining': fdate('join_date'),
        'DateOfBirth': fdate('dob'),
    }

    # Optional new columns — only add if column exists in DB
    optional_str = [
        ('SecondName','second_name'),('ThirdName','third_name'),
        ('FullNameArabic','full_name_arabic'),('MaritalStatus','marital_status'),
        ('Address','address'),('SchoolEmail','school_email'),
        ('EmploymentType','employment_type'),('ReportingManager','reporting_manager'),
        ('BranchLocation','branch_location'),('SalaryType','salary_type'),
        ('BankName','bank_name'),('BankAccountNumber','bank_account_number'),('IBAN','iban'),
        ('TaxRegistrationID','tax_registration_id'),('TaxTreatment','tax_treatment'),
        ('TaxStatus','tax_status'),('InsuranceNumber','insurance_number'),
        ('InsuranceOffice','insurance_office'),('InsuranceStatus','insurance_status'),
        ('InsuranceCategory','insurance_category'),('PreviousEmployer','previous_employer'),
        ('PreviousPensionInfo','previous_pension_info'),
        ('EmployeeType','employee_type'),  # Issue 6: 'Egyptian' or 'Expat'
    ]
    optional_num = [
        ('ProbationPeriod','probation_period'),('WorkingHoursPerDay','working_hours_per_day'),
        ('VariableAllowances','variable_allowances'),('Overtime','overtime'),
        ('Bonuses','bonuses'),('Commissions','commissions'),('InsuranceWage','insurance_wage'),
        ('TaxableBenefits','taxable_benefits'),('OtherTaxExemptions','other_tax_exemptions'),
        ('AnnualTaxAdjustment','annual_tax_adjustment'),
        ('PreviousInsurancePeriod','previous_insurance_period'),
        ('Reimbursement','reimbursement'),
        ('ReallocationAllowance','reallocation_allowance'),
        ('OtherEarnings','other_earnings'),
        # FlightAllowance removed: now a monthly payroll field, not saved on employee
        ('GBPSalary','gbp_salary'),  # Issue 6: GBP salary for expats
    ]
    optional_date = [
        ('ContractStartDate','contract_start_date'),('ContractEndDate','contract_end_date'),
        ('ConfirmationDate','confirmation_date'),('ResignDate','resign_date'),
        ('InsuranceStartDate','insurance_start_date'),
    ]
    optional_bit = [
        ('TaxableEmployee','taxable_employee'),
        ('DocNationalID','doc_national_id'),('DocBirthCert','doc_birth_cert'),
        ('DocEmploymentContract','doc_employment_contract'),
        ('DocEducationalCert','doc_educational_cert'),('DocSyndicate','doc_syndicate'),
        ('DocMedicalForm111','doc_medical_form_111'),
        ('DocSocialInsurance','doc_social_insurance'),('DocBankDocument','doc_bank_document'),
        ('MedicalFitnessCert','medical_fitness_cert'),
    ]

    for col, form_key in optional_str:
        if has_col('Employees', col):
            fields[col] = fv(form_key)
    for col, form_key in optional_num:
        if has_col('Employees', col):
            fields[col] = fnum(form_key)
    for col, form_key in optional_date:
        if has_col('Employees', col):
            date_fields[col] = fdate(form_key)
    for col, form_key in optional_bit:
        if has_col('Employees', col):
            fields[col] = fbit(form_key)

    try:
        if eid:
            all_fields = {**fields, **date_fields}
            sets = ', '.join(f"{k}=?" for k in all_fields)
            vals = list(all_fields.values()) + [eid]
            ex(f"UPDATE Employees SET {sets} WHERE EmpID=?", vals)
            flash('Employee updated.', 'success')
        else:
            all_fields = {**fields, **date_fields}
            keys = list(all_fields.keys())
            vals = list(all_fields.values())
            placeholders = ','.join('?' * len(keys))
            ex(f"INSERT INTO Employees ({','.join(keys)}) VALUES ({placeholders})", vals)
            flash('Employee added.', 'success')
    except Exception as e:
        flash(f'Error: {e}', 'error')
    return redirect(url_for('employees'))

@app.route('/employees/<int:eid>/delete', methods=['POST'])
@login_required
@role_required('Admin', 'HR')
def employee_delete(eid):
    # Delete is disabled — this route kept for URL compatibility but does nothing
    flash('Employee deletion is disabled. To deactivate an employee, set their status to Inactive.', 'info')
    return redirect(url_for('employees'))

import csv, io as _io

def _parse_csv_date(v):
    """Accept dd/mm/yyyy, dd-mm-yyyy or yyyy-mm-dd from CSV → yyyy-mm-dd for SQL Server."""
    v = (v or '').strip()
    if not v: return None
    for fmt in ('%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%m/%d/%Y'):
        try:
            return datetime.strptime(v, fmt).strftime('%Y-%m-%d')
        except ValueError:
            continue
    return None  # unrecognised — store NULL rather than blow up

@app.route('/employees/bulk-upload', methods=['GET', 'POST'])
@login_required
@role_required('Admin', 'HR')
def employee_bulk_upload():
    """Bulk CSV upload for employees.
    Expected CSV columns (order-independent, header row required):
    EmployeeID, FirstName, FamilyName, Department, JobTitle, BasicSalary,
    HousingAllowance, TransportAllowance, OtherFixedAllowances,
    PersonalEmail, WorkSchedule, DateOfJoining, DateOfBirth,
    EmploymentStatus, NationalID, Phone, Gender, Nationality
    """
    results = []
    if request.method == 'POST':
        uploaded = request.files.get('csv_file')
        if not uploaded or not uploaded.filename.lower().endswith('.csv'):
            flash('Please upload a CSV file.', 'error')
            return redirect(url_for('employee_bulk_upload'))

        text = uploaded.read().decode('utf-8-sig', errors='replace')
        reader = csv.DictReader(_io.StringIO(text))
        ok = err = skip = 0
        for i, row in enumerate(reader, start=2):
            def rv(k, d=''):
                return (row.get(k) or '').strip() or d
            emp_id = rv('EmployeeID')
            first  = rv('FirstName')
            family = rv('FamilyName')
            if not first and not family:
                skip += 1
                continue
            try:
                ex("""INSERT INTO Employees
                      (EmployeeID,FirstName,FamilyName,Department,JobTitle,
                       BasicSalary,HousingAllowance,TransportAllowance,OtherFixedAllowances,
                       PersonalEmail,WorkSchedule,DateOfJoining,DateOfBirth,
                       EmploymentStatus,NationalID,Phone,Gender,Nationality)
                      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                   (emp_id, first, family,
                    rv('Department'), rv('JobTitle'),
                    float(rv('BasicSalary') or 0), float(rv('HousingAllowance') or 0),
                    float(rv('TransportAllowance') or 0), float(rv('OtherFixedAllowances') or 0),
                    rv('PersonalEmail'), rv('WorkSchedule'),
                    _parse_csv_date(rv('DateOfJoining')), _parse_csv_date(rv('DateOfBirth')),
                    rv('EmploymentStatus', 'Active'),
                    rv('NationalID'), rv('Phone'), rv('Gender'), rv('Nationality')))
                ok += 1
                results.append({'row': i, 'name': f"{first} {family}", 'status': 'ok', 'msg': ''})
            except Exception as e:
                err += 1
                results.append({'row': i, 'name': f"{first} {family}", 'status': 'error', 'msg': str(e)})

        flash(f'Upload complete: {ok} added, {err} errors, {skip} skipped (blank name).', 'success' if not err else 'info')

    return render_template('employee_bulk_upload.html', results=results)

# ─── Routes: Payroll ──────────────────────────────────────────────────────────

@app.route('/payroll')
@login_required
@role_required('Admin', 'HR')
def payroll_menu():
    return render_template('payroll_menu.html')

@app.route('/payroll/manual', methods=['GET', 'POST'])
@login_required
@role_required('Admin', 'HR')
def payroll_manual():
    emps = qy(f"""SELECT EmpID, ISNULL(EmployeeID,''), ISNULL(FirstName,'')+' '+ISNULL(FamilyName,'')
                  FROM Employees WHERE {ACTIVE_EMP_SQL} ORDER BY FirstName""")
    emp_list = [{'eid': r[0], 'emp_id': r[1], 'name': r[2]} for r in emps]

    result    = None
    sel_emp   = None
    saved_inputs = {}   # pre-populate form from saved PayrollRuns row on GET
    month = (request.form.get('month') or request.args.get('month') or
             date.today().strftime('%Y-%m'))

    # Build set of employee IDs that already have a PayrollRun saved for this month
    saved_rows = qy("SELECT EmpID FROM PayrollRuns WHERE PayrollMonth=?", (month,))
    saved_eids = {r[0] for r in (saved_rows or [])}
    # Fetch the most recently used GBP/EGP rate for this month (if any run was already saved)
    gbp_rate_rows = qy(
        "SELECT TOP 1 GBPEGPRate FROM PayrollRuns WHERE PayrollMonth=? AND GBPEGPRate IS NOT NULL",
        (month,)) if has_col('PayrollRuns', 'GBPEGPRate') else None
    default_gbp_rate = float(gbp_rate_rows[0][0]) if gbp_rate_rows else 0.0
    # Annotate emp_list with done status
    for e in emp_list:
        e['done'] = (e['eid'] in saved_eids)

    def _load_history(eid):
        hist_rows = qy("""
            SELECT TOP 12 PayrollMonth, MonthlyGross,
                   ISNULL(Absences,0), ISNULL(Lateness,0), ISNULL(SickLeave,0),
                   ISNULL(Sanctions,0), ISNULL(OtherDeductions,0),
                   ISNULL(Loan,0), ISNULL(IncomeTax,0), ISNULL(SocialInsurance,0),
                   TotalDeductions, NetSalary, RunDate, ISNULL(Notes,'')
            FROM PayrollRuns WHERE EmpID=?
            ORDER BY PayrollMonth DESC
        """, (eid,))
        return [{
            'month': r[0], 'gross': r[1],
            'absences': r[2], 'lateness': r[3], 'sick': r[4],
            'sanctions': r[5], 'other': r[6], 'loan': r[7],
            'tax': r[8], 'si': r[9],
            'total_ded': r[10], 'net': r[11],
            'run_date': r[12], 'notes': r[13],
        } for r in (hist_rows or [])]

    # ── GET: pre-load employee + saved deductions if eid & month in query string ──
    if request.method == 'GET':
        eid_get = request.args.get('eid')
        if eid_get:
            emp = _enrich_emp_expat(load_emp_basic(int(eid_get)))
            if emp:
                sel_emp = emp
                # Try to restore saved deduction inputs for the selected month
                # Include GBPEGPRate if column exists
                _gbp_col  = ", ISNULL(GBPEGPRate,0)"    if has_col('PayrollRuns','GBPEGPRate')    else ", 0"
                _fa_col   = ", ISNULL(FlightAllowance,0)" if has_col('PayrollRuns','FlightAllowance') else ", 0"
                # Issues 1 & 2: GBP earn/ded columns
                _gbp_earn_col = (", ISNULL(GBPVariableAllowances,0), ISNULL(GBPOvertime,0)"
                                 ", ISNULL(GBPBonuses,0), ISNULL(GBPCommissions,0)"
                                 ", ISNULL(GBPFlightAllowance,0), ISNULL(GBPReimbursement,0)"
                                 ", ISNULL(GBPReallocationAllowance,0), ISNULL(GBPOtherEarnings,0)"
                                 ", ISNULL(GBPLoanDeduction,0), ISNULL(GBPOtherDeductions,0)"
                                 + (", ISNULL(GBPAbsentDays,0), ISNULL(GBPLateMinutes,0)"
                                    ", ISNULL(GBPSickDays,0), ISNULL(GBPSanctionDays,0)"
                                    if has_col('PayrollRuns','GBPAbsentDays') else ", 0, 0, 0, 0")
                                 ) if has_col('PayrollRuns','GBPBonuses') else ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0"
                saved = qy(f"""
                    SELECT ISNULL(AbsentDays,0), ISNULL(LateMinutes,0),
                           ISNULL(SickDays,0), ISNULL(SanctionDays,0), ISNULL(OtherDays,0),
                           ISNULL(Loan,0), ISNULL(Notes,''),
                           ISNULL(Sanctions,0), ISNULL(OtherDeductions,0),
                           ISNULL(VariableAllowances,0), ISNULL(Overtime,0),
                           ISNULL(Bonuses,0), ISNULL(Commissions,0),
                           ISNULL(Reimbursement,0), ISNULL(ReallocationAllowance,0)
                           {_gbp_col}{_fa_col}{_gbp_earn_col}
                    FROM PayrollRuns WHERE EmpID=? AND PayrollMonth=?
                """, (emp['eid'], month))
                if saved:
                    s = saved[0]
                    saved_inputs = {
                        'absent_days':          float(s[0] or 0),
                        'late_minutes':         float(s[1] or 0),
                        'sick_days':            float(s[2] or 0),
                        'sanction_days':        float(s[3] or 0),
                        'other_days':           float(s[4] or 0),
                        'loan':                 float(s[5] or 0),
                        'notes':                s[6] or '',
                        'other_amount':         float(s[8] or 0),
                        'variable_allowances':  float(s[9] or 0),
                        'overtime':             float(s[10] or 0),
                        'bonuses':              float(s[11] or 0),
                        'commissions':          float(s[12] or 0),
                        'reimbursement':        float(s[13] or 0),
                        'reallocation_allowance': float(s[14] or 0),
                        'gbp_rate':             float(s[15] or 0) if len(s) > 15 else 0.0,
                        'flight_allowance':     float(s[16] or 0) if len(s) > 16 else 0.0,
                        # GBP earn/ded (indices 17-26)
                        'gbp_variable_allowances':    float(s[17] or 0) if len(s) > 17 else 0.0,
                        'gbp_overtime':               float(s[18] or 0) if len(s) > 18 else 0.0,
                        'gbp_bonuses':                float(s[19] or 0) if len(s) > 19 else 0.0,
                        'gbp_commissions':            float(s[20] or 0) if len(s) > 20 else 0.0,
                        'gbp_flight_allowance':       float(s[21] or 0) if len(s) > 21 else 0.0,
                        'gbp_reimbursement':          float(s[22] or 0) if len(s) > 22 else 0.0,
                        'gbp_reallocation_allowance': float(s[23] or 0) if len(s) > 23 else 0.0,
                        'gbp_other_earnings':         float(s[24] or 0) if len(s) > 24 else 0.0,
                        'gbp_loan_deduction':         float(s[25] or 0) if len(s) > 25 else 0.0,
                        'gbp_other_deductions':       float(s[26] or 0) if len(s) > 26 else 0.0,
                        'gbp_absent_days':            float(s[27] or 0) if len(s) > 27 else 0.0,
                        'gbp_late_minutes':           float(s[28] or 0) if len(s) > 28 else 0.0,
                        'gbp_sick_days':              float(s[29] or 0) if len(s) > 29 else 0.0,
                        'gbp_sanction_days':          float(s[30] or 0) if len(s) > 30 else 0.0,
                    }
                    # Overlay monthly earnings on top of fixed profile earnings
                    monthly_earn = dict(emp['earnings'])
                    monthly_earn.update({
                        'VariableAllowances':    saved_inputs['variable_allowances'],
                        'Overtime':              saved_inputs['overtime'],
                        'Bonuses':               saved_inputs['bonuses'],
                        'Commissions':           saved_inputs['commissions'],
                        'Reimbursement':         saved_inputs['reimbursement'],
                        'ReallocationAllowance': saved_inputs['reallocation_allowance'],
                        'FlightAllowance':       saved_inputs.get('flight_allowance', 0),
                    })
                    # Re-run payroll with saved values so result panel shows
                    result = run_payroll(
                        monthly_earn,
                        absent_days=saved_inputs['absent_days'],
                        late_minutes=saved_inputs['late_minutes'],
                        sick_days=saved_inputs['sick_days'],
                        sanction_days=saved_inputs['sanction_days'],
                        loan=saved_inputs['loan'],
                        settings=emp['settings'],
                        other_amount=saved_inputs['other_amount'] or None,
                    )
                    # Inject monthly earn overrides into result.inputs for template display
                    result['inputs']['variable_allowances']    = saved_inputs['variable_allowances']
                    result['inputs']['overtime']               = saved_inputs['overtime']
                    result['inputs']['bonuses']                = saved_inputs['bonuses']
                    result['inputs']['commissions']            = saved_inputs['commissions']
                    result['inputs']['reimbursement']          = saved_inputs['reimbursement']
                    result['inputs']['reallocation_allowance'] = saved_inputs['reallocation_allowance']
                    result['inputs']['flight_allowance']       = saved_inputs.get('flight_allowance', 0)
                    # Issues 1 & 2: Restore GBP earn/ded into result for payslip display
                    _gbp_rate_saved = saved_inputs.get('gbp_rate', 0)
                    if emp.get('is_expat'):
                        _gbp_sal_s = float(emp.get('gbp_salary') or 0)
                        result['gbp_data'] = {
                            'rate': _gbp_rate_saved,
                            'gbp_component': round(_gbp_sal_s * 0.75, 2),
                            'egp_component': round(_gbp_sal_s * _gbp_rate_saved * 0.25, 2) if _gbp_rate_saved > 0 else 0,
                            'gbp_variable_allowances':    saved_inputs.get('gbp_variable_allowances', 0),
                            'gbp_overtime':               saved_inputs.get('gbp_overtime', 0),
                            'gbp_bonuses':                saved_inputs.get('gbp_bonuses', 0),
                            'gbp_commissions':            saved_inputs.get('gbp_commissions', 0),
                            'gbp_flight_allowance':       saved_inputs.get('gbp_flight_allowance', 0),
                            'gbp_reimbursement':          saved_inputs.get('gbp_reimbursement', 0),
                            'gbp_reallocation_allowance': saved_inputs.get('gbp_reallocation_allowance', 0),
                            'gbp_other_earnings':         saved_inputs.get('gbp_other_earnings', 0),
                            'gbp_loan_deduction':         saved_inputs.get('gbp_loan_deduction', 0),
                            'gbp_other_deductions':       saved_inputs.get('gbp_other_deductions', 0),
                            'gbp_absent_days':            saved_inputs.get('gbp_absent_days', 0),
                            'gbp_late_minutes':           saved_inputs.get('gbp_late_minutes', 0),
                            'gbp_sick_days':              saved_inputs.get('gbp_sick_days', 0),
                            'gbp_sanction_days':          saved_inputs.get('gbp_sanction_days', 0),
                        }
                        # Compute GBP deduction amounts from stored days/mins
                        _gbp_wdays_g = float((emp.get('settings') or {}).get('working_days', 22) or 22)
                        _gbp_dr_g    = round(result['gbp_data']['gbp_component'] / _gbp_wdays_g, 4) if _gbp_wdays_g else 0
                        _gbp_wmins_g = float((emp.get('settings') or {}).get('work_minutes', 480) or 480)
                        result['gbp_data']['gbp_absent_amount']   = round(result['gbp_data'].get('gbp_absent_days', 0)   * _gbp_dr_g, 2)
                        result['gbp_data']['gbp_late_amount']     = round(result['gbp_data'].get('gbp_late_minutes', 0)  * (_gbp_dr_g / _gbp_wmins_g), 2) if _gbp_wmins_g else 0
                        result['gbp_data']['gbp_sick_amount']     = round(result['gbp_data'].get('gbp_sick_days', 0)     * _gbp_dr_g, 2)
                        result['gbp_data']['gbp_sanction_amount'] = round(result['gbp_data'].get('gbp_sanction_days', 0) * _gbp_dr_g, 2)

    if request.method == 'POST' and request.form.get('action') == 'calculate':
        eid = request.form.get('eid')
        if eid:
            emp = _enrich_emp_expat(load_emp_basic(int(eid)))
            if emp:
                sel_emp = emp
                sanction_days = float(request.form.get('sanction_days') or 0)
                other_amt = request.form.get('other_amount')
                other_egp = float(other_amt) if other_amt not in (None, '') else None
                monthly_earn = dict(emp['earnings'])
                monthly_earn.update({
                    'VariableAllowances':    float(request.form.get('variable_allowances') or 0),
                    'Overtime':              float(request.form.get('overtime') or 0),
                    'Bonuses':               float(request.form.get('bonuses') or 0),
                    'Commissions':           float(request.form.get('commissions') or 0),
                    'Reimbursement':         float(request.form.get('reimbursement') or 0),
                    'ReallocationAllowance': float(request.form.get('reallocation_allowance') or 0),
                    'FlightAllowance':       float(request.form.get('flight_allowance') or 0),
                })
                # Issue 6: Expat — build gbp_data for ALL expat employees (rate may be 0)
                _gbp_rate = float(request.form.get('gbp_rate') or 0)
                _prev_gbp_data = None
                if emp.get('is_expat'):
                    _gbp_sal = float(emp.get('gbp_salary') or 0)
                    _gbp_comp = round(_gbp_sal * 0.75, 2)
                    _egp_comp = round(_gbp_sal * _gbp_rate * 0.25, 2) if _gbp_rate > 0 else 0
                    # Only override BasicSalary when a rate is provided
                    if _gbp_rate > 0:
                        monthly_earn['BasicSalary'] = _egp_comp
                    _prev_gbp_data = {
                        'rate': _gbp_rate, 'gbp_component': _gbp_comp, 'egp_component': _egp_comp,
                        'gbp_variable_allowances':    float(request.form.get('gbp_variable_allowances') or 0),
                        'gbp_overtime':               float(request.form.get('gbp_overtime') or 0),
                        'gbp_bonuses':                float(request.form.get('gbp_bonuses') or 0),
                        'gbp_commissions':            float(request.form.get('gbp_commissions') or 0),
                        'gbp_flight_allowance':       float(request.form.get('gbp_flight_allowance') or 0),
                        'gbp_reimbursement':          float(request.form.get('gbp_reimbursement') or 0),
                        'gbp_reallocation_allowance': float(request.form.get('gbp_reallocation_allowance') or 0),
                        'gbp_other_earnings':         float(request.form.get('gbp_other_earnings') or 0),
                        'gbp_loan_deduction':         float(request.form.get('gbp_loan_deduction') or 0),
                        'gbp_other_deductions':       float(request.form.get('gbp_other_deductions') or 0),
                        'gbp_absent_days':            float(request.form.get('gbp_absent_days') or 0),
                        'gbp_late_minutes':           float(request.form.get('gbp_late_minutes') or 0),
                        'gbp_sick_days':              float(request.form.get('gbp_sick_days') or 0),
                        'gbp_sanction_days':          float(request.form.get('gbp_sanction_days') or 0),
                    }
                result = run_payroll(
                    monthly_earn,
                    absent_days=float(request.form.get('absent_days') or 0),
                    late_minutes=float(request.form.get('late_minutes') or 0),
                    sick_days=float(request.form.get('sick_days') or 0),
                    sanction_days=sanction_days,
                    loan=float(request.form.get('loan') or 0),
                    settings=emp['settings'],
                    other_amount=other_egp,
                )
                # Inject monthly earn overrides into result.inputs for template display
                result['inputs']['variable_allowances']    = monthly_earn.get('VariableAllowances', 0)
                result['inputs']['overtime']               = monthly_earn.get('Overtime', 0)
                result['inputs']['bonuses']                = monthly_earn.get('Bonuses', 0)
                result['inputs']['commissions']            = monthly_earn.get('Commissions', 0)
                result['inputs']['reimbursement']          = monthly_earn.get('Reimbursement', 0)
                result['inputs']['reallocation_allowance'] = monthly_earn.get('ReallocationAllowance', 0)
                result['inputs']['flight_allowance']       = monthly_earn.get('FlightAllowance', 0)
                if _prev_gbp_data:
                    # Compute GBP deduction amounts from days/mins inputs
                    _gbp_wdays = float((emp.get('settings') or {}).get('working_days', 22) or 22)
                    _gbp_dr    = round(_prev_gbp_data['gbp_component'] / _gbp_wdays, 4) if _gbp_wdays else 0
                    _gbp_work_mins = float((emp.get('settings') or {}).get('work_minutes', 480) or 480)
                    _prev_gbp_data['gbp_absent_amount']   = round(_prev_gbp_data.get('gbp_absent_days', 0)   * _gbp_dr, 2)
                    _prev_gbp_data['gbp_late_amount']     = round(_prev_gbp_data.get('gbp_late_minutes', 0)  * (_gbp_dr / _gbp_work_mins), 2) if _gbp_work_mins else 0
                    _prev_gbp_data['gbp_sick_amount']     = round(_prev_gbp_data.get('gbp_sick_days', 0)     * _gbp_dr, 2)
                    _prev_gbp_data['gbp_sanction_amount'] = round(_prev_gbp_data.get('gbp_sanction_days', 0) * _gbp_dr, 2)
                    result['gbp_data'] = _prev_gbp_data

    if request.method == 'POST' and request.form.get('action') in ('save', 'save_next'):
        is_save_next = request.form.get('action') == 'save_next'
        eid = request.form.get('eid')
        if eid and valid_month(month):
            emp = _enrich_emp_expat(load_emp_basic(int(eid)))
            if emp:
                sel_emp = emp
                sanction_days = float(request.form.get('sanction_days') or 0)
                other_amt = request.form.get('other_amount')
                other_egp = float(other_amt) if other_amt not in (None, '') else None
                monthly_earn = dict(emp['earnings'])
                earn_overrides = {
                    'VariableAllowances':    float(request.form.get('variable_allowances') or 0),
                    'Overtime':              float(request.form.get('overtime') or 0),
                    'Bonuses':               float(request.form.get('bonuses') or 0),
                    'Commissions':           float(request.form.get('commissions') or 0),
                    'Reimbursement':         float(request.form.get('reimbursement') or 0),
                    'ReallocationAllowance': float(request.form.get('reallocation_allowance') or 0),
                    'FlightAllowance':       float(request.form.get('flight_allowance') or 0),
                }
                monthly_earn.update(earn_overrides)
                # Issue 6: Expat salary split — build gbp_data for ALL expat employees
                gbp_data = None
                gbp_rate = float(request.form.get('gbp_rate') or 0)
                if emp.get('is_expat'):
                    gbp_salary = float(emp.get('gbp_salary') or 0)
                    gbp_component = round(gbp_salary * 0.75, 2)
                    egp_component = round(gbp_salary * gbp_rate * 0.25, 2) if gbp_rate > 0 else 0
                    # Only override BasicSalary when a rate is provided
                    if gbp_rate > 0:
                        monthly_earn['BasicSalary'] = egp_component
                    gbp_data = {
                        'rate': gbp_rate,
                        'gbp_component': gbp_component,
                        'egp_component': egp_component,
                        'gbp_variable_allowances':    float(request.form.get('gbp_variable_allowances') or 0),
                        'gbp_overtime':               float(request.form.get('gbp_overtime') or 0),
                        'gbp_bonuses':                float(request.form.get('gbp_bonuses') or 0),
                        'gbp_commissions':            float(request.form.get('gbp_commissions') or 0),
                        'gbp_flight_allowance':       float(request.form.get('gbp_flight_allowance') or 0),
                        'gbp_reimbursement':          float(request.form.get('gbp_reimbursement') or 0),
                        'gbp_reallocation_allowance': float(request.form.get('gbp_reallocation_allowance') or 0),
                        'gbp_other_earnings':         float(request.form.get('gbp_other_earnings') or 0),
                        'gbp_loan_deduction':         float(request.form.get('gbp_loan_deduction') or 0),
                        'gbp_other_deductions':       float(request.form.get('gbp_other_deductions') or 0),
                        'gbp_absent_days':            float(request.form.get('gbp_absent_days') or 0),
                        'gbp_late_minutes':           float(request.form.get('gbp_late_minutes') or 0),
                        'gbp_sick_days':              float(request.form.get('gbp_sick_days') or 0),
                        'gbp_sanction_days':          float(request.form.get('gbp_sanction_days') or 0),
                    }
                    # Compute GBP deduction amounts from days/mins inputs
                    _gbp_wdays_s = float((emp.get('settings') or {}).get('working_days', 22) or 22)
                    _gbp_dr_s    = round(gbp_data['gbp_component'] / _gbp_wdays_s, 4) if _gbp_wdays_s else 0
                    _gbp_wmins_s = float((emp.get('settings') or {}).get('work_minutes', 480) or 480)
                    gbp_data['gbp_absent_amount']   = round(gbp_data.get('gbp_absent_days', 0)   * _gbp_dr_s, 2)
                    gbp_data['gbp_late_amount']     = round(gbp_data.get('gbp_late_minutes', 0)  * (_gbp_dr_s / _gbp_wmins_s), 2) if _gbp_wmins_s else 0
                    gbp_data['gbp_sick_amount']     = round(gbp_data.get('gbp_sick_days', 0)     * _gbp_dr_s, 2)
                    gbp_data['gbp_sanction_amount'] = round(gbp_data.get('gbp_sanction_days', 0) * _gbp_dr_s, 2)
                r = run_payroll(
                    monthly_earn,
                    absent_days=float(request.form.get('absent_days') or 0),
                    late_minutes=float(request.form.get('late_minutes') or 0),
                    sick_days=float(request.form.get('sick_days') or 0),
                    sanction_days=sanction_days,
                    loan=float(request.form.get('loan') or 0),
                    settings=emp['settings'],
                    other_amount=other_egp,
                )
                if gbp_data:
                    r['gbp_data'] = gbp_data
                _upsert_payroll_run(eid, month, r, emp,
                                    notes=request.form.get('notes', ''),
                                    earn_overrides=earn_overrides,
                                    gbp_data=gbp_data,
                                    is_manual=True)
                flash(f'Payroll saved for {emp["name"]} — {month}.', 'success')
                # Save & Next: find next employee without a saved run and redirect
                if is_save_next:
                    all_eids = [e['eid'] for e in emp_list]
                    cur_idx = all_eids.index(int(eid)) if int(eid) in all_eids else -1
                    # Update saved_eids with this just-saved employee
                    saved_eids.add(int(eid))
                    next_eid = None
                    # First try employees after current, then wrap around
                    for candidate_eid in all_eids[cur_idx+1:] + all_eids[:cur_idx]:
                        if candidate_eid not in saved_eids:
                            next_eid = candidate_eid
                            break
                    if next_eid:
                        return redirect(url_for('payroll_manual', eid=next_eid, month=month))
                    else:
                        # All done — stay on page, flash a completion message
                        flash(f'All employees done for {month}! 🎉', 'success')
                        return redirect(url_for('payroll_manual', month=month))
                # Inject monthly earn overrides into result.inputs for template display
                r['inputs']['variable_allowances']    = earn_overrides.get('VariableAllowances', 0)
                r['inputs']['overtime']               = earn_overrides.get('Overtime', 0)
                r['inputs']['bonuses']                = earn_overrides.get('Bonuses', 0)
                r['inputs']['commissions']            = earn_overrides.get('Commissions', 0)
                r['inputs']['reimbursement']          = earn_overrides.get('Reimbursement', 0)
                r['inputs']['reallocation_allowance'] = earn_overrides.get('ReallocationAllowance', 0)
                r['inputs']['flight_allowance']       = earn_overrides.get('FlightAllowance', 0)
                result = r

    # Load deduction history for selected employee (last 12 months)
    history = []
    if sel_emp:
        history = _load_history(sel_emp['eid'])

    done_count = sum(1 for e in emp_list if e['done'])
    # calc_gbp_rate: rate entered during a Calculate POST — carry it back so the
    # input field re-populates after the page reload (before the entry is saved to DB)
    calc_gbp_rate = 0.0
    if request.method == 'POST' and request.form.get('action') == 'calculate':
        try:
            calc_gbp_rate = float(request.form.get('gbp_rate') or 0)
        except (ValueError, TypeError):
            calc_gbp_rate = 0.0
    return render_template('payroll_manual.html',
                           emp_list=emp_list, result=result, sel_emp=sel_emp,
                           month=month, history=history, saved_inputs=saved_inputs,
                           done_count=done_count, total_count=len(emp_list),
                           default_gbp_rate=default_gbp_rate,
                           calc_gbp_rate=calc_gbp_rate)

def _enrich_emp_expat(emp):
    """Overlay EmployeeType, GBPSalary and Nationality from DB onto an emp dict (from load_emp_basic)."""
    if not emp:
        return emp
    _has_nat  = has_col('Employees', 'Nationality')
    _has_type = has_col('Employees', 'EmployeeType')
    _has_gbp  = has_col('Employees', 'GBPSalary')
    _sel = []
    if _has_nat:
        _sel.append("ISNULL(Nationality,'Egyptian')")
    if _has_type:
        _sel.append("ISNULL(EmployeeType,'Egyptian')")
    if _has_gbp:
        _sel.append("ISNULL(GBPSalary,0)")
    if _sel:
        rows = qy(f"SELECT {','.join(_sel)} FROM Employees WHERE EmpID=?", (emp['eid'],))
    else:
        rows = []
    if rows:
        vals = rows[0]
        idx = 0
        if _has_nat:
            emp['nationality'] = vals[idx] or 'Egyptian'
            idx += 1
        else:
            emp.setdefault('nationality', 'Egyptian')
        if _has_type:
            emp['employee_type'] = vals[idx] or 'Egyptian'
            idx += 1
        else:
            emp.setdefault('employee_type', 'Egyptian')
        if _has_gbp:
            emp['gbp_salary'] = float(vals[idx] or 0)
        else:
            emp.setdefault('gbp_salary', 0.0)
    else:
        emp.setdefault('nationality', 'Egyptian')
        emp.setdefault('employee_type', 'Egyptian')
        emp.setdefault('gbp_salary', 0.0)
    # Determine is_expat by nationality: non-Egyptian = Expat logic
    emp['is_expat'] = (emp.get('nationality', 'Egyptian') or 'Egyptian').strip().lower() != 'egyptian'
    return emp

def _upsert_payroll_run(eid, month, r, emp, notes='', earn_overrides=None, gbp_data=None, is_manual=False):
    d = r['ded']; t = r['tax']
    day_rate = d.get('day_rate', 0)
    eo = earn_overrides or {}
    gd = gbp_data or {}  # {'rate': float, 'gbp_component': float, 'egp_component': float}
    cols_list = ['EmpID','PayrollMonth','MonthlyGross','SocialInsurance','IncomeTax','Loan',
                 'Absences','Lateness','SickLeave','Sanctions','OtherDeductions',
                 'TotalDeductions','NetSalary','Notes','RunDate',
                 'AbsentDays','LateMinutes','SickDays','SanctionDays','OtherDays','DayRate','UpdatedBy',
                 'VariableAllowances','Overtime','Bonuses','Commissions','Reimbursement','ReallocationAllowance',
                 'FlightAllowance']
    vals_list = [
        eid, month, r['gross'], t['si_monthly'], t['monthly_tax'], r['loan'],
        d['absences'], d['lateness'], d['sick'], d['sanctions'], d['other'],
        r['total_ded'], r['net'], notes, date.today().isoformat(),
        r['inputs']['absent_days'], r['inputs']['late_minutes'],
        r['inputs']['sick_days'], r['inputs']['sanction_days'],
        r['inputs']['other_days'], day_rate,
        session.get('user', 'web'),
        eo.get('VariableAllowances', 0), eo.get('Overtime', 0),
        eo.get('Bonuses', 0), eo.get('Commissions', 0),
        eo.get('Reimbursement', 0), eo.get('ReallocationAllowance', 0),
        eo.get('FlightAllowance', 0),
    ]
    # Request 4: IsManualEntry flag
    if has_col('PayrollRuns', 'IsManualEntry'):
        cols_list.append('IsManualEntry')
        vals_list.append(1 if is_manual else 0)
    # Issue 6: Add GBP salary split columns if they exist in DB
    if gd and has_col('PayrollRuns', 'GBPEGPRate'):
        cols_list += ['GBPEGPRate', 'GBPComponent', 'EGPComponent']
        vals_list += [gd.get('rate', 0), gd.get('gbp_component', 0), gd.get('egp_component', 0)]
    # Issues 1 & 2: GBP monthly earn/ded columns
    if has_col('PayrollRuns', 'GBPBonuses'):
        cols_list += ['GBPVariableAllowances', 'GBPOvertime', 'GBPBonuses', 'GBPCommissions',
                      'GBPFlightAllowance', 'GBPReimbursement', 'GBPReallocationAllowance',
                      'GBPOtherEarnings', 'GBPLoanDeduction', 'GBPOtherDeductions']
        vals_list += [
            gd.get('gbp_variable_allowances', 0),
            gd.get('gbp_overtime', 0),
            gd.get('gbp_bonuses', 0),
            gd.get('gbp_commissions', 0),
            gd.get('gbp_flight_allowance', 0),
            gd.get('gbp_reimbursement', 0),
            gd.get('gbp_reallocation_allowance', 0),
            gd.get('gbp_other_earnings', 0),
            gd.get('gbp_loan_deduction', 0),
            gd.get('gbp_other_deductions', 0),
        ]
        if has_col('PayrollRuns', 'GBPAbsentDays'):
            cols_list += ['GBPAbsentDays', 'GBPLateMinutes', 'GBPSickDays', 'GBPSanctionDays']
            vals_list += [
                gd.get('gbp_absent_days', 0),
                gd.get('gbp_late_minutes', 0),
                gd.get('gbp_sick_days', 0),
                gd.get('gbp_sanction_days', 0),
            ]
    cols = ','.join(cols_list)
    vals = tuple(vals_list)
    existing = qy("SELECT 1 FROM PayrollRuns WHERE EmpID=? AND PayrollMonth=?", (eid, month))
    if existing:
        set_clause = ','.join(f"{c}=?" for c in cols_list[2:])
        ex(f"UPDATE PayrollRuns SET {set_clause} WHERE EmpID=? AND PayrollMonth=?",
           list(vals[2:]) + [eid, month])
    else:
        ex(f"INSERT INTO PayrollRuns ({cols}) VALUES ({','.join('?'*len(vals))})", vals)

def _run_all_payroll(month, overwrite=False):
    """Run payroll for all active employees for the given month.
    Returns list of result dicts. Skips existing unless overwrite=True."""
    s = get_payroll_settings(int(month[:4]))
    emps = qy(f"""SELECT e.EmpID, ISNULL(e.FirstName,'')+' '+ISNULL(e.FamilyName,''),
                         ISNULL(e.Department,''), ISNULL(e.PersonalEmail,'')
                  FROM Employees e WHERE {ACTIVE_EMP_SQL}""")
    results = []
    for eid, name, dept, email in emps:
        # Request 4: Always skip manually-entered payroll (even when overwrite=True)
        if has_col('PayrollRuns', 'IsManualEntry'):
            manual_row = qy("SELECT IsManualEntry FROM PayrollRuns WHERE EmpID=? AND PayrollMonth=?", (eid, month))
            if manual_row and manual_row[0][0]:
                row = qy("""SELECT MonthlyGross,SocialInsurance,IncomeTax,NetSalary
                            FROM PayrollRuns WHERE EmpID=? AND PayrollMonth=?""", (eid, month))
                if row:
                    results.append({'name': name, 'dept': dept, 'email': email,
                                    'gross': float(row[0][0] or 0), 'si': float(row[0][1] or 0),
                                    'tax': float(row[0][2] or 0), 'net': float(row[0][3] or 0),
                                    'eid': eid})
                continue
        # Skip if already processed and overwrite not set
        if not overwrite:
            existing = qy("SELECT 1 FROM PayrollRuns WHERE EmpID=? AND PayrollMonth=?", (eid, month))
            if existing:
                # Load existing for the results display
                row = qy("""SELECT MonthlyGross,SocialInsurance,IncomeTax,NetSalary
                            FROM PayrollRuns WHERE EmpID=? AND PayrollMonth=?""", (eid, month))
                if row:
                    results.append({'name': name, 'dept': dept, 'email': email,
                                    'gross': float(row[0][0] or 0), 'si': float(row[0][1] or 0),
                                    'tax': float(row[0][2] or 0), 'net': float(row[0][3] or 0),
                                    'eid': eid})
                continue
        emp = _enrich_emp_expat(load_emp_basic(eid))
        if not emp:
            continue
        absent = late = sick = sanction = other = loan = 0
        att = qy("""SELECT AbsentDays, LateMinutes, SickDays, SanctionDays, OtherDays
                    FROM AttendanceRecords WHERE EmpID=? AND PayrollMonth=?""", (eid, month))
        if att:
            absent, late, sick, sanction, other = [float(x or 0) for x in att[0]]
        loan_row = qy("SELECT SUM(ISNULL(Amount,0)) FROM LoanDeductions WHERE EmpID=? AND PayrollMonth=?",
                      (eid, month))
        if loan_row:
            loan = float(loan_row[0][0] or 0)
        r = run_payroll(emp['earnings'], absent, late, sick, sanction, other, loan, s)
        # Build GBP data for expat employees
        gbp_data_bulk = None
        if emp.get('is_expat') and emp.get('gbp_salary', 0):
            # Use existing rate from this month if available, else 0 (HR must correct via manual entry)
            rate_row = qy("SELECT TOP 1 GBPEGPRate FROM PayrollRuns WHERE PayrollMonth=? AND GBPEGPRate IS NOT NULL AND GBPEGPRate > 0", (month,))
            rate = float(rate_row[0][0]) if rate_row else 0.0
            gbp_sal = float(emp.get('gbp_salary') or 0)
            gbp_component = round(gbp_sal * 0.75, 2)
            egp_component = round(gbp_sal * rate * 0.25, 2) if rate else 0.0
            gbp_data_bulk = {'rate': rate, 'gbp_component': gbp_component, 'egp_component': egp_component}
        _upsert_payroll_run(eid, month, r, emp, gbp_data=gbp_data_bulk)
        results.append({'name': name, 'dept': dept, 'email': email,
                        'gross': r['gross'], 'net': r['net'],
                        'tax': r['tax']['monthly_tax'], 'si': r['tax']['si_monthly'],
                        'eid': eid})
    return results

@app.route('/payroll/run-all', methods=['GET', 'POST'])
@login_required
@role_required('Admin', 'HR')
def payroll_run_all():
    today = date.today()
    cur_month  = today.strftime('%Y-%m')
    next_month = (today.replace(day=1) + timedelta(days=32)).replace(day=1).strftime('%Y-%m')
    month    = request.form.get('month') or cur_month
    overwrite = bool(request.form.get('overwrite'))
    action   = request.form.get('action', 'run')
    results  = None
    email_results = False

    if request.method == 'POST':
        if month < cur_month or month > next_month:
            flash(f'Payroll can only be run for {cur_month} (current) or {next_month} (next month).', 'error')
        elif action == 'run':
            results = _run_all_payroll(month, overwrite)
            flash(f'Payroll run complete: {len(results)} employees processed for {month}.', 'success')
        elif action == 'excel':
            # Build bulk Excel with one sheet per employee
            try:
                import openpyxl
                from openpyxl.styles import Font, PatternFill, Alignment
                wb = openpyxl.Workbook()
                wb.remove(wb.active)  # remove default sheet
                rows_all = _run_all_payroll(month, overwrite=False)
                for rec in rows_all:
                    eid  = rec['eid']
                    emp  = load_emp_basic(eid)
                    if not emp:
                        continue
                    emp = _enrich_emp_expat(emp)
                    _gbp_sel = (", ISNULL(GBPEGPRate,0), ISNULL(GBPComponent,0), ISNULL(EGPComponent,0)"
                                if has_col('PayrollRuns','GBPEGPRate') else ", 0, 0, 0")
                    row_db = qy(f"""SELECT MonthlyGross,SocialInsurance,IncomeTax,Loan,
                                          Absences,Lateness,SickLeave,Sanctions,OtherDeductions,
                                          TotalDeductions,NetSalary,DayRate,Notes{_gbp_sel}
                                   FROM PayrollRuns WHERE EmpID=? AND PayrollMonth=?""", (eid, month))
                    if not row_db:
                        continue
                    r0 = row_db[0]
                    gross=float(r0[0] or 0); si=float(r0[1] or 0); tax=float(r0[2] or 0)
                    loan=float(r0[3] or 0)
                    abs_=float(r0[4] or 0); lat=float(r0[5] or 0); sick=float(r0[6] or 0)
                    sanc=float(r0[7] or 0); oth=float(r0[8] or 0)
                    total_ded=float(r0[9] or 0); net=float(r0[10] or 0); notes=r0[12] or ''
                    gbp_rate=float(r0[13] or 0); gbp_comp=float(r0[14] or 0); egp_comp=float(r0[15] or 0)

                    ws = wb.create_sheet(title=(emp['name'] or str(eid))[:31])
                    ws.column_dimensions['A'].width = 30
                    ws.column_dimensions['B'].width = 18

                    def hdr(rn, txt, col='143256'):
                        c = ws.cell(row=rn, column=1, value=txt)
                        c.font = Font(bold=True, color='FFFFFF', size=10)
                        c.fill = PatternFill('solid', fgColor=col)
                        ws.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=2)

                    def row2(rn, lbl, val, bold=False, prefix=''):
                        ws.cell(row=rn, column=1, value=lbl)
                        display_val = round(float(val or 0), 2)
                        c2 = ws.cell(row=rn, column=2,
                                     value=(prefix + f'{display_val:,.2f}') if prefix else display_val)
                        if not prefix:
                            c2.number_format = '#,##0.00'
                        c2.alignment = Alignment(horizontal='right')
                        if bold:
                            ws.cell(row=rn, column=1).font = Font(bold=True)
                            c2.font = Font(bold=True)

                    rn = 1
                    ws.cell(row=rn, column=1, value='KINGS SCHOOL EGYPT — PAYSLIP').font = Font(bold=True, size=13)
                    ws.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=2); rn+=1
                    ws.cell(row=rn, column=1,
                            value=f"Employee: {emp['name']}  |  Month: {month}  |  "
                                  f"Nationality: {emp.get('nationality','') or 'Egyptian'}"); rn+=2

                    if emp.get('is_expat') and gbp_comp:
                        # GBP section
                        hdr(rn, 'GBP EARNINGS (EXPAT)'); rn+=1
                        if gbp_rate: row2(rn, '  GBP/EGP Rate', gbp_rate, prefix=''); rn+=1
                        row2(rn, '  Base GBP Pay (75%)', gbp_comp, prefix='£'); rn+=1
                        row2(rn, 'Total GBP Earnings', gbp_comp, bold=True, prefix='£'); rn+=2
                        hdr(rn, 'GBP NET PAY', col='059669'); rn+=1
                        row2(rn, 'GBP Net Salary', gbp_comp, bold=True, prefix='£'); rn+=2
                        # EGP section (25% portion)
                        hdr(rn, 'EGP EARNINGS (25% PORTION)'); rn+=1
                        for col_k, lbl in [('BasicSalary','Basic Salary'),('HousingAllowance','Housing'),
                                           ('TransportAllowance','Transport'),('OtherFixedAllowances','Other Fixed'),
                                           ('VariableAllowances','Variable'),('Overtime','Overtime'),
                                           ('Bonuses','Bonuses'),('Commissions','Commissions')]:
                            v = float(emp['earnings'].get(col_k) or 0)
                            if v: row2(rn, '  '+lbl, v, prefix='EGP '); rn+=1
                        row2(rn, 'Total EGP Gross', gross, bold=True, prefix='EGP '); rn+=2
                        hdr(rn, 'EGP DEDUCTIONS', col='DC001E'); rn+=1
                        for lbl, val in [('  SI',si),('  Tax',tax),('  Loan',loan),
                                         ('  Absences',abs_),('  Lateness',lat),
                                         ('  Sick',sick),('  Sanctions',sanc),('  Other',oth)]:
                            if val: row2(rn, lbl, val, prefix='EGP '); rn+=1
                        row2(rn, 'Total EGP Deductions', total_ded, bold=True, prefix='EGP '); rn+=2
                        hdr(rn, 'EGP NET (25%)', col='16a34a'); rn+=1
                        row2(rn, 'EGP Net Salary', net, bold=True, prefix='EGP ')
                    else:
                        hdr(rn, 'GROSS EARNINGS (EGP)'); rn+=1
                        for col_k, lbl in [('BasicSalary','Basic Salary'),('HousingAllowance','Housing'),
                                           ('TransportAllowance','Transport'),('OtherFixedAllowances','Other Fixed'),
                                           ('VariableAllowances','Variable'),('Overtime','Overtime'),
                                           ('Bonuses','Bonuses'),('Commissions','Commissions')]:
                            v = float(emp['earnings'].get(col_k) or 0)
                            if v: row2(rn, '  '+lbl, v); rn+=1
                        row2(rn, 'Total Gross (EGP)', gross, bold=True); rn+=2
                        hdr(rn, 'DEDUCTIONS (EGP)', col='DC001E'); rn+=1
                        for lbl, val in [('  Social Insurance', si),('  Income Tax', tax),
                                         ('  Loan', loan),('  Absences', abs_),('  Lateness', lat),
                                         ('  Sick Leave', sick),('  Sanctions', sanc),('  Other', oth)]:
                            if val: row2(rn, lbl, val); rn+=1
                        row2(rn, 'Total Deductions', total_ded, bold=True); rn+=2
                        hdr(rn, 'NET SALARY (EGP)', col='16a34a'); rn+=1
                        row2(rn, 'Net Salary', net, bold=True)

                buf = io.BytesIO()
                wb.save(buf); buf.seek(0)
                return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                 as_attachment=True, download_name=f'Payroll_{month}_All.xlsx')
            except Exception as e:
                flash(f'Excel export error: {e}', 'error')
        elif action == 'email':
            return redirect(url_for('email_payslips', month=month))

    return render_template('payroll_run_all.html',
                           cur_month=cur_month, next_month=next_month,
                           month=month, overwrite=overwrite,
                           results=results, email_results=email_results)

def _html_to_pdf(html_content):
    """Convert HTML string to PDF bytes using wkhtmltopdf. Returns (bytes, 'pdf') or (html_bytes, 'html')."""
    import subprocess, tempfile, os as _os, shutil
    # Find wkhtmltopdf — check PATH first, then common Windows install location
    exe = shutil.which('wkhtmltopdf')
    if not exe:
        win_path = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
        if _os.path.exists(win_path):
            exe = win_path
    if not exe:
        return html_content.encode('utf-8'), 'html'
    tmp_html = tmp_pdf = None
    try:
        with tempfile.NamedTemporaryFile(suffix='.html', delete=False, mode='w', encoding='utf-8') as tf:
            tf.write(html_content); tmp_html = tf.name
        tmp_pdf = tmp_html.replace('.html', '.pdf')
        result = subprocess.run(
            [exe, '--quiet', '--encoding', 'UTF-8', '--page-size', 'A4', tmp_html, tmp_pdf],
            capture_output=True, timeout=30)
        if result.returncode == 0 and _os.path.exists(tmp_pdf):
            with open(tmp_pdf, 'rb') as pf:
                return pf.read(), 'pdf'
    except Exception:
        pass
    finally:
        for _f in [tmp_html, tmp_pdf]:
            if _f:
                try: _os.unlink(_f)
                except Exception: pass
    return html_content.encode('utf-8'), 'html'

def _get_smtp_config():
    """Read SMTP config from DB (SmtpConfig table). Returns dict."""
    try:
        rows = qy("SELECT ConfigKey, ConfigValue FROM SmtpConfig")
        cfg = {r[0]: r[1] for r in rows} if rows else {}
    except Exception:
        cfg = {}
    return cfg

def _save_smtp_config(cfg):
    """Upsert key/value pairs into SmtpConfig table."""
    try:
        ex("""IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME='SmtpConfig')
              CREATE TABLE SmtpConfig (ConfigKey NVARCHAR(100) PRIMARY KEY, ConfigValue NVARCHAR(500))""")
    except Exception:
        pass
    for k, v in cfg.items():
        try:
            ex("""MERGE SmtpConfig AS t USING (VALUES(?,?)) AS s(k,v) ON t.ConfigKey=s.k
                  WHEN MATCHED THEN UPDATE SET ConfigValue=s.v
                  WHEN NOT MATCHED THEN INSERT (ConfigKey,ConfigValue) VALUES(s.k,s.v)""", (k, v))
        except Exception:
            pass

@app.route('/payroll/email-payslips', methods=['GET', 'POST'])
@login_required
@role_required('Admin', 'HR')
def email_payslips():
    today = date.today()
    cur_month  = today.strftime('%Y-%m')
    month = request.args.get('month') or request.form.get('month') or cur_month
    log_lines = []
    cfg = _get_smtp_config()

    if request.method == 'POST':
        # Save config from form
        new_cfg = {
            'smtp_host': request.form.get('smtp_host', '').strip(),
            'smtp_port': request.form.get('smtp_port', '587').strip(),
            'smtp_user': request.form.get('smtp_user', '').strip(),
            'smtp_from': request.form.get('smtp_from', '').strip(),
            'smtp_subject': request.form.get('smtp_subject', 'Your Payslip — Kings School Egypt').strip(),
        }
        # Only update password if provided (don't blank it out accidentally)
        new_pass = request.form.get('smtp_pass', '').strip()
        if new_pass:
            new_cfg['smtp_pass'] = new_pass
        else:
            new_cfg['smtp_pass'] = cfg.get('smtp_pass', '')
        _save_smtp_config(new_cfg)
        cfg = new_cfg

        smtp_host = cfg.get('smtp_host', '')
        smtp_port = int(cfg.get('smtp_port', 587) or 587)
        smtp_user = cfg.get('smtp_user', '')
        smtp_pass = cfg.get('smtp_pass', '')
        smtp_from = cfg.get('smtp_from', '') or smtp_user
        smtp_subject = cfg.get('smtp_subject', 'Your Payslip — Kings School Egypt')

        if not smtp_host or not smtp_user or not smtp_pass:
            flash('Please fill in all email configuration fields (host, user, password).', 'error')
        else:
            import smtplib
            from email.mime.multipart import MIMEMultipart
            from email.mime.text import MIMEText
            from email.mime.application import MIMEApplication

            rows_all = _run_all_payroll(month, overwrite=False)
            server = None
            try:
                server = smtplib.SMTP(smtp_host, smtp_port, timeout=15)
                server.ehlo(); server.starttls(); server.login(smtp_user, smtp_pass)
                log_lines.append(('ok', f'Connected to {smtp_host}:{smtp_port}'))
            except Exception as e:
                log_lines.append(('err', f'Connection failed: {e}'))
                flash(f'Could not connect to email server: {e}', 'error')

            sent = skipped = failed = 0
            for rec in rows_all:
                eid     = rec['eid']
                to_addr = rec.get('email', '').strip()
                emp     = load_emp_basic(eid)
                if not emp:
                    continue
                _fa_sel = ", ISNULL(FlightAllowance,0)" if has_col('PayrollRuns','FlightAllowance') else ", 0"
                row_db = qy(f"""SELECT MonthlyGross,SocialInsurance,IncomeTax,Loan,
                                      Absences,Lateness,SickLeave,Sanctions,OtherDeductions,
                                      TotalDeductions,NetSalary,DayRate,Notes,
                                      ISNULL(VariableAllowances,0),ISNULL(Overtime,0),
                                      ISNULL(Bonuses,0),ISNULL(Commissions,0),
                                      ISNULL(Reimbursement,0),ISNULL(ReallocationAllowance,0)
                                      {_fa_sel}
                               FROM PayrollRuns WHERE EmpID=? AND PayrollMonth=?""", (eid, month))
                if not row_db:
                    log_lines.append(('skip', f'{emp["name"]} — no payroll entry for {month}'))
                    skipped += 1
                    continue
                r0 = row_db[0]
                gross=float(r0[0] or 0); si=float(r0[1] or 0); tax=float(r0[2] or 0)
                loan=float(r0[3] or 0); abs_=float(r0[4] or 0); lat=float(r0[5] or 0)
                sick=float(r0[6] or 0); sanc=float(r0[7] or 0); oth=float(r0[8] or 0)
                total_ded=float(r0[9] or 0); net=float(r0[10] or 0)
                day_rate=float(r0[11] or 0); notes=r0[12] or ''
                # Overlay monthly variable earnings onto emp['earnings'] for payslip display
                emp['earnings'].update({
                    'VariableAllowances':    float(r0[13] or 0),
                    'Overtime':              float(r0[14] or 0),
                    'Bonuses':               float(r0[15] or 0),
                    'Commissions':           float(r0[16] or 0),
                    'Reimbursement':         float(r0[17] or 0),
                    'ReallocationAllowance': float(r0[18] or 0),
                    'FlightAllowance':       float(r0[19] or 0) if len(r0) > 19 else 0,
                })
                r_full = run_payroll(emp['earnings'], settings=emp['settings'])
                tax_detail = r_full['tax']

                if not to_addr:
                    log_lines.append(('skip', f'{emp["name"]} — no email address on file'))
                    skipped += 1
                    continue
                if not server:
                    log_lines.append(('err', f'{emp["name"]} — skipped (no connection)'))
                    failed += 1
                    continue
                try:
                    today_str = date.today().strftime('%d %b %Y')
                    html_content = render_template('payslip_pdf.html',
                        emp=emp, month=month, gross=gross, si=si, tax=tax, loan=loan,
                        absences=abs_, lateness=lat, sick=sick, sanctions=sanc,
                        other_ded=oth, total_ded=total_ded, net=net, day_rate=day_rate,
                        notes=notes, tax_detail=tax_detail, today_str=today_str)
                    att_data, fmt = _html_to_pdf(html_content)
                    att_mime     = 'application/pdf' if fmt == 'pdf' else 'text/html'
                    att_filename = f'Payslip_{month}_{emp.get("emp_id","")}.{fmt}'
                    msg = MIMEMultipart('mixed')
                    msg['From']    = smtp_from
                    msg['To']      = to_addr
                    msg['Subject'] = f'{smtp_subject} — {month}'
                    body_html = (f"<p>Dear {emp['name']},</p>"
                                 f"<p>Please find your payslip for <strong>{month}</strong> attached.</p>"
                                 f"<p>Net Salary: <strong>EGP {net:,.2f}</strong></p>"
                                 "<p>Kings School Egypt — HR Department</p>")
                    msg.attach(MIMEText(body_html, 'html'))
                    att = MIMEApplication(att_data, _subtype=att_mime.split('/')[-1])
                    att.add_header('Content-Disposition', 'attachment', filename=att_filename)
                    msg.attach(att)
                    server.sendmail(smtp_from, to_addr, msg.as_string())
                    log_lines.append(('ok', f'{emp["name"]} → {to_addr}'))
                    sent += 1
                except Exception as e:
                    log_lines.append(('err', f'{emp["name"]} — {str(e)[:80]}'))
                    failed += 1

            if server:
                try: server.quit()
                except Exception: pass

            flash(f'Done: {sent} sent, {skipped} skipped, {failed} failed.', 'success' if sent else 'error')

    return render_template('email_payslips.html',
                           cfg=cfg, month=month, log_lines=log_lines)

@app.route('/payroll/history')
@login_required
@role_required('Admin', 'HR')
def payroll_history():
    months = qy("""SELECT PayrollMonth, COUNT(*) as EmpCount,
                   SUM(MonthlyGross) as TotalGross,
                   SUM(SocialInsurance) as TotalSI,
                   SUM(IncomeTax) as TotalTax,
                   SUM(Absences+Lateness+SickLeave+Sanctions+OtherDeductions) as TotalDed,
                   SUM(Loan) as TotalLoan,
                   SUM(TotalDeductions) as TotalAllDed,
                   SUM(NetSalary) as TotalNet,
                   MIN(RunDate) as RunDate
                   FROM PayrollRuns GROUP BY PayrollMonth ORDER BY PayrollMonth DESC""") or []
    months_data = [{
        'month': r[0], 'emps': r[1], 'gross': r[2], 'si': r[3], 'tax': r[4],
        'ded': r[5], 'loan': r[6], 'total_ded': r[7], 'net': r[8], 'run_date': r[9],
    } for r in months]
    return render_template('payroll_history.html', months=months_data)

@app.route('/payroll/history/<month>')
@login_required
@role_required('Admin', 'HR')
def payroll_history_detail(month):
    if not valid_month(month):
        flash('Invalid month.', 'error')
        return redirect(url_for('payroll_history'))
    _gbp_hist = (", ISNULL(p.GBPComponent,0), ISNULL(p.GBPEGPRate,0), ISNULL(p.IsManualEntry,0)"
                 if has_col('PayrollRuns','GBPEGPRate') else ", 0, 0, 0")
    rows = qy(f"""SELECT e.EmpID, ISNULL(e.EmployeeID,''),
                 ISNULL(e.FirstName,'')+' '+ISNULL(e.FamilyName,''),
                 ISNULL(e.Department,''), ISNULL(e.JobTitle,''),
                 ISNULL(e.Nationality,'Egyptian'),
                 p.MonthlyGross, p.SocialInsurance, p.IncomeTax, p.Loan,
                 p.Absences, p.Lateness, p.SickLeave, p.Sanctions, p.OtherDeductions,
                 p.TotalDeductions, p.NetSalary,
                 ISNULL(p.AbsentDays,0), ISNULL(p.LateMinutes,0),
                 ISNULL(p.SickDays,0), ISNULL(p.SanctionDays,0), ISNULL(p.OtherDays,0),
                 ISNULL(p.DayRate,0), ISNULL(p.Notes,''), ISNULL(p.RunDate,''), ISNULL(p.UpdatedBy,'')
                 {_gbp_hist}
                 FROM PayrollRuns p
                 JOIN Employees e ON e.EmpID=p.EmpID
                 WHERE p.PayrollMonth=? ORDER BY e.FirstName""", (month,))
    detail = [{
        'eid': r[0], 'emp_id': r[1], 'name': r[2], 'dept': r[3], 'job': r[4],
        'nationality': r[5],
        'gross': r[6], 'si': r[7], 'tax': r[8], 'loan': r[9],
        'absences': r[10], 'lateness': r[11], 'sick': r[12],
        'sanctions': r[13], 'other_ded': r[14], 'total_ded': r[15], 'net': r[16],
        'absent_days': r[17], 'late_minutes': r[18], 'sick_days': r[19],
        'sanction_days': r[20], 'other_days': r[21], 'day_rate': r[22],
        'notes': r[23], 'run_date': r[24], 'updated_by': r[25],
        'gbp_component': float(r[26] or 0), 'gbp_rate': float(r[27] or 0),
        'is_manual': bool(r[28]),
    } for r in rows]
    totals = {
        'gross': sum(d['gross'] or 0 for d in detail),
        'si': sum(d['si'] or 0 for d in detail),
        'tax': sum(d['tax'] or 0 for d in detail),
        'loan': sum(d['loan'] or 0 for d in detail),
        'total_ded': sum(d['total_ded'] or 0 for d in detail),
        'net': sum(d['net'] or 0 for d in detail),
    }
    return render_template('payroll_history_detail.html',
                           month=month, detail=detail, totals=totals)

@app.route('/payroll/history/export/<fmt>')
@login_required
@role_required('Admin', 'HR')
def payroll_history_export(fmt):
    """Export payroll history summary as Excel or PDF."""
    months = qy("""SELECT PayrollMonth, COUNT(*), SUM(MonthlyGross), SUM(SocialInsurance),
                   SUM(IncomeTax), SUM(Absences+Lateness+SickLeave+Sanctions+OtherDeductions),
                   SUM(Loan), SUM(TotalDeductions), SUM(NetSalary), MIN(RunDate)
                   FROM PayrollRuns GROUP BY PayrollMonth ORDER BY PayrollMonth DESC""") or []
    headers = ['Month','Employees','Gross Salary','SI','Income Tax',
               'Day Deductions','Loan','Total Deductions','Net Salary','Run Date']
    rows_data = [list(r) for r in months]

    if fmt == 'xlsx':
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = Workbook()
        ws = wb.active
        ws.title = 'Payroll History'
        hfill = PatternFill('solid', fgColor='4F46E5')
        hfont = Font(bold=True, color='FFFFFF')
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=ci, value=h)
            c.font = hfont; c.fill = hfill; c.alignment = Alignment(horizontal='center')
        num_cols = {2,3,4,5,6,7,8,9}
        for ri, row in enumerate(rows_data, 2):
            for ci, val in enumerate(row, 1):
                cell = ws.cell(row=ri, column=ci)
                if ci in num_cols and val is not None:
                    try: cell.value = float(val); cell.number_format = '#,##0.00'
                    except: cell.value = val
                elif ci == 10:  # date
                    cell.value = str(val)[:10] if val else ''
                else:
                    cell.value = val
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = max(14, len(str(col[0].value or ''))+4)
        buf = io.BytesIO()
        wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name='payroll_history.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    elif fmt == 'pdf':
        months_data = [{'month':r[0],'emps':r[1],'gross':r[2],'si':r[3],'tax':r[4],
                        'ded':r[5],'loan':r[6],'total_ded':r[7],'net':r[8],'run_date':r[9]}
                       for r in months]
        html = render_template('payroll_history_pdf.html', months=months_data,
                               generated=date.today().strftime('%d/%m/%Y'))
        try:
            from weasyprint import HTML as WH
            pdf = WH(string=html).write_pdf()
            return send_file(io.BytesIO(pdf), as_attachment=True,
                             download_name='payroll_history.pdf', mimetype='application/pdf')
        except ImportError:
            return html, 200, {'Content-Type': 'text/html'}
    else:
        flash('Unknown export format.', 'error')
        return redirect(url_for('payroll_history'))


@app.route('/payroll/history/<month>/export/<fmt>')
@login_required
@role_required('Admin', 'HR')
def payroll_history_detail_export(month, fmt):
    """Export one month's payroll detail as Excel or PDF."""
    if not valid_month(month):
        flash('Invalid month.', 'error')
        return redirect(url_for('payroll_history'))
    _gbp_hde = (", ISNULL(p.GBPComponent,0), ISNULL(p.GBPEGPRate,0)"
                if has_col('PayrollRuns','GBPEGPRate') else ", 0, 0")
    rows = qy(f"""SELECT e.EmpID, ISNULL(e.EmployeeID,''),
                 ISNULL(e.FirstName,'')+' '+ISNULL(e.FamilyName,''),
                 ISNULL(e.Department,''), ISNULL(e.JobTitle,''),
                 ISNULL(e.Nationality,'Egyptian'),
                 p.MonthlyGross, p.SocialInsurance, p.IncomeTax, p.Loan,
                 p.Absences, p.Lateness, p.SickLeave, p.Sanctions, p.OtherDeductions,
                 p.TotalDeductions, p.NetSalary
                 {_gbp_hde}
                 FROM PayrollRuns p
                 JOIN Employees e ON e.EmpID=p.EmpID
                 WHERE p.PayrollMonth=? ORDER BY e.FirstName""", (month,))
    headers = ['Emp ID','Name','Department','Job Title','Nationality',
               'EGP Gross','SI','Tax','Loan',
               'Absences','Lateness','Sick','Sanctions','Other Ded','Total Ded','EGP Net',
               'GBP Pay (£)','GBP Rate']
    rows_data = [[r[1],r[2],r[3],r[4],r[5],
                  r[6],r[7],r[8],r[9],r[10],r[11],r[12],r[13],r[14],r[15],r[16],
                  r[17],r[18]] for r in rows]
    num_start = 6  # columns 6+ are money

    if fmt == 'xlsx':
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = Workbook()
        ws = wb.active
        ws.title = f'Payroll {month}'
        ws.append([f'Payroll Detail — {month}'])
        ws['A1'].font = Font(bold=True, size=13)
        ws.append([])
        hfill = PatternFill('solid', fgColor='4F46E5')
        hfont = Font(bold=True, color='FFFFFF')
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=3, column=ci, value=h)
            c.font = hfont; c.fill = hfill; c.alignment = Alignment(horizontal='center')
        for ri, row in enumerate(rows_data, 4):
            for ci, val in enumerate(row, 1):
                cell = ws.cell(row=ri, column=ci)
                if ci >= num_start and val is not None:
                    try: cell.value = float(val); cell.number_format = '#,##0.00'
                    except: cell.value = val
                else:
                    cell.value = val
        # Totals row
        tr = len(rows_data) + 4
        ws.cell(row=tr, column=1, value='TOTAL').font = Font(bold=True)
        for ci in range(num_start, len(headers)+1):
            col_letter = ws.cell(row=3, column=ci).column_letter
            ws.cell(row=tr, column=ci).value = f'=SUM({col_letter}4:{col_letter}{tr-1})'
            ws.cell(row=tr, column=ci).number_format = '#,##0.00'
            ws.cell(row=tr, column=ci).font = Font(bold=True)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = max(12, len(str(col[2].value or ''))+4)
        buf = io.BytesIO()
        wb.save(buf); buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name=f'payroll_{month}.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    elif fmt == 'pdf':
        detail = [{'emp_id':r[1],'name':r[2],'dept':r[3],'job':r[4],
                   'gross':r[5],'si':r[6],'tax':r[7],'loan':r[8],
                   'absences':r[9],'lateness':r[10],'sick':r[11],
                   'sanctions':r[12],'other_ded':r[13],'total_ded':r[14],'net':r[15]}
                  for r in rows]
        totals = {k: sum(d[k] or 0 for d in detail)
                  for k in ['gross','si','tax','loan','total_ded','net']}
        html = render_template('payroll_history_detail_pdf.html', month=month,
                               detail=detail, totals=totals,
                               generated=date.today().strftime('%d/%m/%Y'))
        try:
            from weasyprint import HTML as WH
            pdf = WH(string=html).write_pdf()
            return send_file(io.BytesIO(pdf), as_attachment=True,
                             download_name=f'payroll_{month}.pdf', mimetype='application/pdf')
        except ImportError:
            return html, 200, {'Content-Type': 'text/html'}
    else:
        flash('Unknown export format.', 'error')
        return redirect(url_for('payroll_history_detail', month=month))


@app.route('/payroll/rollback', methods=['GET', 'POST'])
@login_required
@role_required('Admin', 'HR')
def payroll_rollback():
    months = qy("SELECT DISTINCT PayrollMonth FROM PayrollRuns ORDER BY PayrollMonth DESC") or []
    month_list = [r[0] for r in months]
    if request.method == 'POST':
        m = request.form.get('month', '').strip()
        if not valid_month(m):
            flash('Invalid month.', 'error')
        else:
            ex("DELETE FROM PayrollRuns WHERE PayrollMonth=?", (m,))
            flash(f'Payroll for {m} rolled back successfully.', 'success')
            return redirect(url_for('payroll_history'))
    return render_template('payroll_rollback.html', month_list=month_list)

@app.route('/payroll/settings', methods=['GET', 'POST'])
@login_required
@role_required('Admin', 'HR')
def payroll_settings():
    year = int(request.args.get('year') or date.today().year)
    s = get_payroll_settings(year)
    if request.method == 'POST':
        f = request.form
        try:
            ex("""IF EXISTS (SELECT 1 FROM PayrollSettings WHERE PayrollYear=?)
                      UPDATE PayrollSettings SET DayDivisor=?,HourDivisor=?,SICap=?,SIRate=?,
                        LatenessDivisor=?,PersonalExemptionAnnual=?,
                        AnnualExemption1=?,AnnualExemption2=?,
                        TaxRate1=?,TaxRate2=?,TaxRate3=?,TaxRate4=?,TaxRate5=?,TaxRate6=?,TaxRate7=?,
                        TaxBracket1=?,TaxBracket2=?
                        WHERE PayrollYear=?
                  ELSE
                      INSERT INTO PayrollSettings (PayrollYear,DayDivisor,HourDivisor,SICap,SIRate,
                        LatenessDivisor,PersonalExemptionAnnual,AnnualExemption1,AnnualExemption2,
                        TaxRate1,TaxRate2,TaxRate3,TaxRate4,TaxRate5,TaxRate6,TaxRate7,
                        TaxBracket1,TaxBracket2)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
               (year,
                float(f['DayDivisor']), float(f['HourDivisor']),
                float(f['SICap']), float(f['SIRate']),
                float(f['LatenessDivisor']), float(f['PersonalExemptionAnnual']),
                float(f['AnnualExemption1']), float(f['AnnualExemption2']),
                float(f['TaxRate1']), float(f['TaxRate2']), float(f['TaxRate3']),
                float(f['TaxRate4']), float(f['TaxRate5']), float(f['TaxRate6']), float(f['TaxRate7']),
                float(f['TaxBracket1']), float(f['TaxBracket2']),
                year,
                year,
                float(f['DayDivisor']), float(f['HourDivisor']),
                float(f['SICap']), float(f['SIRate']),
                float(f['LatenessDivisor']), float(f['PersonalExemptionAnnual']),
                float(f['AnnualExemption1']), float(f['AnnualExemption2']),
                float(f['TaxRate1']), float(f['TaxRate2']), float(f['TaxRate3']),
                float(f['TaxRate4']), float(f['TaxRate5']), float(f['TaxRate6']), float(f['TaxRate7']),
                float(f['TaxBracket1']), float(f['TaxBracket2'])))
            flash(f'Payroll settings saved for {year}.', 'success')
        except Exception as e:
            flash(f'Error saving: {e}', 'error')
    return render_template('payroll_settings.html', s=s, year=year)

# ─── Routes: Reports ──────────────────────────────────────────────────────────

@app.route('/reports')
@login_required
@role_required('Admin', 'HR')
def reports_menu():
    return render_template('reports_menu.html')

@app.route('/reports/run', methods=['GET', 'POST'])
@login_required
@role_required('Admin', 'HR')
def reports_run():
    _REPORT_META = {
        'employee_directory':  {'page_title': 'Employee Directory',        'num_text_cols': 7},
        'payroll_summary':     {'page_title': 'Payroll Summary by Month',  'num_text_cols': 3},
        'tax_report':          {'page_title': 'Tax Report',                'num_text_cols': 3},
        'deductions_report':   {'page_title': 'Deductions Report',         'num_text_cols': 3},
        'headcount':           {'page_title': 'Headcount by Department',   'num_text_cols': 2},
        'salary_cost_dept':    {'page_title': 'Salary Cost by Department', 'num_text_cols': 2},
        'ytd_payroll':         {'page_title': 'Year-to-Date Payroll',      'num_text_cols': 3},
        'salary_register':     {'page_title': 'Salary Register',           'num_text_cols': 4},
    }

    report_type   = request.args.get('type', 'payroll_summary')
    meta          = _REPORT_META.get(report_type, {})
    page_title    = meta.get('page_title', 'Report')
    num_text_cols = meta.get('num_text_cols', 3)

    # Load department list for the Employee Directory dropdown
    try:
        dept_rows = qy(f"SELECT DISTINCT Department FROM Employees WHERE {ACTIVE_EMP_SQL} AND Department IS NOT NULL AND LTRIM(RTRIM(Department)) <> '' ORDER BY Department")
        departments = [r[0] for r in dept_rows]
    except Exception:
        departments = []

    ctx = dict(
        report_type=report_type, page_title=page_title,
        num_text_cols=num_text_cols,
        title=None, cols=[], data=[],
        footer_row=None, summary_stats=None,
        month='', year='', dept='',
        departments=departments,
    )

    if request.method == 'GET':
        return render_template('reports_run.html', **ctx)

    # ── POST: collect filter params ───────────────────────────────────────────
    month = (request.form.get('month') or '').strip()   # 'YYYY-MM'
    year  = (request.form.get('year')  or '').strip()   # 'YYYY'
    dept  = (request.form.get('dept')  or '').strip()
    ctx.update(month=month, year=year, dept=dept)

    try:
        # ── 1. Employee Directory ─────────────────────────────────────────────
        if report_type == 'employee_directory':
            sql = f"""SELECT ISNULL(EmployeeID,''), ISNULL(FirstName,'')+' '+ISNULL(FamilyName,''),
                             ISNULL(Department,''), ISNULL(JobTitle,''),
                             ISNULL(Nationality,''), ISNULL(Phone,''),
                             ISNULL(PersonalEmail,''),
                             CONVERT(varchar(10), DateOfJoining, 23)
                      FROM Employees WHERE {ACTIVE_EMP_SQL}"""
            params = []
            if dept:
                sql += " AND Department = ?"
                params.append(dept)
            sql += " ORDER BY Department, FirstName"
            data = [list(r) for r in qy(sql, params)]
            cols = ['Emp ID', 'Name', 'Department', 'Job Title',
                    'Nationality', 'Phone', 'Email', 'Joining Date']
            ctx.update(
                title=f'Employee Directory{" – " + dept if dept else ""}',
                cols=cols, data=data,
                summary_stats=[{'label': 'Total Active', 'value': len(data), 'is_money': False}],
            )

        # ── 2. Payroll Summary by Month ───────────────────────────────────────
        elif report_type == 'payroll_summary':
            if not month:
                return render_template('reports_run.html', **ctx)
            _gbp_ps = (", ISNULL(p.GBPComponent,0), ISNULL(p.GBPEGPRate,0)"
                       if has_col('PayrollRuns','GBPEGPRate') else ", 0, 0")
            data = [list(r) for r in qy(f"""
                SELECT ISNULL(e.EmployeeID,''),
                       ISNULL(e.FirstName,'')+' '+ISNULL(e.FamilyName,''),
                       ISNULL(e.Department,''),
                       ISNULL(e.Nationality,'Egyptian'),
                       ISNULL(p.MonthlyGross,0), ISNULL(p.SocialInsurance,0),
                       ISNULL(p.IncomeTax,0),    ISNULL(p.Loan,0),
                       ISNULL(p.TotalDeductions,0), ISNULL(p.NetSalary,0)
                       {_gbp_ps}
                FROM PayrollRuns p JOIN Employees e ON e.EmpID=p.EmpID
                WHERE p.PayrollMonth=?
                ORDER BY e.Department, e.FirstName
            """, (month,))]
            cols = ['Emp ID', 'Name', 'Department', 'Nationality',
                    'EGP Gross', 'SI', 'Tax', 'Loan', 'Total Ded.', 'EGP Net',
                    'GBP Pay (£)', 'GBP Rate']
            totals = [sum(float(r[i] or 0) for r in data) for i in range(4, 12)]
            footer = ['', 'TOTAL', '', ''] + totals
            ctx.update(
                title=f'Payroll Summary – {month}',
                cols=cols, data=data, footer_row=footer,
                num_text_cols=4,
                summary_stats=[
                    {'label': 'Employees',   'value': len(data),    'is_money': False},
                    {'label': 'Total EGP Gross', 'value': totals[0], 'is_money': True},
                    {'label': 'Total EGP Net',   'value': totals[5], 'is_money': True},
                ],
            )

        # ── 3. Tax Report ─────────────────────────────────────────────────────
        elif report_type == 'tax_report':
            if not month:
                return render_template('reports_run.html', **ctx)
            data = [list(r) for r in qy("""
                SELECT ISNULL(e.EmployeeID,''),
                       ISNULL(e.FirstName,'')+' '+ISNULL(e.FamilyName,''),
                       ISNULL(e.Department,''),
                       ISNULL(p.MonthlyGross,0),
                       ISNULL(p.SocialInsurance,0),
                       ISNULL(p.IncomeTax,0),
                       ISNULL(p.SocialInsurance,0)+ISNULL(p.IncomeTax,0) AS TotalTaxSI
                FROM PayrollRuns p JOIN Employees e ON e.EmpID=p.EmpID
                WHERE p.PayrollMonth=?
                ORDER BY e.Department, e.FirstName
            """, (month,))]
            cols   = ['Emp ID', 'Name', 'Department',
                      'Gross', 'SI Deduction', 'Income Tax', 'Total Tax+SI']
            totals = [sum(r[i] or 0 for r in data) for i in range(3, 7)]
            # totals[0]=Gross [1]=SI [2]=Tax [3]=TotalTaxSI
            footer = ['', 'TOTAL', ''] + totals
            ctx.update(
                title=f'Tax Report – {month}',
                cols=cols, data=data, footer_row=footer,
                summary_stats=[
                    {'label': 'Employees',    'value': len(data),  'is_money': False},
                    {'label': 'Total SI',     'value': totals[1],  'is_money': True},
                    {'label': 'Total Tax',    'value': totals[2],  'is_money': True},
                    {'label': 'Total Tax+SI', 'value': totals[3],  'is_money': True},
                ],
            )

        # ── 4. Deductions Report ──────────────────────────────────────────────
        elif report_type == 'deductions_report':
            if not month:
                return render_template('reports_run.html', **ctx)
            _gbp_ded = (", ISNULL(p.GBPLoanDeduction,0), ISNULL(p.GBPOtherDeductions,0)"
                        if has_col('PayrollRuns','GBPBonuses') else ", 0, 0")
            data = [list(r) for r in qy(f"""
                SELECT ISNULL(e.EmployeeID,''),
                       ISNULL(e.FirstName,'')+' '+ISNULL(e.FamilyName,''),
                       ISNULL(e.Department,''),
                       ISNULL(e.Nationality,'Egyptian'),
                       ISNULL(p.SocialInsurance,0),
                       ISNULL(p.IncomeTax,0),
                       ISNULL(p.Loan,0),
                       ISNULL(p.Absences,0),
                       ISNULL(p.Lateness,0),
                       ISNULL(p.SickLeave,0),
                       ISNULL(p.Sanctions,0),
                       ISNULL(p.OtherDeductions,0),
                       ISNULL(p.TotalDeductions,0)
                       {_gbp_ded}
                FROM PayrollRuns p JOIN Employees e ON e.EmpID=p.EmpID
                WHERE p.PayrollMonth=?
                ORDER BY e.Department, e.FirstName
            """, (month,))]
            cols = ['Emp ID', 'Name', 'Department', 'Nationality',
                    'SI', 'Tax', 'Loan', 'Absences',
                    'Lateness', 'Sick', 'Sanctions', 'Other', 'Total Ded.',
                    'GBP Loan Ded.', 'GBP Other Ded.']
            totals = [sum(float(r[i] or 0) for r in data) for i in range(4, 15)]
            footer = ['', 'TOTAL', '', ''] + totals
            ctx.update(
                title=f'Deductions Report – {month}',
                cols=cols, data=data, footer_row=footer,
                num_text_cols=4,
                summary_stats=[
                    {'label': 'Employees',      'value': len(data),    'is_money': False},
                    {'label': 'Total EGP Ded.', 'value': totals[8],   'is_money': True},
                ],
            )

        # ── 5. Headcount by Department ────────────────────────────────────────
        elif report_type == 'headcount':
            data = [list(r) for r in qy(f"""
                SELECT ISNULL(Department,'Unassigned'),
                       COUNT(*),
                       SUM(CASE WHEN LOWER(ISNULL(Nationality,'egyptian'))='egyptian' THEN 1 ELSE 0 END),
                       SUM(CASE WHEN LOWER(ISNULL(Nationality,'egyptian'))<>'egyptian' THEN 1 ELSE 0 END),
                       SUM(ISNULL(BasicSalary,0)),
                       SUM(ISNULL(HousingAllowance,0)+ISNULL(TransportAllowance,0)+ISNULL(OtherFixedAllowances,0)),
                       SUM(ISNULL(BasicSalary,0)+ISNULL(HousingAllowance,0)+ISNULL(TransportAllowance,0)+ISNULL(OtherFixedAllowances,0))
                FROM Employees WHERE {ACTIVE_EMP_SQL}
                GROUP BY Department ORDER BY Department
            """)]
            cols   = ['Department', 'Total', 'Egyptian', 'Expat (Non-Egyptian)',
                      'Total Basic (EGP)', 'Total Allowances (EGP)', 'Total Gross (EGP)']
            footer = ['TOTAL',
                      sum(r[1] for r in data),
                      sum(r[2] or 0 for r in data),
                      sum(r[3] or 0 for r in data),
                      sum(r[4] or 0 for r in data),
                      sum(r[5] or 0 for r in data),
                      sum(r[6] or 0 for r in data)]
            ctx.update(
                title='Headcount by Department',
                cols=cols, data=data, footer_row=footer,
                num_text_cols=2,
                summary_stats=[
                    {'label': 'Departments',  'value': len(data),                    'is_money': False},
                    {'label': 'Total Emps',   'value': sum(r[1] for r in data),      'is_money': False},
                    {'label': 'Egyptians',    'value': sum(r[2] or 0 for r in data), 'is_money': False},
                    {'label': 'Expats',       'value': sum(r[3] or 0 for r in data), 'is_money': False},
                    {'label': 'Total Gross',  'value': sum(r[6] or 0 for r in data), 'is_money': True},
                ],
            )

        # ── 6. Salary Cost by Department ──────────────────────────────────────
        elif report_type == 'salary_cost_dept':
            if not month:
                return render_template('reports_run.html', **ctx)
            data = [list(r) for r in qy("""
                SELECT ISNULL(e.Department,'Unassigned'),
                       COUNT(DISTINCT p.EmpID),
                       SUM(ISNULL(p.MonthlyGross,0)),
                       SUM(ISNULL(p.TotalDeductions,0)),
                       SUM(ISNULL(p.NetSalary,0))
                FROM PayrollRuns p JOIN Employees e ON e.EmpID=p.EmpID
                WHERE p.PayrollMonth=?
                GROUP BY e.Department ORDER BY e.Department
            """, (month,))]
            cols   = ['Department', 'Headcount', 'Total Gross', 'Total Deductions', 'Total Net']
            footer = ['TOTAL',
                      sum(r[1] for r in data),
                      sum(r[2] or 0 for r in data),
                      sum(r[3] or 0 for r in data),
                      sum(r[4] or 0 for r in data)]
            ctx.update(
                title=f'Salary Cost by Department – {month}',
                cols=cols, data=data, footer_row=footer,
                num_text_cols=2,
                summary_stats=[
                    {'label': 'Departments', 'value': len(data),                    'is_money': False},
                    {'label': 'Total Gross', 'value': sum(r[2] or 0 for r in data), 'is_money': True},
                    {'label': 'Total Net',   'value': sum(r[4] or 0 for r in data), 'is_money': True},
                ],
            )

        # ── 7. Year-to-Date Payroll ───────────────────────────────────────────
        elif report_type == 'ytd_payroll':
            if not year:
                return render_template('reports_run.html', **ctx)
            data = [list(r) for r in qy("""
                SELECT ISNULL(e.EmployeeID,''),
                       ISNULL(e.FirstName,'')+' '+ISNULL(e.FamilyName,''),
                       ISNULL(e.Department,''),
                       SUM(ISNULL(p.MonthlyGross,0))     AS YTDGross,
                       SUM(ISNULL(p.SocialInsurance,0))  AS YTDSI,
                       SUM(ISNULL(p.IncomeTax,0))        AS YTDTax,
                       SUM(ISNULL(p.Loan,0))             AS YTDLoan,
                       SUM(ISNULL(p.TotalDeductions,0))  AS YTDTotalDed,
                       SUM(ISNULL(p.NetSalary,0))        AS YTDNet,
                       COUNT(p.PayrollMonth)             AS MonthsPaid
                FROM PayrollRuns p JOIN Employees e ON e.EmpID=p.EmpID
                WHERE LEFT(p.PayrollMonth,4)=?
                GROUP BY e.EmpID, e.EmployeeID, e.FirstName, e.FamilyName, e.Department
                ORDER BY e.Department, e.FirstName
            """, (year,))]
            cols   = ['Emp ID', 'Name', 'Department',
                      'YTD Gross', 'YTD SI', 'YTD Tax',
                      'YTD Loan', 'YTD Total Ded.', 'YTD Net', 'Months Paid']
            totals = [sum(r[i] or 0 for r in data) for i in range(3, 10)]
            footer = ['', 'TOTAL', ''] + totals
            ctx.update(
                title=f'Year-to-Date Payroll – {year}',
                cols=cols, data=data, footer_row=footer,
                summary_stats=[
                    {'label': 'Employees', 'value': len(data),  'is_money': False},
                    {'label': 'YTD Gross', 'value': totals[0],  'is_money': True},
                    {'label': 'YTD Net',   'value': totals[5],  'is_money': True},
                ],
            )

        # ── 8. Salary Register ────────────────────────────────────────────────
        elif report_type == 'salary_register':
            _gbp_col = ", ISNULL(GBPSalary,0)" if has_col('Employees','GBPSalary') else ", 0"
            data = [list(r) for r in qy(f"""
                SELECT ISNULL(EmployeeID,''),
                       ISNULL(FirstName,'')+' '+ISNULL(FamilyName,''),
                       ISNULL(Department,''), ISNULL(JobTitle,''),
                       ISNULL(Nationality,'Egyptian'),
                       ISNULL(BasicSalary,0), ISNULL(HousingAllowance,0),
                       ISNULL(TransportAllowance,0), ISNULL(OtherFixedAllowances,0),
                       ISNULL(BasicSalary,0)+ISNULL(HousingAllowance,0)+
                       ISNULL(TransportAllowance,0)+ISNULL(OtherFixedAllowances,0)
                       {_gbp_col}
                FROM Employees WHERE {ACTIVE_EMP_SQL} ORDER BY Department, FirstName
            """)]
            cols   = ['Emp ID', 'Name', 'Department', 'Job Title', 'Nationality',
                      'Basic (EGP)', 'Housing (EGP)', 'Transport (EGP)', 'Other (EGP)',
                      'Total EGP Gross', 'GBP Salary (£)']
            totals = [sum(float(r[i] or 0) for r in data) for i in range(5, 11)]
            footer = ['', 'TOTAL', '', '', ''] + totals
            ctx.update(
                title='Salary Register',
                cols=cols, data=data, footer_row=footer,
                num_text_cols=5,
                summary_stats=[
                    {'label': 'Active Employees', 'value': len(data),    'is_money': False},
                    {'label': 'Total EGP Payroll', 'value': totals[4],   'is_money': True},
                ],
            )

    except Exception as e:
        ctx['error'] = str(e)

    return render_template('reports_run.html', **ctx)

# ─── Routes: User Manager ─────────────────────────────────────────────────────

@app.route('/users')
@login_required
@role_required('Admin')
def user_manager():
    if not app.config.get('_users_cols_checked'):
        _detect_users_cols()
        app.config['_users_cols_checked'] = True
    active_col = app.config.get('_active_col', 'IsActive')
    active_sel = f"ISNULL({active_col},1)" if active_col else "1"
    rows = qy(f"SELECT UserID, Username, ISNULL(FullName,Username), Role, {active_sel} FROM Users ORDER BY Username") or []
    users = [{'uid': r[0], 'username': r[1], 'fullname': r[2],
               'role': r[3], 'active': r[4]} for r in rows]
    return render_template('user_manager.html', users=users, roles=list(ROLE_LABELS.keys()))

@app.route('/users/save', methods=['POST'])
@login_required
@role_required('Admin')
def user_save():
    f = request.form
    uid = f.get('uid', '').strip()
    username = f.get('username', '').strip()
    fullname = f.get('fullname', '').strip()
    role     = f.get('role', '').strip()
    pwd      = f.get('password', '').strip()
    active   = 1 if f.get('active') else 0
    if not app.config.get('_users_cols_checked'):
        _detect_users_cols()
        app.config['_users_cols_checked'] = True
    pwd_col    = app.config.get('_pwd_col', 'PasswordHash')
    active_col = app.config.get('_active_col')
    try:
        if uid:
            if pwd and active_col:
                ex(f"UPDATE Users SET FullName=?,Role=?,{active_col}=?,{pwd_col}=? WHERE UserID=?",
                   (fullname, role, active, pwd, uid))
            elif pwd:
                ex(f"UPDATE Users SET FullName=?,Role=?,{pwd_col}=? WHERE UserID=?",
                   (fullname, role, pwd, uid))
            elif active_col:
                ex(f"UPDATE Users SET FullName=?,Role=?,{active_col}=? WHERE UserID=?",
                   (fullname, role, active, uid))
            else:
                ex("UPDATE Users SET FullName=?,Role=? WHERE UserID=?",
                   (fullname, role, uid))
            flash('User updated.', 'success')
        else:
            if active_col:
                ex(f"INSERT INTO Users (Username,FullName,Role,{pwd_col},{active_col}) VALUES (?,?,?,?,?)",
                   (username, fullname, role, pwd, active))
            else:
                ex(f"INSERT INTO Users (Username,FullName,Role,{pwd_col}) VALUES (?,?,?,?)",
                   (username, fullname, role, pwd))
            flash('User created.', 'success')
    except Exception as e:
        flash(f'Error: {e}', 'error')
    return redirect(url_for('user_manager'))

@app.route('/users/<int:uid>/delete', methods=['POST'])
@login_required
@role_required('Admin')
def user_delete(uid):
    try:
        ex("DELETE FROM Users WHERE UserID=?", (uid,))
        flash('User deleted.', 'success')
    except Exception as e:
        flash(f'Error: {e}', 'error')
    return redirect(url_for('user_manager'))

# ─── API helpers ──────────────────────────────────────────────────────────────

@app.route('/api/payroll/emp-earnings', methods=['GET'])
@login_required
def api_emp_earnings():
    """Return earnings breakdown for a given employee (for real-time deduction display)."""
    eid = request.args.get('eid')
    if not eid:
        return jsonify({'error': 'No employee'}), 400
    emp = _enrich_emp_expat(load_emp_basic(int(eid)))
    if not emp:
        return jsonify({'error': 'Not found'}), 404
    earnings = emp['earnings']
    s = emp['settings']
    basic = float(earnings.get('BasicSalary') or 0)
    gross = sum(float(earnings.get(k) or 0) for k in EARNING_COLS)
    div = float(s.get('DayDivisor', 30))
    day_rate = basic / div if div else 0
    # Check if this employee already has a saved run for the requested month
    req_month = request.args.get('month', date.today().strftime('%Y-%m'))
    saved_check = qy("SELECT 1 FROM PayrollRuns WHERE EmpID=? AND PayrollMonth=?",
                     (emp['eid'], req_month))
    return jsonify({
        'name': emp['name'],
        'emp_id': emp['emp_id'],
        'dept': emp['dept'],
        'job': emp['job'],
        'earnings': {k: round(float(earnings.get(k) or 0), 2) for k in EARNING_COLS},
        'gross': round(gross, 2),
        'day_rate': round(day_rate, 2),
        'settings': {k: float(v) for k, v in s.items()},
        'already_saved': bool(saved_check),
        'employee_type': emp.get('employee_type', 'Egyptian') or 'Egyptian',
        'nationality': emp.get('nationality', 'Egyptian') or 'Egyptian',
        'is_expat': bool(emp.get('is_expat', False)),
        'gbp_salary': float(emp.get('gbp_salary') or 0),
    })

@app.route('/api/payroll/preview', methods=['POST'])
@login_required
def api_payroll_preview():
    data = request.json or {}
    eid = data.get('eid')
    if not eid:
        return jsonify({'error': 'No employee selected'}), 400
    emp = load_emp_basic(int(eid))
    if not emp:
        return jsonify({'error': 'Employee not found'}), 404
    s = emp['settings']
    sanction_days = float(data.get('sanction_days') or 0)
    other_amt = data.get('other_amount')
    other_egp = float(other_amt) if other_amt not in (None, '') else None
    monthly_earn = dict(emp['earnings'])
    monthly_earn.update({
        'VariableAllowances':    float(data.get('variable_allowances') or 0),
        'Overtime':              float(data.get('overtime') or 0),
        'Bonuses':               float(data.get('bonuses') or 0),
        'Commissions':           float(data.get('commissions') or 0),
        'Reimbursement':         float(data.get('reimbursement') or 0),
        'ReallocationAllowance': float(data.get('reallocation_allowance') or 0),
    })
    r = run_payroll(
        monthly_earn,
        absent_days=float(data.get('absent_days') or 0),
        late_minutes=float(data.get('late_minutes') or 0),
        sick_days=float(data.get('sick_days') or 0),
        sanction_days=sanction_days,
        loan=float(data.get('loan') or 0),
        settings=s,
        other_amount=other_egp,
    )
    return jsonify({
        'gross': round(r['gross'], 2),
        'net': round(r['net'], 2),
        'si': round(r['tax']['si_monthly'], 2),
        'tax': round(r['tax']['monthly_tax'], 2),
        'total_ded': round(r['total_ded'], 2),
        'absences': round(r['ded']['absences'], 2),
        'lateness': round(r['ded']['lateness'], 2),
        'sick': round(r['ded']['sick'], 2),
        'sanctions': round(r['ded']['sanctions'], 2),
        'other_ded': round(r['ded']['other'], 2),
        'day_rate': round(r['ded']['day_rate'], 2),
        'taxable': round(r['tax']['taxable_monthly'], 2),
        'taxable_annual': round(r['tax']['taxable_annual'], 2),
        'annual_tax': round(r['tax']['annual_tax'], 2),
        'net_before_tax': round(r['tax']['net_before_tax'], 2),
        'pe': round(r['tax']['pe_monthly'], 2),
        'brackets': [{'label': b['label'], 'rate': b['rate'],
                      'base': round(b['base'], 2), 'tax': round(b['tax'], 2)}
                     for b in r['tax'].get('brackets', [])],
    })

@app.route('/payroll/export/pdf')
@login_required
@role_required('Admin', 'HR')
def payroll_export_pdf():
    eid = request.args.get('eid')
    month = request.args.get('month')
    if not eid or not valid_month(month or ''):
        flash('Invalid export parameters.', 'error')
        return redirect(url_for('payroll_manual'))
    emp = load_emp_basic(int(eid))
    if not emp:
        flash('Employee not found.', 'error')
        return redirect(url_for('payroll_manual'))
    # Load saved run if available
    emp = _enrich_emp_expat(emp)
    _fa_sel2    = ", ISNULL(FlightAllowance,0)"    if has_col('PayrollRuns','FlightAllowance')    else ", 0"
    _gbp_sel2   = ", ISNULL(GBPEGPRate,0), ISNULL(GBPComponent,0), ISNULL(EGPComponent,0)" \
                  if has_col('PayrollRuns','GBPEGPRate') else ", 0, 0, 0"
    _gbped_sel2 = (", ISNULL(GBPVariableAllowances,0), ISNULL(GBPOvertime,0)"
                   ", ISNULL(GBPBonuses,0), ISNULL(GBPCommissions,0)"
                   ", ISNULL(GBPFlightAllowance,0), ISNULL(GBPReimbursement,0)"
                   ", ISNULL(GBPReallocationAllowance,0), ISNULL(GBPOtherEarnings,0)"
                   ", ISNULL(GBPLoanDeduction,0), ISNULL(GBPOtherDeductions,0)"
                   + (", ISNULL(GBPAbsentDays,0), ISNULL(GBPLateMinutes,0)"
                      ", ISNULL(GBPSickDays,0), ISNULL(GBPSanctionDays,0)"
                      if has_col('PayrollRuns','GBPAbsentDays') else ", 0, 0, 0, 0")
                   ) if has_col('PayrollRuns','GBPBonuses') else ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0"
    row = qy(f"""SELECT MonthlyGross,SocialInsurance,IncomeTax,Loan,
                       Absences,Lateness,SickLeave,Sanctions,OtherDeductions,
                       TotalDeductions,NetSalary,DayRate,Notes,
                       ISNULL(VariableAllowances,0),ISNULL(Overtime,0),
                       ISNULL(Bonuses,0),ISNULL(Commissions,0),
                       ISNULL(Reimbursement,0),ISNULL(ReallocationAllowance,0)
                       {_fa_sel2}{_gbp_sel2}{_gbped_sel2}
                FROM PayrollRuns WHERE EmpID=? AND PayrollMonth=?""", (eid, month))
    gbp_data = None
    if row:
        r0 = row[0]
        gross = float(r0[0] or 0); si = float(r0[1] or 0); tax = float(r0[2] or 0)
        loan  = float(r0[3] or 0)
        absences=float(r0[4] or 0); lateness=float(r0[5] or 0); sick=float(r0[6] or 0)
        sanctions=float(r0[7] or 0); other_ded=float(r0[8] or 0)
        total_ded=float(r0[9] or 0); net=float(r0[10] or 0)
        day_rate=float(r0[11] or 0); notes=r0[12] or ''
        # Overlay monthly variable earnings for payslip display
        emp['earnings'].update({
            'VariableAllowances':    float(r0[13] or 0),
            'Overtime':              float(r0[14] or 0),
            'Bonuses':               float(r0[15] or 0),
            'Commissions':           float(r0[16] or 0),
            'Reimbursement':         float(r0[17] or 0),
            'ReallocationAllowance': float(r0[18] or 0),
            'FlightAllowance':       float(r0[19] or 0) if len(r0) > 19 else 0,
        })
        # GBP data for Expat payslip section — build for ALL expats (rate may be 0)
        _gbp_rate_pdf = float(r0[20] or 0) if len(r0) > 20 else 0
        if emp.get('is_expat'):
            gbp_data = {
                'rate':                       _gbp_rate_pdf,
                'gbp_component':              float(r0[21] or 0) if len(r0) > 21 else 0,
                'egp_component':              float(r0[22] or 0) if len(r0) > 22 else 0,
                'gbp_variable_allowances':    float(r0[23] or 0) if len(r0) > 23 else 0,
                'gbp_overtime':               float(r0[24] or 0) if len(r0) > 24 else 0,
                'gbp_bonuses':                float(r0[25] or 0) if len(r0) > 25 else 0,
                'gbp_commissions':            float(r0[26] or 0) if len(r0) > 26 else 0,
                'gbp_flight_allowance':       float(r0[27] or 0) if len(r0) > 27 else 0,
                'gbp_reimbursement':          float(r0[28] or 0) if len(r0) > 28 else 0,
                'gbp_reallocation_allowance': float(r0[29] or 0) if len(r0) > 29 else 0,
                'gbp_other_earnings':         float(r0[30] or 0) if len(r0) > 30 else 0,
                'gbp_loan_deduction':         float(r0[31] or 0) if len(r0) > 31 else 0,
                'gbp_other_deductions':       float(r0[32] or 0) if len(r0) > 32 else 0,
                'gbp_absent_days':            float(r0[33] or 0) if len(r0) > 33 else 0,
                'gbp_late_minutes':           float(r0[34] or 0) if len(r0) > 34 else 0,
                'gbp_sick_days':              float(r0[35] or 0) if len(r0) > 35 else 0,
                'gbp_sanction_days':          float(r0[36] or 0) if len(r0) > 36 else 0,
            }
            # Compute GBP deduction amounts for PDF display
            _gbp_pdf_wdays = float((emp.get('settings') or {}).get('working_days', 22) or 22)
            _gbp_pdf_dr    = round(gbp_data['gbp_component'] / _gbp_pdf_wdays, 4) if _gbp_pdf_wdays else 0
            _gbp_pdf_wmins = float((emp.get('settings') or {}).get('work_minutes', 480) or 480)
            gbp_data['gbp_absent_amount']   = round(gbp_data.get('gbp_absent_days', 0)   * _gbp_pdf_dr, 2)
            gbp_data['gbp_late_amount']     = round(gbp_data.get('gbp_late_minutes', 0)  * (_gbp_pdf_dr / _gbp_pdf_wmins), 2) if _gbp_pdf_wmins else 0
            gbp_data['gbp_sick_amount']     = round(gbp_data.get('gbp_sick_days', 0)     * _gbp_pdf_dr, 2)
            gbp_data['gbp_sanction_amount'] = round(gbp_data.get('gbp_sanction_days', 0) * _gbp_pdf_dr, 2)
    else:
        r = run_payroll(emp['earnings'], settings=emp['settings'])
        gross=r['gross']; si=r['tax']['si_monthly']; tax=r['tax']['monthly_tax']
        loan=0; absences=lateness=sick=sanctions=other_ded=0
        total_ded=r['total_ded']; net=r['net']; day_rate=r['ded']['day_rate']; notes=''

    # Recalculate for bracket detail
    r_full = run_payroll(emp['earnings'], settings=emp['settings'])
    tax_detail = r_full['tax']
    html_content = render_template('payslip_pdf.html',
        emp=emp, month=month, gross=gross, si=si, tax=tax, loan=loan,
        absences=absences, lateness=lateness, sick=sick, sanctions=sanctions,
        other_ded=other_ded, total_ded=total_ded, net=net, day_rate=day_rate,
        notes=notes, tax_detail=tax_detail, gbp_data=gbp_data,
        today_str=date.today().strftime('%d %b %Y'))
    pdf_bytes, fmt = _html_to_pdf(html_content)
    if fmt == 'pdf':
        fname = f"Payslip_{emp.get('emp_id') or emp.get('eid','0')}_{month}.pdf"
        return send_file(io.BytesIO(pdf_bytes), mimetype='application/pdf',
                         as_attachment=True, download_name=fname)
    from flask import Response
    return Response(html_content, mimetype='text/html')

@app.route('/payroll/export/excel')
@login_required
@role_required('Admin', 'HR')
def payroll_export_excel():
    eid = request.args.get('eid')
    month = request.args.get('month')
    if not eid or not valid_month(month or ''):
        flash('Invalid export parameters.', 'error')
        return redirect(url_for('payroll_manual'))
    emp = load_emp_basic(int(eid))
    if not emp:
        flash('Employee not found.', 'error')
        return redirect(url_for('payroll_manual'))
    try:
        import openpyxl  # noqa: F401 — verify available before heavy work
    except ImportError:
        flash('Excel export requires openpyxl. Run: pip install openpyxl on the server and restart.', 'error')
        return redirect(url_for('payroll_manual'))
    # ── same rich query as PDF export (includes GBP columns) ──────────────
    _fa_sel_xl    = ", ISNULL(FlightAllowance,0)"    if has_col('PayrollRuns','FlightAllowance')    else ", 0"
    _gbp_sel_xl   = ", ISNULL(GBPEGPRate,0), ISNULL(GBPComponent,0), ISNULL(EGPComponent,0)" \
                    if has_col('PayrollRuns','GBPEGPRate') else ", 0, 0, 0"
    _gbped_sel_xl = (", ISNULL(GBPVariableAllowances,0), ISNULL(GBPOvertime,0)"
                     ", ISNULL(GBPBonuses,0), ISNULL(GBPCommissions,0)"
                     ", ISNULL(GBPFlightAllowance,0), ISNULL(GBPReimbursement,0)"
                     ", ISNULL(GBPReallocationAllowance,0), ISNULL(GBPOtherEarnings,0)"
                     ", ISNULL(GBPLoanDeduction,0), ISNULL(GBPOtherDeductions,0)"
                     + (", ISNULL(GBPAbsentDays,0), ISNULL(GBPLateMinutes,0)"
                        ", ISNULL(GBPSickDays,0), ISNULL(GBPSanctionDays,0)"
                        if has_col('PayrollRuns','GBPAbsentDays') else ", 0, 0, 0, 0")
                     ) if has_col('PayrollRuns','GBPBonuses') else ", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0"
    row = qy(f"""SELECT MonthlyGross,SocialInsurance,IncomeTax,Loan,
                       Absences,Lateness,SickLeave,Sanctions,OtherDeductions,
                       TotalDeductions,NetSalary,DayRate,Notes,
                       ISNULL(VariableAllowances,0),ISNULL(Overtime,0),
                       ISNULL(Bonuses,0),ISNULL(Commissions,0),
                       ISNULL(Reimbursement,0),ISNULL(ReallocationAllowance,0)
                       {_fa_sel_xl}{_gbp_sel_xl}{_gbped_sel_xl}
                FROM PayrollRuns WHERE EmpID=? AND PayrollMonth=?""", (eid, month))

    gbp_data = None
    if row:
        r0 = row[0]
        gross=float(r0[0] or 0); si=float(r0[1] or 0); tax=float(r0[2] or 0)
        loan=float(r0[3] or 0)
        absences=float(r0[4] or 0); lateness=float(r0[5] or 0); sick=float(r0[6] or 0)
        sanctions=float(r0[7] or 0); other_ded=float(r0[8] or 0)
        total_ded=float(r0[9] or 0); net=float(r0[10] or 0); notes=r0[12] or ''
        # Update earnings with monthly variable values saved in PayrollRuns
        emp['earnings'].update({
            'VariableAllowances':    float(r0[13] or 0),
            'Overtime':              float(r0[14] or 0),
            'Bonuses':               float(r0[15] or 0),
            'Commissions':           float(r0[16] or 0),
            'Reimbursement':         float(r0[17] or 0),
            'ReallocationAllowance': float(r0[18] or 0),
            'FlightAllowance':       float(r0[19] or 0) if len(r0) > 19 else 0,
        })
        # Build GBP data for expat employees
        if emp.get('is_expat'):
            gbp_data = {
                'rate':                       float(r0[20] or 0) if len(r0) > 20 else 0,
                'gbp_component':              float(r0[21] or 0) if len(r0) > 21 else 0,
                'egp_component':              float(r0[22] or 0) if len(r0) > 22 else 0,
                'gbp_variable_allowances':    float(r0[23] or 0) if len(r0) > 23 else 0,
                'gbp_overtime':               float(r0[24] or 0) if len(r0) > 24 else 0,
                'gbp_bonuses':               float(r0[25] or 0) if len(r0) > 25 else 0,
                'gbp_commissions':            float(r0[26] or 0) if len(r0) > 26 else 0,
                'gbp_flight_allowance':       float(r0[27] or 0) if len(r0) > 27 else 0,
                'gbp_reimbursement':          float(r0[28] or 0) if len(r0) > 28 else 0,
                'gbp_reallocation_allowance': float(r0[29] or 0) if len(r0) > 29 else 0,
                'gbp_other_earnings':         float(r0[30] or 0) if len(r0) > 30 else 0,
                'gbp_loan_deduction':         float(r0[31] or 0) if len(r0) > 31 else 0,
                'gbp_other_deductions':       float(r0[32] or 0) if len(r0) > 32 else 0,
                'gbp_absent_days':            float(r0[33] or 0) if len(r0) > 33 else 0,
                'gbp_late_minutes':           float(r0[34] or 0) if len(r0) > 34 else 0,
                'gbp_sick_days':              float(r0[35] or 0) if len(r0) > 35 else 0,
                'gbp_sanction_days':          float(r0[36] or 0) if len(r0) > 36 else 0,
            }
            _xl_wdays = float((emp.get('settings') or {}).get('working_days', 22) or 22)
            _xl_dr    = round(gbp_data['gbp_component'] / _xl_wdays, 4) if _xl_wdays else 0
            _xl_wmins = float((emp.get('settings') or {}).get('work_minutes', 480) or 480)
            gbp_data['gbp_absent_amount']   = round(gbp_data['gbp_absent_days']   * _xl_dr, 2)
            gbp_data['gbp_late_amount']     = round(gbp_data['gbp_late_minutes']  * (_xl_dr / _xl_wmins), 2) if _xl_wmins else 0
            gbp_data['gbp_sick_amount']     = round(gbp_data['gbp_sick_days']     * _xl_dr, 2)
            gbp_data['gbp_sanction_amount'] = round(gbp_data['gbp_sanction_days'] * _xl_dr, 2)
    else:
        r = run_payroll(emp['earnings'], settings=emp['settings'])
        gross=r['gross']; si=r['tax']['si_monthly']; tax=r['tax']['monthly_tax']
        loan=absences=lateness=sick=sanctions=other_ded=0
        total_ded=r['total_ded']; net=r['net']; notes=''

    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Payslip {month}"
        ws.column_dimensions['A'].width = 36
        ws.column_dimensions['B'].width = 22

        def hdr(row, txt, col='4F46E5'):
            c = ws.cell(row=row, column=1, value=txt)
            c.font = Font(bold=True, color='FFFFFF', size=11)
            c.fill = PatternFill('solid', fgColor=col)
            c.alignment = Alignment(horizontal='left')
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)

        def row2(r, lbl, val, bold=False, red=False, prefix=''):
            c1 = ws.cell(row=r, column=1, value=lbl)
            display_val = round(float(val or 0), 2)
            c2 = ws.cell(row=r, column=2, value=(prefix + f'{display_val:,.2f}') if prefix else display_val)
            if not prefix:
                c2.number_format = '#,##0.00'
            c2.alignment = Alignment(horizontal='right')
            if bold:
                c1.font = Font(bold=True, color='CC0000' if red else '000000')
                c2.font = Font(bold=True, color='CC0000' if red else '000000')
            elif red:
                c1.font = Font(color='CC0000')
                c2.font = Font(color='CC0000')

        rn = 1
        t = ws.cell(row=rn, column=1, value='KINGS SCHOOL EGYPT — PAYSLIP')
        t.font = Font(bold=True, size=14)
        ws.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=2); rn+=1
        ws.cell(row=rn, column=1, value=f"Employee: {emp['name']}  |  ID: {emp.get('emp_id','')}")
        ws.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=2); rn+=1
        ws.cell(row=rn, column=1, value=f"Month: {month}  |  Nationality: {emp.get('nationality','')}")
        ws.merge_cells(start_row=rn, start_column=1, end_row=rn, end_column=2); rn+=2

        if emp.get('is_expat') and gbp_data:
            # ── EXPAT: GBP payslip only ──────────────────────────────────
            gd = gbp_data
            gbp_earn_tot = (gd.get('gbp_component', 0) + gd.get('gbp_variable_allowances', 0)
                            + gd.get('gbp_overtime', 0) + gd.get('gbp_bonuses', 0)
                            + gd.get('gbp_commissions', 0) + gd.get('gbp_flight_allowance', 0)
                            + gd.get('gbp_reimbursement', 0) + gd.get('gbp_reallocation_allowance', 0)
                            + gd.get('gbp_other_earnings', 0))
            gbp_ded_tot  = (gd.get('gbp_absent_amount', 0) + gd.get('gbp_late_amount', 0)
                            + gd.get('gbp_sick_amount', 0) + gd.get('gbp_sanction_amount', 0)
                            + gd.get('gbp_other_deductions', 0) + gd.get('gbp_loan_deduction', 0))
            gbp_net_val  = gbp_earn_tot - gbp_ded_tot

            hdr(rn, 'GBP EARNINGS (EXPAT)'); rn+=1
            if gd.get('rate', 0):
                row2(rn, '  GBP/EGP Rate', gd['rate'], prefix=''); rn+=1
            row2(rn, '  Base GBP Pay (75%)', gd.get('gbp_component', 0), prefix='£'); rn+=1
            for key, lbl in [('gbp_variable_allowances','Variable Allowances'),
                              ('gbp_overtime','Overtime'),('gbp_bonuses','Bonuses'),
                              ('gbp_commissions','Commissions/Other'),
                              ('gbp_flight_allowance','Flight Allowance'),
                              ('gbp_reimbursement','Re-Imbursement'),
                              ('gbp_reallocation_allowance','Re-Allocation Allowance'),
                              ('gbp_other_earnings','Other Earnings')]:
                v = gd.get(key, 0)
                if v: row2(rn, '  '+lbl, v, prefix='£'); rn+=1
            row2(rn, 'Total GBP Earnings', gbp_earn_tot, bold=True, prefix='£'); rn+=2

            hdr(rn, 'GBP DEDUCTIONS', col='DC2626'); rn+=1
            for key, lbl in [('gbp_absent_amount','Absences'),('gbp_late_amount','Late/Early'),
                              ('gbp_sick_amount','Sick Leave'),('gbp_sanction_amount','Sanctions'),
                              ('gbp_other_deductions','Other Deductions'),
                              ('gbp_loan_deduction','Loan Deduction')]:
                v = gd.get(key, 0)
                if v: row2(rn, '  '+lbl, v, red=True, prefix='— £'); rn+=1
            row2(rn, 'TOTAL GBP DEDUCTIONS', gbp_ded_tot, bold=True, red=True, prefix='— £'); rn+=2

            hdr(rn, 'GBP NET PAY', col='059669'); rn+=1
            row2(rn, 'GBP NET SALARY', gbp_net_val, bold=True, prefix='£'); rn+=1
            # ── EGP section for expat (25% portion) ─────────────────────
            rn+=1
            hdr(rn, 'EGP EARNINGS (25% PORTION)'); rn+=1
            for col_key, lbl in [('BasicSalary','Basic Salary (EGP)'),
                                  ('HousingAllowance','Housing Allowance'),
                                  ('TransportAllowance','Transport Allowance'),
                                  ('OtherFixedAllowances','Other Fixed Allowances'),
                                  ('VariableAllowances','Variable Allowances'),
                                  ('Overtime','Overtime'),('Bonuses','Bonuses'),
                                  ('Commissions','Commissions'),('Reimbursement','Re-Imbursement'),
                                  ('ReallocationAllowance','Re-Allocation Allowance'),
                                  ('FlightAllowance','Flight Allowance')]:
                v = float(emp['earnings'].get(col_key) or 0)
                if v: row2(rn, '  '+lbl, v, prefix='EGP '); rn+=1
            row2(rn, 'Total EGP Gross', gross, bold=True, prefix='EGP '); rn+=2

            hdr(rn, 'EGP DEDUCTIONS', col='DC2626'); rn+=1
            for lbl, v in [('Social Insurance (11%)', si),('Income Tax', tax),
                            ('Absences', absences),('Late/Early', lateness),
                            ('Sick Leave', sick),('Sanctions', sanctions),
                            ('Other Deductions', other_ded),('Loan', loan)]:
                if v: row2(rn, '  '+lbl, v, red=True, prefix='EGP '); rn+=1
            row2(rn, 'TOTAL EGP DEDUCTIONS', total_ded, bold=True, red=True, prefix='EGP '); rn+=2

            hdr(rn, 'EGP NET (25% PORTION)', col='059669'); rn+=1
            row2(rn, 'EGP Net Salary', net, bold=True, prefix='EGP '); rn+=1

        else:
            # ── EGYPTIAN: EGP payslip only ───────────────────────────────
            hdr(rn, 'GROSS EARNINGS (EGP)'); rn+=1
            for col_key, lbl in [('BasicSalary','Basic Salary'),('HousingAllowance','Housing Allowance'),
                                  ('TransportAllowance','Transport Allowance'),
                                  ('OtherFixedAllowances','Other Fixed Allowances'),
                                  ('VariableAllowances','Variable Allowances'),
                                  ('Overtime','Overtime'),('Bonuses','Bonuses'),
                                  ('Commissions','Commissions'),('Reimbursement','Re-Imbursement'),
                                  ('ReallocationAllowance','Re-Allocation Allowance'),
                                  ('FlightAllowance','Flight Allowance'),
                                  ('OtherEarnings','Other Earnings')]:
                v = float(emp['earnings'].get(col_key) or 0)
                if v: row2(rn, '  '+lbl, v); rn+=1
            row2(rn, 'Total Gross (EGP)', gross, bold=True); rn+=2

            hdr(rn, 'DEDUCTIONS (EGP)', col='DC2626'); rn+=1
            for lbl, v in [('Absences', absences),('Late/Early', lateness),('Sick Leave', sick),
                            ('Sanctions', sanctions),('Other Deductions', other_ded),
                            ('Loan', loan),('Social Insurance (11%)', si),('Income Tax', tax)]:
                if v: row2(rn, '  '+lbl, v, red=True); rn+=1
            row2(rn, 'TOTAL DEDUCTIONS (EGP)', total_ded, bold=True, red=True); rn+=2

            hdr(rn, 'NET TAKE HOME (EGP)', col='059669'); rn+=1
            row2(rn, 'NET SALARY (EGP)', net, bold=True); rn+=1

        buf = io.BytesIO()
        wb.save(buf); buf.seek(0)
        fname = f"Payslip_{emp.get('emp_id') or emp['eid']}_{month}.xlsx"
        return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name=fname)
    except Exception as exc:
        flash(f'Excel export failed: {exc}', 'error')
        return redirect(url_for('payroll_manual'))

if __name__ == '__main__':
    import sys
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 80
    try:
        from waitress import serve
        print(f"Starting waitress on port {port}...")
        serve(app, host='0.0.0.0', port=port, threads=4)
    except ImportError:
        print("waitress not found, falling back to Flask dev server")
        app.run(host='0.0.0.0', port=port, debug=False)
