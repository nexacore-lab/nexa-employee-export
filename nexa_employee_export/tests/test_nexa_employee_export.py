import io
from datetime import date, datetime

from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook

from odoo import fields
from odoo.exceptions import AccessError
from odoo.tests import HttpCase, TransactionCase, new_test_user, tagged


def give_contract(employee, start, end=False):
    """Contract of ``employee``: hr.contract (Odoo 18) or contract dates of the
    employee record (Odoo 19)."""
    env = employee.env
    if 'hr.contract' in env:
        return env['hr.contract'].create({
            'name': 'Contract %s' % employee.name, 'employee_id': employee.id,
            'date_start': start, 'date_end': end, 'wage': 1000, 'state': 'open',
        })
    version = employee.version_id
    version.write({'contract_date_start': start, 'contract_date_end': end})
    return version


class NexExportCommon(TransactionCase):

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        # English texts are checked, whatever the main language of the database
        cls.env = cls.env(context=dict(cls.env.context, tracking_disable=True, lang='en_US'))
        cls.company = cls.env.company
        cls.officer = new_test_user(cls.env, lang='en_US', login='nex_exp_officer', groups='hr.group_hr_user',
            company_id=cls.company.id, company_ids=[(6, 0, cls.company.ids)])
        cls.employee_user = new_test_user(cls.env, lang='en_US', login='nex_exp_employee', groups='base.group_user',
            company_id=cls.company.id, company_ids=[(6, 0, cls.company.ids)])
        cls.dept = cls.env['hr.department'].create({'name': 'NEX Export Dept'})
        Employee = cls.env['hr.employee']
        cls.emp_a = Employee.create({
            'name': 'NEX Export Alpha', 'department_id': cls.dept.id, 'job_title': 'Analyst',
            'birthday': date(1990, 5, 17), 'private_email': 'alpha@example.com',
        })
        cls.emp_b = Employee.create({'name': 'NEX Export Beta', 'department_id': cls.dept.id})
        cls.emp_archived = Employee.create({'name': 'NEX Export Gone', 'active': False})
        cls.today = fields.Date.context_today(Employee)
        give_contract(cls.emp_a, cls.today - relativedelta(years=3, months=2))

    def _wizard(self, user=None, **vals):
        return self.env['nex.employee.export.wizard'].with_user(user or self.officer).create(vals)

    def _sheet(self, wizard, name=None):
        content, _filename = wizard._nex_build_xlsx()
        book = load_workbook(io.BytesIO(content))
        return book[name] if name else book.worksheets[0], book


@tagged('post_install', '-at_install')
class TestNexaEmployeeExport(NexExportCommon):

    def _rows(self, sheet):
        return [list(row) for row in sheet.iter_rows(values_only=True)]

    def test_report_all_employees(self):
        wizard = self._wizard()
        self.assertIn('all the', wizard.summary)
        sheet, _book = self._sheet(wizard)
        rows = self._rows(sheet)
        self.assertIn('IDENTITY', rows[0])
        self.assertIn('JOB', rows[0])
        headers = rows[1]
        name_title = self.env['hr.employee']._fields['name']._description_string(self.env)
        self.assertEqual(headers[0], name_title)
        self.assertIn('Age', headers)
        self.assertIn('Length of Service', headers)
        names = [row[0] for row in rows[2:]]
        self.assertIn('NEX Export Alpha', names)
        self.assertIn('NEX Export Beta', names)
        self.assertNotIn('NEX Export Gone', names, 'archived employees only on request')
        self.assertNotIn('wage', [h.lower() for h in headers if h], 'no salary in the export')

        wizard.include_archived = True
        sheet, _book = self._sheet(wizard)
        self.assertIn('NEX Export Gone', [row[0] for row in self._rows(sheet)[2:]])

    def test_computed_columns(self):
        wizard = self._wizard(employee_ids=[(6, 0, self.emp_a.ids)])
        self.assertIn('1 selected', wizard.summary)
        sheet, _book = self._sheet(wizard)
        rows = self._rows(sheet)
        self.assertEqual(len(rows), 3, 'header rows and one employee')
        values = dict(zip(rows[1], rows[2]))
        self.assertEqual(values['Age'], relativedelta(self.today, date(1990, 5, 17)).years)
        self.assertEqual(values['Join Date'],
                         datetime.combine(self.today - relativedelta(years=3, months=2), datetime.min.time()))
        self.assertEqual(values['Length of Service'], '3 y 2 m')
        self.assertEqual(values[self.env['hr.employee']._fields['department_id']
                                ._description_string(self.env)], 'NEX Export Dept')

    def test_groups_and_additional_fields(self):
        Option = self.env['nex.employee.export.field']
        Option._nex_sync()
        pin = Option.search([('name', '=', 'pin')])
        self.assertEqual(len(pin), 1)
        wizard = self._wizard(
            employee_ids=[(6, 0, (self.emp_a | self.emp_b).ids)],
            group_identity=True, group_job=False, group_contract=False, group_private=False,
            group_emergency=False, group_bank=False, group_permit=False,
            extra_field_ids=[(6, 0, pin.ids)])
        rows = self._rows(self._sheet(wizard)[0])
        self.assertIn('IDENTITY', rows[0])
        self.assertNotIn('JOB', rows[0])
        self.assertIn('ADDITIONAL FIELDS', rows[0])
        pin_title = self.env['hr.employee']._fields['pin']._description_string(self.env)
        self.assertIn(pin_title, rows[1])
        self.assertEqual(len(rows), 4)

    def test_offered_fields(self):
        Option = self.env['nex.employee.export.field']
        Option._nex_sync()
        officer_option = Option.with_user(self.officer)
        names = officer_option.search(officer_option._nex_readable_domain()).mapped('name')
        self.assertIn('pin', names)
        self.assertFalse([n for n in names if 'wage' in n or 'salary' in n], 'never a salary')
        self.assertFalse([n for n in names if n.startswith('message_')])
        Employee = self.env['hr.employee']
        restricted = [n for n in names if Employee._fields[n].groups
                      and not self.officer.has_groups(Employee._fields[n].groups)]
        self.assertFalse(restricted, 'only the fields the officer can read')
        # A second sync changes nothing
        count = Option.search_count([])
        Option._nex_sync()
        self.assertEqual(Option.search_count([]), count)

    def test_reimport_updates_the_same_employees(self):
        wizard = self._wizard(mode='import', employee_ids=[(6, 0, (self.emp_a | self.emp_b).ids)])
        sheet, book = self._sheet(wizard, 'Employees')
        self.assertIn('Reference', book.sheetnames)
        rows = self._rows(sheet)
        header = rows[0]
        self.assertEqual(header[0], 'id')
        self.assertIn('department_id/id', header)
        self.assertIn('job_title', header)
        self.assertNotIn('birthday/id', header)

        def cell(value):
            if value is None:
                return ''
            if isinstance(value, datetime):
                return value.date().isoformat()
            return str(value)

        data = [[cell(v) for v in row] for row in rows[1:]]
        job_col = header.index('job_title')
        name_col = header.index('name')
        for row in data:
            if row[name_col] == 'NEX Export Alpha':
                row[job_col] = 'Senior Analyst'
        count = self.env['hr.employee'].search_count([])
        result = self.env['hr.employee'].with_user(self.officer).load(header, data)
        self.assertFalse([m for m in result['messages'] if m['type'] == 'error'], result['messages'])
        self.assertEqual(self.emp_a.job_title, 'Senior Analyst')
        self.assertEqual(self.emp_a.department_id, self.dept)
        self.assertEqual(self.env['hr.employee'].search_count([]), count, 'no new employee')

    def test_only_hr_officers(self):
        with self.assertRaises(AccessError):
            self._wizard(user=self.employee_user)
        arch = self.env['hr.employee'].with_user(self.officer).get_views([(False, 'list')])
        self.assertIn('action_nex_export_xlsx', arch['views']['list']['arch'])
        # The file of a wizard belongs to the user who prepared it
        other = new_test_user(self.env, lang='en_US', login='nex_exp_officer2', groups='hr.group_hr_user')
        wizard = self._wizard()
        with self.assertRaises(AccessError):
            wizard.with_user(other)._nex_build_xlsx()

    def test_open_from_list(self):
        action = (self.emp_a | self.emp_b).with_user(self.officer).action_nex_export_xlsx()
        self.assertEqual(action['res_model'], 'nex.employee.export.wizard')
        wizard = self.env['nex.employee.export.wizard'].with_user(self.officer).with_context(
            action['context']).create({})
        self.assertEqual(wizard.employee_ids, self.emp_a | self.emp_b)
        url = wizard.action_export()['url']
        self.assertEqual(url, '/nexa_employee_export/xlsx/%s' % wizard.id)

    def test_standard_export_untouched(self):
        data = (self.emp_a | self.emp_b).export_data(['name', 'department_id'])['datas']
        self.assertEqual(sorted(row[0] for row in data), ['NEX Export Alpha', 'NEX Export Beta'])


@tagged('post_install', '-at_install')
class TestNexaEmployeeExportHttp(HttpCase):

    def test_download(self):
        officer = new_test_user(self.env, lang='en_US', login='nex_exp_http', password='nex_exp_http',
                                groups='hr.group_hr_user')
        other = new_test_user(self.env, lang='en_US', login='nex_exp_http2', password='nex_exp_http2',
                              groups='hr.group_hr_user')
        self.env['hr.employee'].create({'name': 'NEX Export Http'})
        wizard = self.env['nex.employee.export.wizard'].with_user(officer).create({})
        self.authenticate('nex_exp_http', 'nex_exp_http')
        response = self.url_open('/nexa_employee_export/xlsx/%s' % wizard.id)
        self.assertEqual(response.status_code, 200)
        self.assertIn('spreadsheetml', response.headers['Content-Type'])
        self.assertTrue(response.content.startswith(b'PK'))
        self.authenticate('nex_exp_http2', 'nex_exp_http2')
        response = self.url_open('/nexa_employee_export/xlsx/%s' % wizard.id)
        self.assertNotEqual(response.status_code, 200, 'the file of another user is not served')
        self.assertTrue(other)
