import io

import xlsxwriter
from dateutil.relativedelta import relativedelta

from odoo import _, api, fields, models
from odoo.exceptions import AccessError, UserError

from ..models.nexa_employee_export_field import SALARY_WORDS, SKIPPED_TYPES

# Columns of the report, per group: (path, width). A path is a field of the
# employee, a field of a related record ('user_id.login') or a computed column
# ('@age'). A column whose field does not exist in the database (other Odoo
# version, module not installed) is left out, so one list serves Odoo 18 and 19.
COLUMNS = [
    ('identity', [
        # gender: Odoo 18, sex: Odoo 19
        ('name', 26), ('legal_name', 26), ('gender', 10), ('sex', 10), ('birthday', 12), ('@age', 7),
        ('place_of_birth', 15), ('country_id', 15), ('identification_id', 18), ('ssnid', 16),
        ('passport_id', 13), ('marital', 12), ('children', 8), ('spouse_complete_name', 22),
        ('spouse_birthdate', 12),
    ]),
    ('job', [
        ('company_id', 24), ('department_id', 20), ('job_id', 22), ('job_title', 24),
        ('employee_type', 12), ('parent_id', 22), ('coach_id', 22), ('address_id', 24),
        ('work_location_id', 16), ('resource_calendar_id', 20), ('leave_manager_id', 22),
        ('user_id.login', 26), ('category_ids', 20), ('barcode', 12), ('work_email', 28),
        ('work_phone', 15), ('mobile_phone', 16), ('active', 8),
    ]),
    ('contract', [
        # Odoo 18: current contract (hr_contract)
        ('contract_id.name', 26), ('contract_id.contract_type_id', 14),
        ('contract_id.date_start', 12), ('contract_id.date_end', 12), ('contract_id.state', 11),
        # Odoo 19: contract dates of the employee record
        ('contract_type_id', 14), ('contract_date_start', 12), ('contract_date_end', 12),
        ('@join_date', 12), ('@service', 13), ('departure_date', 12), ('departure_reason_id', 16),
    ]),
    ('private', [
        ('private_email', 26), ('private_phone', 16), ('private_street', 30), ('private_street2', 20),
        ('private_city', 14), ('private_state_id', 16), ('private_zip', 9), ('private_country_id', 13),
        ('lang', 12), ('certificate', 12), ('study_field', 18), ('study_school', 22),
        ('km_home_work', 10),
    ]),
    ('emergency', [
        ('emergency_contact', 22), ('emergency_phone', 16),
    ]),
    ('bank', [
        # Odoo 18: one bank account, Odoo 19: several
        ('bank_account_id.acc_number', 20), ('bank_account_id.bank_id', 16), ('bank_account_ids', 26),
    ]),
    ('permit', [
        ('visa_no', 13), ('visa_expire', 12), ('permit_no', 14), ('work_permit_expiration_date', 12),
    ]),
]
GROUP_COLORS = ['#1F3864', '#2F5496']


class NexEmployeeExportWizard(models.TransientModel):
    _name = 'nex.employee.export.wizard'
    _description = 'Employee Excel Export'

    mode = fields.Selection([
        ('report', 'Report'),
        ('import', 'Update and Re-import'),
    ], string='Format', required=True, default='report')
    employee_ids = fields.Many2many(
        'hr.employee', 'nex_employee_export_wizard_employee_rel', 'wizard_id', 'employee_id',
        string='Employees', context={'active_test': False})
    include_archived = fields.Boolean(
        'Include Archived Employees',
        help='Only used when no employee is selected in the list.')
    group_identity = fields.Boolean('Identity', default=True)
    group_job = fields.Boolean('Job', default=True)
    group_contract = fields.Boolean('Contract', default=True)
    group_private = fields.Boolean('Private Information', default=True)
    group_emergency = fields.Boolean('Emergency Contact', default=True)
    group_bank = fields.Boolean('Bank', default=True)
    group_permit = fields.Boolean('Work Permit', default=True)
    extra_field_ids = fields.Many2many(
        'nex.employee.export.field', 'nex_employee_export_wizard_field_rel', 'wizard_id', 'field_id',
        string='Additional Fields', domain=lambda self: self.env['nex.employee.export.field']._nex_readable_domain(),
        help='Other fields of the employee to add at the end, for example the fields of your own '
             'modules.')
    summary = fields.Char(compute='_compute_summary')

    # ------------------------------------------------------------------
    # Wizard
    # ------------------------------------------------------------------

    @api.model
    def default_get(self, fields_list):
        self.env['nex.employee.export.field']._nex_sync()
        return super().default_get(fields_list)

    @api.model
    def _nex_open_action(self, employees):
        return {
            'type': 'ir.actions.act_window',
            'name': _('Export Employees'),
            'res_model': self._name,
            'view_mode': 'form',
            'views': [(False, 'form')],
            'target': 'new',
            'context': {'default_employee_ids': [(6, 0, employees.ids)]},
        }

    @api.depends('employee_ids', 'include_archived')
    def _compute_summary(self):
        for wizard in self:
            count = len(wizard._nex_employees())
            if wizard.employee_ids:
                wizard.summary = _('%s selected employees.', count)
            else:
                wizard.summary = _('No row selected: all the %s employees of the current companies.', count)

    def _nex_employees(self):
        self.ensure_one()
        selected = self.with_context(active_test=False).employee_ids
        if selected:
            employees = selected
        else:
            domain = [('company_id', 'in', self.env.companies.ids)]
            if not self.include_archived:
                domain.append(('active', '=', True))
            employees = self.env['hr.employee'].with_context(active_test=False).search(domain)
        return employees.sorted(
            lambda e: (e.company_id.name or '', e.department_id.name or '', e.name or ''))

    def action_export(self):
        self.ensure_one()
        if not self._nex_employees():
            raise UserError(_('There is no employee to export.'))
        return {
            'type': 'ir.actions.act_url',
            'url': '/nexa_employee_export/xlsx/%s' % self.id,
            'target': 'download',
        }

    # ------------------------------------------------------------------
    # Columns
    # ------------------------------------------------------------------

    def _nex_group_titles(self):
        return {
            'identity': _('Identity'),
            'job': _('Job'),
            'contract': _('Contract'),
            'private': _('Private Information'),
            'emergency': _('Emergency Contact'),
            'bank': _('Bank'),
            'permit': _('Work Permit'),
            'extra': _('Additional Fields'),
        }

    def _nex_computed_titles(self):
        return {
            '@age': _('Age'),
            '@join_date': _('Join Date'),
            '@service': _('Length of Service'),
        }

    def _nex_path_fields(self, path):
        """Fields of the path, or None when a field does not exist (other Odoo
        version, module not installed), is a salary or cannot be read by the user."""
        model = self.env['hr.employee']
        user = self.env.user
        result = []
        for name in path.split('.'):
            field = model._fields.get(name)
            if not field or field.type in SKIPPED_TYPES or any(word in name for word in SALARY_WORDS):
                return None
            if field.groups and not user.has_groups(field.groups):
                return None
            result.append(field)
            if field.relational:
                model = self.env[field.comodel_name]
        return result

    def _nex_title(self, path):
        if path.startswith('@'):
            return self._nex_computed_titles()[path]
        Employee = self.env['hr.employee']
        names = path.split('.')
        title = Employee._fields[names[0]]._description_string(self.env)
        if len(names) > 1:
            sub = self.env[Employee._fields[names[0]].comodel_name]._fields[names[1]]
            title = '%s - %s' % (title, sub._description_string(self.env))
        return title

    def _nex_columns(self):
        """[(group key, [(title, path, width)])] of the chosen groups."""
        self.ensure_one()
        groups, used = [], set()
        for key, columns in COLUMNS:
            if not self['group_%s' % key]:
                continue
            cols = []
            for path, width in columns:
                if path.startswith('@') or self._nex_path_fields(path):
                    cols.append((self._nex_title(path), path, width))
                    used.add(path)
            if cols:
                groups.append((key, cols))
        extra = [
            (self._nex_title(option.name), option.name, 18)
            for option in self.extra_field_ids
            if option.name not in used and self._nex_path_fields(option.name)
        ]
        if extra:
            groups.append(('extra', extra))
        return groups

    # ------------------------------------------------------------------
    # Values
    # ------------------------------------------------------------------

    @api.model
    def _nex_join_date(self, employee):
        if 'first_contract_date' in employee._fields:
            return employee.first_contract_date or None
        if hasattr(employee, '_get_first_contract_date'):
            return employee._get_first_contract_date() or None
        return None

    def _nex_value(self, employee, path, today):
        if path == '@age':
            return relativedelta(today, employee.birthday).years if employee.birthday else None
        if path == '@join_date':
            return self._nex_join_date(employee)
        if path == '@service':
            start = self._nex_join_date(employee)
            if not start:
                return None
            end = employee.departure_date if not employee.active and employee.departure_date else today
            delta = relativedelta(end, start)
            return _('%(years)s y %(months)s m', years=delta.years, months=delta.months)

        record = employee
        *parents, name = path.split('.')
        for part in parents:
            record = record[part][:1]
            if not record:
                return None
        field = record._fields[name]
        value = record[name]
        if field.type == 'many2one':
            return value.display_name or None
        if field.type == 'many2many':
            return ', '.join(value.mapped('display_name')) or None
        if field.type == 'selection':
            return dict(field._description_selection(record.env)).get(value) if value else None
        if field.type == 'boolean':
            return _('Yes') if value else _('No')
        if field.type in ('integer', 'float', 'monetary'):
            return value
        return value or None

    # ------------------------------------------------------------------
    # Excel file
    # ------------------------------------------------------------------

    def _nex_check_owner(self):
        """Only the user who prepared the export downloads the file."""
        self.ensure_one()
        if not self.env.su and self.create_uid != self.env.user:
            raise AccessError(_('This export was prepared by another user.'))

    def _nex_build_xlsx(self):
        """(content, file name) of the Excel file."""
        self.ensure_one()
        self._nex_check_owner()
        employees = self._nex_employees()
        # Read with sudo: the employees were selected with the rights of the user
        # and the columns are limited to the fields this user can read; sudo
        # only gives access to the related records (contract, bank account).
        employees = employees.sudo()
        today = fields.Date.context_today(self)
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        if self.mode == 'import':
            self._nex_write_import(workbook, employees)
            filename = _('Employees (re-import) %s.xlsx', today)
        else:
            self._nex_write_report(workbook, employees, today)
            filename = _('Employees %s.xlsx', today)
        workbook.close()
        return output.getvalue(), filename

    def _nex_write_report(self, workbook, employees, today):
        sheet = workbook.add_worksheet(_('Employees'))
        groups = self._nex_columns()
        titles = self._nex_group_titles()
        base = {'font_name': 'Arial', 'font_size': 9, 'valign': 'vcenter', 'border': 1,
                'border_color': '#BFBFBF'}
        band = ['#FFFFFF', '#F2F5F9']
        fmt_text = [workbook.add_format(dict(base, bg_color=c)) for c in band]
        fmt_date = [workbook.add_format(dict(base, bg_color=c, num_format='dd mmm yyyy', align='left'))
                    for c in band]
        fmt_num = [workbook.add_format(dict(base, bg_color=c, align='center')) for c in band]

        col = 0
        paths = []
        for index, (key, cols) in enumerate(groups):
            color = GROUP_COLORS[index % 2]
            fmt_group = workbook.add_format(dict(base, bold=True, font_size=10, font_color='#FFFFFF',
                                                 bg_color=color, align='center', border_color='#FFFFFF'))
            fmt_head = workbook.add_format(dict(base, bold=True, font_color='#FFFFFF', bg_color=color,
                                                align='center', text_wrap=True, border_color='#FFFFFF'))
            if len(cols) > 1:
                sheet.merge_range(0, col, 0, col + len(cols) - 1, titles[key].upper(), fmt_group)
            else:
                sheet.write(0, col, titles[key].upper(), fmt_group)
            for title, path, width in cols:
                sheet.set_column(col, col, width)
                sheet.write(1, col, title, fmt_head)
                paths.append(path)
                col += 1
        sheet.set_row(0, 18)
        sheet.set_row(1, 32)
        sheet.freeze_panes(2, 1)

        for index, employee in enumerate(employees):
            row, shade = index + 2, index % 2
            for c, path in enumerate(paths):
                value = self._nex_value(employee, path, today)
                if value is None or value == '':
                    sheet.write_blank(row, c, None, fmt_text[shade])
                elif hasattr(value, 'year'):
                    sheet.write_datetime(row, c, fields.Datetime.to_datetime(value), fmt_date[shade])
                elif isinstance(value, (int, float)) and not isinstance(value, bool):
                    sheet.write_number(row, c, value, fmt_num[shade])
                else:
                    sheet.write_string(row, c, str(value), fmt_text[shade])
        if paths:
            sheet.autofilter(1, 0, max(len(employees), 1) + 1, len(paths) - 1)

    # ------------------------------------------------------------------
    # Re-import file
    # ------------------------------------------------------------------

    def _nex_import_fields(self):
        """Simple fields of the chosen columns that the import of Odoo can write."""
        Employee = self.env['hr.employee']
        names = []
        for _key, cols in self._nex_columns():
            for _title, path, _width in cols:
                if path.startswith('@') or '.' in path or path in names:
                    continue
                field = Employee._fields[path]
                if field.type in SKIPPED_TYPES + ('many2many',):
                    continue
                # Fields of the employee record (Odoo 19) are written through the employee
                if (field.store or getattr(field, 'inherited', False)) and not field.readonly:
                    names.append(path)
        return names

    def _nex_write_import(self, workbook, employees):
        names = self._nex_import_fields()
        m2o = [name for name in names if employees._fields[name].type == 'many2one']
        # Relations as external IDs: no ambiguity with duplicated or archived names
        paths = [['id']] + [[name, 'id'] if name in m2o else [name] for name in names]
        rows = employees.with_context(import_compat=True)._export_rows(paths)

        sheet = workbook.add_worksheet(_('Employees'))
        base = {'font_name': 'Arial', 'font_size': 9, 'border': 1, 'border_color': '#BFBFBF'}
        fmt_head = workbook.add_format(dict(base, bold=True, font_color='#FFFFFF', bg_color='#1F3864'))
        fmt_id = workbook.add_format(dict(base, bg_color='#FFF2CC'))
        fmt_cell = workbook.add_format(base)
        fmt_date = workbook.add_format(dict(base, num_format='yyyy-mm-dd', align='left'))

        for col, path in enumerate(paths):
            header = '/'.join(path)
            if path == ['id']:
                label = _('External ID of the employee: do not change it.')
            else:
                label = employees._fields[path[0]]._description_string(self.env)
                if len(path) > 1:
                    label += ' - ' + _('external ID, see the Reference sheet')
            sheet.write(0, col, header, fmt_head)
            sheet.write_comment(0, col, label)
            sheet.set_column(col, col, max(len(header), 12) + 4)
        sheet.freeze_panes(1, 1)

        for r, row in enumerate(rows, start=1):
            for col, value in enumerate(row):
                cell = fmt_id if col == 0 else fmt_cell
                if isinstance(value, bool):
                    sheet.write_boolean(r, col, value, cell)
                elif value is None or value == '':
                    sheet.write_blank(r, col, None, cell)
                elif hasattr(value, 'year'):
                    sheet.write_datetime(r, col, fields.Datetime.to_datetime(value), fmt_date)
                elif isinstance(value, (int, float)):
                    sheet.write_number(r, col, value, cell)
                else:
                    sheet.write_string(r, col, str(value), cell)

        self._nex_write_reference(workbook, employees, m2o)

        help_sheet = workbook.add_worksheet(_('How to Import'))
        help_sheet.set_column(0, 0, 120)
        for r, line in enumerate([
            _('How to update the employees with this file:'),
            _('1. Change the data in the "Employees" sheet. Do not change or delete the "id" column '
              'and the header row.'),
            _('2. In Odoo, open Employees, list view, gear icon, Import records, and upload this file.'),
            _('3. Choose the "Employees" sheet: the columns are mapped by themselves, the headers are '
              'the technical names of the fields.'),
            _('4. Click Test, then Import. A row with the same "id" updates that employee.'),
            '',
            _('The columns ending with "/id" (department, job position, manager...) hold external '
              'IDs, not names. To change one, look for the new value in the "Reference" sheet and '
              'copy its external ID.'),
        ]):
            help_sheet.write(r, 0, line)

    def _nex_write_reference(self, workbook, employees, m2o):
        """External ID and name of the related records. Small tables are listed
        entirely, large ones (partners, states) only with the records used."""
        full_limit = 500
        by_model = {}
        for name in m2o:
            field = employees._fields[name]
            entry = by_model.setdefault(field.comodel_name, {
                'labels': [], 'records': self.env[field.comodel_name].sudo()})
            entry['labels'].append('%s (%s/id)' % (field._description_string(self.env), name))
            entry['records'] |= employees.mapped(name)

        sheet = workbook.add_worksheet(_('Reference'))
        base = {'font_name': 'Arial', 'font_size': 9, 'border': 1, 'border_color': '#BFBFBF'}
        fmt_head = workbook.add_format(dict(base, bold=True, font_color='#FFFFFF', bg_color='#1F3864'))
        fmt_cell = workbook.add_format(base)
        for col, (title, width) in enumerate([
            (_('Used in the columns'), 46), (_('External ID'), 42), (_('Name'), 46), (_('Archived'), 10),
        ]):
            sheet.write(0, col, title, fmt_head)
            sheet.set_column(col, col, width)
        sheet.freeze_panes(1, 0)

        r = 1
        for model_name, entry in by_model.items():
            Model = self.env[model_name].sudo().with_context(active_test=False)
            records = entry['records']
            if Model.search_count([]) <= full_limit:
                records |= Model.search([])
            records = records.sorted(lambda rec: rec.display_name or '')
            data = records.with_context(import_compat=True)._export_rows([['id'], ['display_name']])
            has_active = 'active' in Model._fields
            used_in = ', '.join(entry['labels'])
            for record, (xid, name) in zip(records, data):
                sheet.write_string(r, 0, used_in, fmt_cell)
                sheet.write_string(r, 1, xid or '', fmt_cell)
                sheet.write_string(r, 2, name or '', fmt_cell)
                sheet.write_string(r, 3, _('Yes') if has_active and not record.active else '', fmt_cell)
                r += 1
        sheet.autofilter(0, 0, max(r - 1, 1), 3)
