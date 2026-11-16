from odoo import api, fields, models

# Field types that cannot be exported as one cell
SKIPPED_TYPES = ('binary', 'one2many', 'properties', 'properties_definition', 'json', 'reference',
                 'many2one_reference')
# Technical fields never offered
SKIPPED_PREFIXES = ('message_', 'activity_', 'website_message', 'rating_', 'my_activity_')
SKIPPED_NAMES = ('id', 'create_uid', 'write_uid', 'create_date', 'write_date', 'display_name')
# No salary in the export, whatever the rights of the user
SALARY_WORDS = ('wage', 'salary')


class NexEmployeeExportField(models.Model):
    """Fields of the employee that can be added to the export.

    The Fields model of Odoo is reserved to the administrators; this list lets
    the HR Officers choose additional columns. It is refreshed from the
    employee model each time the export opens."""

    _name = 'nex.employee.export.field'
    _description = 'Exportable Employee Field'
    _order = 'label, name'
    _rec_names_search = ['label', 'name']

    name = fields.Char('Technical Name', required=True, readonly=True)
    label = fields.Char('Label', readonly=True)

    @api.depends('name', 'label')
    @api.depends_context('lang')
    def _compute_display_name(self):
        Employee = self.env['hr.employee']
        for option in self:
            field = Employee._fields.get(option.name)
            label = field._description_string(self.env) if field else option.label or option.name
            option.display_name = '%s (%s)' % (label, option.name)

    @api.model
    def _nex_is_exportable(self, name, field):
        if field.type in SKIPPED_TYPES or name in SKIPPED_NAMES or name.startswith(SKIPPED_PREFIXES):
            return False
        if any(word in name for word in SALARY_WORDS):
            return False
        # Stored fields, and the fields of the employee record (Odoo 19)
        return bool(field.store or getattr(field, 'inherited', False))

    @api.model
    def _nex_sync(self):
        """Create or remove the options to match the fields of the employee."""
        Option = self.sudo()
        wanted = {
            name: field for name, field in self.env['hr.employee']._fields.items()
            if self._nex_is_exportable(name, field)
        }
        existing = {option.name: option for option in Option.search([])}
        stale = Option.browse([option.id for name, option in existing.items() if name not in wanted])
        if stale:
            stale.unlink()
        missing = [name for name in wanted if name not in existing]
        if missing:
            Option.create([{'name': name, 'label': wanted[name].string} for name in missing])

    @api.model
    def _nex_readable_domain(self):
        """Options of the fields that the current user can read."""
        Employee = self.env['hr.employee']
        user = self.env.user
        names = [
            name for name, field in Employee._fields.items()
            if self._nex_is_exportable(name, field) and (not field.groups or user.has_groups(field.groups))
        ]
        return [('name', 'in', names)]
