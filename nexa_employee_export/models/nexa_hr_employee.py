from odoo import models


class HrEmployee(models.Model):
    _inherit = 'hr.employee'

    def action_nex_export_xlsx(self):
        """Export wizard of the selected employees (none: all employees)."""
        return self.env['nex.employee.export.wizard']._nex_open_action(self)
