from odoo import http
from odoo.exceptions import AccessError
from odoo.http import content_disposition, request


class NexaEmployeeExport(http.Controller):

    @http.route('/nexa_employee_export/xlsx/<int:wizard_id>', type='http', auth='user', methods=['GET'])
    def nex_export_xlsx(self, wizard_id, **kwargs):
        wizard = request.env['nex.employee.export.wizard'].browse(wizard_id).exists()
        try:
            content, filename = wizard._nex_build_xlsx() if wizard else (None, None)
        except AccessError:
            content = None
        if not content:
            return request.not_found()
        return request.make_response(content, headers=[
            ('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
            ('Content-Disposition', content_disposition(filename)),
        ])
