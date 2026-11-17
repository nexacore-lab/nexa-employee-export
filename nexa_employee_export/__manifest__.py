{
    'name': 'Employee Excel Export | HR Report, Bulk Update & Re-import',
    'version': '19.0.1.0.0',
    'summary': 'Employee export, export employees to Excel, employee Excel export, employee data export, '
               'employee report Excel, employee list export, HR export, export employee data, employee '
               'master data, employee database export, employee import, update employees from Excel, '
               'mass update employees, bulk update employees, re-import employees, HR Excel report, '
               'export karyawan, data karyawan Excel, HRIS, free',
    'description': """
Nexa Employee Export
====================
An "Export Excel" button on the employee list:

* report mode: one readable Excel file with grouped headers (identity, job,
  contract, private address, emergency contact, bank, work permit), computed
  columns such as age, join date and length of service, and any other employee
  field you pick;
* re-import mode: a file in the import format of Odoo (external ID and
  technical names), with a reference sheet of the related records, so HR
  updates many employees in Excel and imports the file back with the standard
  Import of Odoo.

The export follows the selected rows, or all the employees of the current
companies. No salary is exported. The module adds a button and a wizard; the
standard Employees application and its export keep working as before. It uses
no external service.
""",
    'category': 'Human Resources',
    'author': 'NexaCore Solution',
    'maintainer': 'NexaCore Solution',
    'website': 'https://nexacoresolution.id',
    'support': 'halo@nexacoresolution.id',
    'license': 'LGPL-3',
    'images': ['images/main_screenshot.png'],
    'depends': [
        'hr',
    ],
    'data': [
        'security/ir.model.access.csv',
        'wizard/nexa_employee_export_wizard_views.xml',
        'views/nexa_hr_employee_views.xml',
    ],
    'installable': True,
    'application': True,
}
