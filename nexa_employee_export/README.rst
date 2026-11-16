====================
Nexa Employee Export
====================

Export the employees to Excel in one click, for Odoo 19.0 (Community or
Enterprise). Free and open source (LGPL-3).

* **Report**: one readable sheet with grouped headers (identity, job, contract,
  private information, emergency contact, bank, work permit) and computed
  columns: age, join date and length of service.
* **Your columns**: switch the groups on or off and add any other employee
  field, including the fields of your own modules.
* **Update and re-import**: a file in the import format of Odoo (external ID
  and technical names) with a reference sheet of the related records. Change
  the data in Excel and import it back: the same employees are updated.
* No salary is exported, and a user only gets the fields he can read.

Requirements
============

* Odoo 19.0 with the Employees application (``hr``).
* No Python package to install, no external service.

Installation
============

1. Copy the ``nexa_employee_export`` folder into your addons path.
2. Restart Odoo and update the apps list.
3. Install **Nexa Employee Export**.

Usage
=====

1. **Employees**, list view: select employees, or none for all the employees
   of the current companies.
2. Click **Export Excel**, choose the format and the columns, then **Export**.
3. To update many employees: choose **Update and Re-import**, change the file,
   then in the employee list use the gear icon, **Import records**, and upload
   the file. Keep the ``id`` column: it makes the import update the same
   employees.

The button is for the **Employees: Officer** role and above.

Good to know
============

* The module only adds a button and a wizard. The standard export and import
  of Odoo keep working as before.
* The file of an export can only be downloaded by the user who prepared it.

Support
=======

halo@nexacoresolution.id

License
=======

GNU Lesser General Public License v3 (LGPL-3). See the ``LICENSE`` file.

Copyright (c) 2026 NexaCore Solution.
