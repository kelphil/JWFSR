# JWFSR Readme
Thanks for checking this out. Hope you will find it useful.

# Setup Instructions
1. Open field service report (FSR) template spreadsheet provided to you
2. Click 'File' -> 'Make a copy' to copy the FSR template Google spreadsheet; provide a meaningful filename 
3. Enter 'publisher name', 'email address', 'phone number',  etc in the 'Email List' tab of your FSR spreadsheet
4. Enter 'congregation/group name' and 'language' on 'Controls' tab (row 14) of your FSR spreadsheet
5. Enter first/last email day and first email summary day on 'Controls' tab (row 15) of your FSR spreadsheet
6. Select Group Overseer name from the dropdown on 'Controls' tab (row 16, column B) of your FSR spreadsheet
7. Click 'FSR Menu' -> 'Create Form'
8. Click 'Continue' and then 'Allow' to authorize script access to your new Google spreadsheet
9. Click 'Form' -> 'Go to live form' to validate the field service report entry online form. Optional: Enter a sample field service report to validate the system.
10. If you entered sample data, Click 'FSR Menu' -> 'Update Dashboard'. Check if the sample report shows up on the 'Dashboard' tab. The 'Progress' tab should also reflect the sample data status.
11. Click FSR Menu -> Create Triggers to enable automated emails to publishers. Click 'Continue' and then 'Allow' to authorize script access to your new Google spreadsheet.

# Reminder Emails
A reminder email will be sent on alternate days to each publisher from the First Email Day (Controls-B15) to Last Email Day (Controls-D15) with his/her unique code (reset each month) and a link to the online field service report form. Further reminder emails will not be sent once a publisher submits his/her field service report for the month via the online form. 

Reminder emails are sent to publishers who are marked as YES in the FIELD SERVICE REPORT (Email List - ColumnF) of your FSR spreadsheet.

# Summary Emails
A summary email will be sent to the Group Overseer (Controls-B16) each day from First Email Summary Day (Controls-F15) to Last Email Day (Controls-D15).
