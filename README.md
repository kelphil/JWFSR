# JWFSR Readme
Thanks for checking this out. Hope you will find it useful.

# Setup Instructions
1. Open field service report (FSR) template spreadsheet provided to you
2. Click 'File' -> 'Make a copy' to copy the FSR template Google spreadsheet; provide a meaningful filename 
3. Enter 'publisher name', 'email address', 'phone number' in the 'Email List' tab of your FSR spreadsheet. 'UserCode' will be generated automatically.
4. Enter 'congregation/group name' and 'language' on 'Controls' tab (row 14) of your FSR spreadsheet
5. Enter first/last email day and first email summary day on 'Controls' tab (row 15) of your FSR spreadsheet
6. Select Group Overseer name from the dropdown on 'Controls' tab (row 16, column B) of your FSR spreadsheet
7. Click 'FSR Menu' -> 'Create Form'
8. Click 'Continue' and then 'Allow' to authorize script access to your new Google spreadsheet
9. Click 'Form' -> 'Go to live form' to validate the field service report entry online form. Optional: Enter a sample field service report to validate the system.
10. If you entered sample data, Click 'FSR Menu' -> 'Update Dashboard'. Check if the sample report shows up on the 'Dashboard' tab. The 'Progress' tab should also reflect the sample data status.
11. If you entered sample data and if the 'Dashboard' reflects the new sample report you entered, click 'FSR Menu' -> 'Send FSR Summary Email' to send a sample summary email to the group overseer (Controls-B16)
12. Click FSR Menu -> Create Triggers to enable automated emails to publishers. Click 'Continue' and then 'Allow' to authorize script access to your new Google spreadsheet.
13. Wait for the 'First Email Day' (Controls-B15) of next month to start getting automated emails.

# Reminder Emails
A reminder email will be sent on alternate days to each publisher from the First Email Day (Controls-B15) to Last Email Day (Controls-D15) with his/her unique code (reset each month) and a link to the online field service report form. Further reminder emails will not be sent once a publisher submits his/her field service report for the month via the online form. 

Reminder emails are sent to publishers who are marked as YES in the FIELD SERVICE REPORT (Email List - ColumnD) of your FSR spreadsheet.

# Summary Emails
A field service group summary email will be sent to the every Group Overseer each day from First Email Summary Day (Controls-F15) to Last Email Day (Controls-D15). Assistant Group Overseer will be copied on these emails if Controls-F14 is marked as YES. These emails will show report grouped by Publishers/RP/AP etc

A congregation summary email will be sent to the Congregation Secretary (Controls-B17) on Congregation Report Email Summary Day (Controls-D17). All Congregation Elders will be copied on this email if Controls-F14 is marked as YES. This email will also show report statistics (totals/averages) grouped by Publishers/RP/AP etc.
