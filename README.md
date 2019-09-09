# Automated email report
Python script that extracts short interest data from a database, checks whether this has been updated on a daily base (weekdays) and subsequently compiles and sends an automated email report on Fridays

## Overview

- This repository contains a script that will run on weekdays and
  - Check whether data has been updated on weekdays and send an email if the data was not updated
  - Compiles and send an automated email report on Fridays based on the weekly difference in values
- The automated email report contains:
  - The top 15 companies with the highest short interest
  - The top 15 companies with the most days to cover
  - The top 15 weekly increases in short interest
  - The top 15 weekly decreases in short interest
  - Overview of companies that reported no data either this week or last week and thus have not been included in the report
- The scripts use SQL and the following (external) Python packages/dependencies, read the documentation for the correct use of the packages:
  - 
- When compiling the email report, the script will compile the report based on the weekly difference. It will use the latest available data in the database and compile a weekly report based on that date. For example, today is Friday 20 September, but the latest available data in the database is from Wednesday 18 September. The report will then compile the report between data from Wednesday 11 September and Wednesday 18 September. 
- A separate script was used to check for new data and write that data to the database on weekdays.
- The script was set to run through Windows Task Scheduler on weekdays and is designed to run remotely in a cloud environment

- Pseudocode for the entire setup on Monday - Thursday:
  - 15.00 start script
  - if data has been updated:
    - do nothing
  - else (data was not updated):
    - send email to request data to be updated
    - set timer for 30 minutes
  - 15.30 if data still has not been updated:
    - send email that data has not been updated
  - else (if data was updated):
    - do nothing
    
- Pseudocode for Friday:
  - 15.00 start script
  - if data has been updated:
    - compile and send automated email report
  - else: 
    - send email to request data to be updated
    - set timer for 30 minutes
  - 15.30:
    - compile and send automated email report with the latest available data, regardless of whether data has been updated
   
  
