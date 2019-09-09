# Automated short interest email report
Python script that extracts [short interest](https://www.investopedia.com/terms/s/shortinterest.asp) data from a database, checks whether the data today has been updated on a daily base (weekdays) and subsequently compiles and sends an automated email report on Fridays containing a weekly update.

## Overview

- This repository contains a script that will run on weekdays and
  - Check whether data has been updated on weekdays and send an email if the data was not updated, requesting to update the data within 20 minutes.
  - Compiles and send an automated email report on Fridays based on the weekly differences in short interest data.
- The short interest data focuses on Dutch (AEX) and Belgian (BELMID29) exchange-listed companies.
- The automated email report contains:
  - The top 15 companies with the highest short interest.
  - The top 15 companies with the most days to cover.
  - The top 15 weekly increases in short interest.
  - The top 15 weekly decreases in short interest.
  - Overview of companies that reported no data either this week or last week and thus have not been included in the report.
- The script uses the following (external) Python packages/dependencies, read the documentation for the correct use of the packages:
  - [datetime](https://docs.python.org/3/library/datetime.html)
  - [email](https://docs.python.org/3/library/email.html)
  - [IPython](https://ipython.readthedocs.io/en/stable/)
  - [numpy](https://docs.scipy.org/)
  - [pandas](https://pandas.pydata.org/pandas-docs/stable/)
  - [pyodbc](https://github.com/mkleehammer/pyodbc/wiki)
  - [smtplib](https://docs.python.org/3/library/smtplib.html)
  - [statistics](https://docs.python.org/3/library/statistics.html)
  - [time](https://docs.python.org/2/library/time.html)
- When compiling the email report, the script will compile the report based on the weekly differences. It will use the latest available data in the database and compile a weekly report based on that date. For example, today is Friday 20 September, but the latest available data in the database is from Wednesday 18 September, the report will compile the report from data between Wednesday 11 September and Wednesday 18 September. 
- The [report.py](/report.py) file can be used to run the file in a cloud environment, the [report_manual_send.py](/report_manual_send.py) file can be used to run the script manually. 
- A separate script was used to check for new data and write that data to the database.
- The script should be set to run through Windows Task Scheduler on weekdays and is designed to run remotely in a cloud environment.
- Pseudocode for the entire setup on Monday - Thursday:
  - 15.00 start script
  - if data has been updated:
    - do nothing
  - else (data was not updated):
    - send email to request data to be updated within 20 minutes
    - set timer for 30 minutes
    - 15.25 external script checks whether new data was placed in folder, if yes write to database
  - 15.30 if data still has not been updated:
    - send email that data has not been updated
  - else (if data was updated):
    - do nothing
    
- Pseudocode for Friday:
  - 15.00 start script
  - if data has been updated:
    - compile and send automated email report
  - else: 
    - send email to request data to be updated within 20 minutes
    - set timer for 30 minutes
    - 15.25 external script checks whether new data was placed in folder, if yes write to database
  - 15.30:
    - compile and send automated email report with the latest available data, regardless of whether data has been updated
