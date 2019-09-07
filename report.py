##############################
#                            #
#   Automated email report   #
#          oliverk1          #
#         July  2019         #
#                            #
##############################

# Import packages
import pyodbc, smtplib, time
import pandas.io.sql
import pandas as pd
import numpy as np
from datetime import date, timedelta, datetime
from IPython.display import display, HTML
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from statistics import mean 
from email import encoders

########################################
#  Setup for SQL and email connection  #
########################################

# Connect to database and select data from the past 16 weeks
conn = pyodbc.connect(
    "Driver={SQL Server Native Client 11.0};"
    "Server={***};"
    "Database=***;"
    "Trusted_Connection=yes;"
)

sql = """
SELECT *
FROM dbo.SI_DATA
WHERE Date >= DATEADD(wk,-16, getdate())
ORDER BY UpdateTimeStamp DESC
"""

# Set up email settings
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')

# Convert database SQL query to pandas
datapoints = pandas.io.sql.read_sql(sql, conn)

# Set up function
def runscript(datapoints):

    ###############################################
    #  Merge databases and obtain table per date  #
    ###############################################

    # Load column values and tickers
    datatypes = pandas.io.sql.read_sql("""SELECT * FROM dbo.SI_DataType""", conn)
    tickerids = pandas.io.sql.read_sql("""SELECT * FROM dbo.SI_TickerInfo""", conn)

    # Merge other tables with query results and select relevant columns
    new = pd.merge(datapoints, datatypes, how='left', on='DataTypeId')
    new = pd.merge(new, tickerids, how='left', on='TickerId')
    new = new.sort_values(['Date','TickerId','DataType'], ascending=False)
    new = new[['Id','TickerId','Value','Date','UpdateTimeStamp','DataType','IsPercentage','Ticker']]

    # Add full names to dataframe
    fullname = pd.read_csv("fullname.csv", sep=";", names=['Ticker', 'FullName'], skiprows=[0])
    new = pd.merge(new, fullname, how='left', on='Ticker')

    # Apply pivot table to obtain table per date
    newdf = pd.pivot_table(new, values='Value', index=['UpdateTimeStamp','TickerId','Ticker','FullName'], columns=['DataType'])
    newdf = newdf.sort_values('UpdateTimeStamp', ascending=False)
    newdf.reset_index(inplace=True)
    newdf.index.names = ['Index']

    #############################
    #  Access last week's data  #
    #############################

    # Function to obtain ordered set of dates
    def datesinorder(seq):
        seen = set()
        seen_add = seen.add
        return [x for x in seq if not (x in seen or seen_add(x))]

    # Obtain latest date in database and subsequent data
    latestdate = datesinorder(newdf.UpdateTimeStamp)[0]
    latestdatedf = newdf.loc[newdf['UpdateTimeStamp'] == latestdate].sort_values('Ticker')
    latestdatedf.reset_index(inplace=True)
    latestdatedf = latestdatedf.drop('Index', axis=1)

    # NOTE: Function that will find date for 1 week ago. If no date is found, the date will increment until found
    # For example if no data was found for last Wednesday, it will take the data from Tuesday, else Monday etc. etc.
    # The function returns the 'best' date in a datetime format

    def findlastdate():
        a = True
        datapointsdate = list(datapoints.UpdateTimeStamp.apply(lambda x: x.date()))
        days = 7
        while a == True:
            lastweek = (latestdate - timedelta(days)).date()
            if lastweek in datapointsdate:
                a = False
                return(lastweek)
            else:
                days+=1
    lastweekdate = findlastdate()

    # Obtain data from last week
    newdf['UpdateTimeStamp'] = newdf['UpdateTimeStamp'].apply(lambda x: x.date())
    lastweekdatedf = newdf.loc[newdf['UpdateTimeStamp'] == lastweekdate].sort_values('Ticker')
    lastweekdatedf.reset_index(inplace=True)
    lastweekdatedf = lastweekdatedf.drop('Index', axis=1)

    #####################
    #  Compute metrics  #
    #####################

    # Top 15 shorts % of free float
    def topshorts():
        topshorts = latestdatedf[['Ticker', 'FullName', 'SIPct']]
        topshorts = topshorts.sort_values('SIPct', ascending=False).head(15)
        topshorts['SIPct'] = topshorts['SIPct'].apply(lambda x: x * 100)

        html_topshorts = ""        
        for i in range(0,len(topshorts)):
            line = str(i+1) + '. ' + topshorts.FullName.iloc[i] + " <b>" + str(round(topshorts.SIPct.iloc[i],1)) + "%</b>"
            html_topshorts += line + "<br>"

        return(html_topshorts)

    # Top 15 short days to cover
    def daystocover():
        daystocover = latestdatedf[['Ticker', 'FullName', 'SIPct', 'DTC']]
        daystocover = daystocover.sort_values('DTC', ascending=False).head(15)
        daystocover['SIPct'] = daystocover['SIPct'].apply(lambda x: round(x * 100,1))
        daystocover['DTC'] = daystocover['DTC'].apply(lambda x: int(round(x,1)))

        html_daystocover = ""
        for i in range(0,len(daystocover)):
            line = str(i+1)+ '. ' + daystocover.FullName.iloc[i] + " " + "(SI " + str(daystocover.SIPct.iloc[i]) +"%) <b>" + str(daystocover.DTC.iloc[i])+ " days</b> to cover"
            html_daystocover += line + "<br>"

        return(html_daystocover)

    # Top 10 weekly increases and decreases
    def percentages():

        # Calculate percentual differences 
        nons = []
        names = []
        percs = []
        latestsi = []
        no_data = []

        for item in lastweekdatedf.Ticker:

            # If all values in row are empty, data is considered as missing
            if mean([item for item in lastweekdatedf.loc[lastweekdatedf['Ticker'] == item][['SIShares','SINotional','SIPct','SIDaily','SIWeekly','DTC','DTCDaily','DTCWeekly']].iloc[0]]) == 0.0:
                no_data.append(item)
            else:
                # Obtain this week's SI, store ticker if data is empty
                try:
                    latestrow = (latestdatedf.loc[latestdatedf['Ticker'] == item]['SIPct'].iloc[0]) * 100
                except:
                    no_data.append(item)
                    latestrow = ""

                # Obtain last week's SI, store ticker if data is empty
                try:
                    lastweekrow = (lastweekdatedf.loc[lastweekdatedf['Ticker'] == item]['SIPct'].iloc[0]) * 100
                except:
                    no_data.append(item)
                    lastweekrow = ""   

                # If both rows have data append the data
                if latestrow != "" and lastweekrow != "":
                    diff = (latestrow - lastweekrow)
                    names.append(item)
                    percs.append(diff)
                    latestsi.append(latestrow)    

        # Find percentual differences
        result = pd.DataFrame({'Ticker': names, 'Percent': percs, 'SI': latestsi})
        result = pd.merge(result, fullname, how='left', on='Ticker')
        topincrease = result.sort_values('Percent', ascending=False).head(15)
        topdecrease = result.sort_values('Percent', ascending=True).head(15)

        no_data = pd.DataFrame({'Ticker': no_data})
        no_data = pd.merge(no_data, fullname, how='left', on='Ticker')
        no_data = ", ".join(list(no_data['FullName']))
        
        # Calculate top 15 weekly increases
        html_topincr = ""
        for i in range(0,len(topincrease)):
            line = str(i+1)+ '. ' + topincrease.FullName.iloc[i] + " <b>+" + str(round(topincrease.Percent.iloc[i],1))+"%</b> (SI " + str(round(topincrease.SI.iloc[i],1)) +"%)"
            html_topincr += line + "<br>"

        # Calculate top 15 weekly decreases
        html_topdecr = ""
        for i in range(0,len(topdecrease)):
            line = str(i+1)+ '. ' + topdecrease.FullName.iloc[i] + " <b>" + str(round(topdecrease.Percent.iloc[i],1))+"%</b> (SI " + str(round(topdecrease.SI.iloc[i],1)) +"%)"
            html_topdecr += line + "<br>"

        return(html_topincr, html_topdecr, no_data)

    ###################
    #  Compose email  #
    ###################
    
    sender_email = "example@example.com"
    receiver_email = ['example@example.com']

    message = MIMEMultipart("alternative")
    message["Subject"] = "Weekly short interest report"
    message["From"] = "Sender name you want to appear in email header"
    message["To"] = ", ".join(receiver_email)

    html = """\
    <html>
    <p>Dear colleague,</p> <p>Please find the highlights of this week's automated short interest report below and the complete report in the attachment.</p> <p>The data has been calculated between <em>""" + str(lastweekdate.strftime('%A %d %B %Y')) + """</em> and <em>""" + str(latestdate.strftime('%A %d %B %Y')) + """</em> (data has a 1-day delay).</p> <p>Please note that if no data was present for one week ago, the tool will find the nearest historical date that has data available.</p> <h4>Top 15 shorts % of free float</h4> <p>""" + str(topshorts()) + """</p> <h4>Top 15 shorts days to cover</h4> <p>""" + str(daystocover()) + """</p> <h4>Top 15 weekly increases</h4> <p>""" + str(percentages()[0]) + """</p> <h4>Top 15 weekly decreases</h4> <p>""" + str(percentages()[1]) + """</p> <h4>Excluded data</h4> <p>For the following companies either missing and/or incorrect data was found: &nbsp;</p> <p><em>""" + percentages()[2] + """</em> &nbsp;</p> <p>The above tickers have therefore <strong>not</strong> been included in the short report! Please check the data if a company's name returns periodically.</p> <p>If you have any questions and/or remarks, please contact XXX <a href="mailto:example@example.com">here</a>.</p> <p>Kind regards,</p> <p>Name<span style="font-size: 7.0pt; text-transform: uppercase; letter-spacing: .5pt;">&nbsp; &nbsp;</span></p> <p><span style="font-size: 7.0pt; text-transform: uppercase; letter-spacing: .5pt;">THIS IS AN AUTOMATED MESSAGE, PLEASE DO NOT REPLY TO THIS EMAIL ADDRESS!<br /><br /></span></p>
    </html>
    """
    
    message.attach(MIMEText(html, "html"))
    
    # Include attachment
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open("attachment.xlsx", "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="attachment.xlsx"')

    message.attach(part)

    # Enter server details
    server = smtplib.SMTP('smtp.servername.nl', 25)
    server.connect("smtp.servername.nl", 25)
    server.sendmail(sender_email, receiver_email, message.as_string())
    server.quit()

# Check if script has already updated today and if today is a Friday
if datapoints.UpdateTimeStamp.iloc[0].date() == date.today() and date.today().strftime("%A") == 'Friday':

    ################
    # Run function #
    ################

    runscript(datapoints) 

# ############################################      
# #  If data has not yet been updated today  #
# ############################################

if not datapoints.UpdateTimeStamp.iloc[0].date() == date.today():
    
    # Send email to notify data has to be updated   
    sender_email = "example@example.com"
    receiver_email = ['example@example.com']

    message = MIMEMultipart("alternative")
    message["Subject"] = "Please update the short interest information!"
    message["From"] = "Sender name you want to appear in email header"
    message["To"] = ", ".join(receiver_email)

    html = """\
    <html>
    <h4>Please update the short interest information!</h4> <p>Dear colleague,</p> <p>The automatic short interest tool has noticed that today's data has not yet been updated</p> <p>Please update the short interest data file <strong>within 20 minutes </strong>in the designated folder&nbsp;</p> <p>Thank you</p> <p><span style="font-size: 7.0pt; text-transform: uppercase; letter-spacing: .5pt;">THIS IS AN AUTOMATED MESSAGE, PLEASE DO NOT REPLY TO THIS EMAIL ADDRESS!</span></p> <p>&nbsp;</p>
    </html>
    """
    message.attach(MIMEText(html, "html"))

    server = smtplib.SMTP('smtp.servername.nl', 25)
    server.connect("smtp.servername.nl", 25)
    server.sendmail(sender_email, receiver_email, message.as_string())
    server.quit()
    
    # Set timer for 30 minutes to allow for data to be updated
    time.sleep(1800)
    
    # After timer ends, check if data has already been updated in database and if today is Friday. If so send short interest report
    datapoints_new = pandas.io.sql.read_sql(sql, conn)
    if datapoints_new.UpdateTimeStamp.iloc[0].date() == date.today() and date.today().strftime("%A") == 'Friday':
        runscript(datapoints_new)
    
    # If data has been updated today but today is not a Friday, do nothing
    if datapoints_new.UpdateTimeStamp.iloc[0].date() == date.today():
        pass
    
    # If data still was not updated after 30 minutes, send email that data data was not updated and ask to update tomorrow
    else:        
        sender_email = "example@example.com"
        receiver_email = ['example@example.com']

        message = MIMEMultipart("alternative")
        message["Subject"] = "ALERT: You have not updated the short interest data on time"
        message["From"] = "Sender name you want to appear in email header"
        message["To"] = ", ".join(receiver_email)

        html = """\
        <html>
        <h4>You have not updated today's short interest data, please update the data tomorrow!</h4> <p>Dear colleague,</p> <p>The automatic short interest tool has noticed that today's data still has not been updated.</p> <p>The tool will run again tomorrow (except on weekend days) at 15.00, please make sure that you update the data on time!</p> <p>If you have updated the data but still received this email, please contact XXX.</p> <p>Thank you<span style="font-size: 7.0pt; text-transform: uppercase; letter-spacing: .5pt;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></p> <p><span style="font-size: 7.0pt; text-transform: uppercase; letter-spacing: .5pt;">THIS IS AN AUTOMATED MESSAGE, PLEASE DO NOT REPLY TO THIS EMAIL ADDRESS!</span></p>        </html>
        """
        message.attach(MIMEText(html, "html"))

        server = smtplib.SMTP('smtp.servername.nl', 25)
        server.connect("smtp.servername.nl", 25)
        server.sendmail(sender_email, receiver_email, message.as_string())
        server.quit()
