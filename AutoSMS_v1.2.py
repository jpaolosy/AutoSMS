# AutoSMS v1.2 - June 28, 2017

# Imports

import imapclient
import smtplib
import pyzmail
import openpyxl
import os
import time
import sys
import json

from timeit import default_timer as timer





# Initiate Global Variables

email='ki.tdm.sms@gmail.com'

password='Republic123'

bulk_sms_code='TDM123'





def readMail():

        

    print("Starting IMAP process...")

    try:

        imap_resolution = 0

        retry_counter = 0        

        imap_obj = imapclient.IMAPClient('imap.gmail.com', ssl=True)

        print("Connected to IMAP")

        updateCounter("imap_success")

        imap_obj.login(email, password)

        imap_obj.select_folder('INBOX', readonly=False)

        print("Checking for new alerts...")



        while retry_counter < 60:

            uids=imap_obj.search('UNSEEN')

            

            if len(uids) < 1:

                sys.stdout.write("\r" + "[No new alerts - Retry " + str(retry_counter+1) + "/60]")

                retry_counter += 1

                imap_resolution = 60

            else:

                print("[" + str(len(uids)) + " new alert/s found]")

                max_uids = max(uids)

                for uid in uids:

                    raw_message = imap_obj.fetch([uid], ['BODY[]', 'FLAGS'])

                    message = pyzmail.PyzMessage.factory(raw_message[uid]['BODY[]'])

                    mail_subj = message.get_subject()

                    message_html = message.html_part.get_payload().decode(message.html_part.charset)

                    plant_site, alert_type, alert_code = subjectParser(mail_subj)

                    message_body = bodyParser(message_html)

                    mail_to = getMailingList(uid, max_uids, plant_site, alert_code, message_body)

                    sendMail(uid, max_uids, plant_site, alert_type,alert_code, message_body, mail_to)		

                    # put codes here to mark messages as read after smtp succeeded

                retry_counter = 60

                imap_resolution = 0

            sys.stdout.flush()

            sleepTime(imap_resolution)

                

        print("\nDisconnecting from IMAP...\n")

        imap_obj.logout()

        

    except:

        updateCounter("imap_failed")

        imap_resolution = 20

        

        try:

            print("IMAP process error!")

            imap_obj.logout()

        except:

            print("IMAP connection failed!!!")

   



def sendMail(uid, max_uids, plant_site, alert_type, alert_code, message_body, mail_to):

  

    smtp_resolution = 0

    retry_counter = 0

    print("Starting SMTP process...")

    while retry_counter < 10:

        sleepTime(smtp_resolution)

        try:

            smtp_obj = smtplib.SMTP('smtp.gmail.com', 587)

            print("Connected to SMTP")

            updateCounter("smtp_success")

            #smtp_obj.set_debuglevel(1)

            smtp_obj.ehlo()

            smtp_obj.starttls()

            smtp_obj.login(email, password)

            

            if uid <= max_uids:

                print("Sending " + alert_type + " message to " + plant_site + " group:")

                for recipients in mail_to:

                    print("\t" + recipients)

                    #smtp_obj.sendmail(email, recipients, "Subject:" + str(bulk_sms_code) + "\r\n"+ str(message_body))

                    updateCounter("credits_available")

                print("\nMessage: " + message_body)

                updateCounter("alerts_delivered")

                if uid == max_uids:

                    retry_counter=10

                    print("Disconnecting from SMTP...")

                    smtp_obj.quit()



        except:

            updateCounter("smtp_failed")

            retry_counter += 1

            smtp_resolution = 20

            

            try:

                print("SMTP process error!")

                smtp_obj.quit()

            except:

                print("SMTP connection failed!!!")







def getMailingList(uid, max_uids, plant_site, alert_code, message_body):



    print("Getting mailing list...")

    os.chdir('/home/rctdm/Desktop/AutoSMS/directory')

    work_book = openpyxl.load_workbook('directory.xlsx', data_only=True)

    lookup_sheet = work_book.get_sheet_by_name('LOOKUP')

    directory_sheet = work_book.get_sheet_by_name(plant_site)



    # Plant site lookup

    for r in range(1,8):

        plant_site_lookup = lookup_sheet.cell(row=r, column=8).value

        if plant_site_lookup == plant_site:

            row_count=lookup_sheet.cell(row=r, column=9).internal_value



    # Append contacts to mail_to list        

    mail_to = []

    for r in range (2, row_count+2):

        if directory_sheet['F' + str(r)].internal_value.find(alert_code) > -1:

            mail_to.append(directory_sheet['E' + str(r)].internal_value)

    return mail_to





def bodyParser(message_html):



    if message_html != None:

        html_len = 31500

        start_len = message_html.find("<BODY>",html_len)+6

        end_len = message_html.find("</BODY>",html_len)

        sliced_body = message_html[start_len:end_len]

        message_body = sliced_body.replace(':', '.')

        return message_body

    else:

        print("Mail body is empty!")





def subjectParser(mail_subj):

	

    if mail_subj != None:

        plant_site = mail_subj[-4:-1]

        end_len = mail_subj.find(" ")

        alert_type = mail_subj[:end_len]

        alert_code=""

        if alert_type == "StopLog":

        	alert_code = "S"

        elif alert_type == "DailyProduction":

        	alert_code = "P"

        elif alert_type == "DailyQuality":

        	alert_code = "Q"

        elif alert_type == "Environment":

        	alert_code = "E"

        #elif alertType == "Department":

        	#alert_code = "X"

        return plant_site, alert_type, alert_code

    else:

        print("Mail subject is empty!")

        

def sleepTime(resolutionTime):

	time.sleep(resolutionTime)

	        

        

def updateCounter(counter_object):

        

    with open ("/home/rctdm/Desktop/AutoSMS/counterLog.json") as counter_append:

        json_data = json.load(counter_append)

    if counter_object != "credits_available":

        json_data["counterLog"][counter_object] += 1

    else:

        json_data["counterLog"][counter_object] -= 1

    with open("/home/rctdm/Desktop/AutoSMS/counterLog.json", "w") as counter_save:

        counter_save.write(json.dumps(json_data))

        

		

def printCounterValues():

    with open ("/home/rctdm/Desktop/AutoSMS/counterLog.json") as counter_values:

        json_data = json.load(counter_values)



    imap_failed = json_data["counterLog"]["imap_failed"]

    imap_success = json_data["counterLog"]["imap_success"]

    smtp_failed = json_data["counterLog"]["smtp_failed"]

    smtp_success = json_data["counterLog"]["smtp_success"]

    alerts_delivered = json_data["counterLog"]["alerts_delivered"]

    credits_available = json_data["counterLog"]["credits_available"]



    print("Sent:" + str(alerts_delivered) + "\t\tCredits:" + str(credits_available) + "\nIMAP:" + str(imap_success) + "/" + str(imap_failed) + "\tSMTP:" + str(smtp_success) + "/" + str(smtp_failed))



     

class CounterLogging():

    pass







def main():



    while (True):

        os.system("clear")

        print("\tAutoSMS v1.2")

        print("-" * 30)

        start_timer = timer()

        printCounterValues()

        print("-" * 30)

        print(">>>Starting AutoSMS")

        readMail()

        end_timer = timer()

        elapsed_time = round((end_timer-start_timer),2)

        print("Time elapsed: " + str(elapsed_time) + " seconds")

        print(">>>Restarting AutoSMS")

        sleepTime(20)

        os.system("clear")

        



if __name__ == "__main__":

	main()

