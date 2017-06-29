# AutoSMS
This project is made for sending text messages to customers of alerts coming from a Knowledge Management System that sends email from triggers/events or at a set periodical time.
As with any new projects coming from a complete beginner, it may not be as pragmatic as I want it to be.
Several methods have been used in order to come up with a working solution for the requirement.
PL/SQL is used in the Knowledge Management System side to trigger an email that will be sent to an email that needs to be setup only for recieving specific mails, so as not to pollute the alerts pool.
Basic HTML is used in order to pull data from the KM system and provide the needed information.
The email is then processed by a Python 2.7 Script with various libraries used sucha as imapclient, smtplib, json, openpyxl, etc.
The Python script is being run in a Linux Platform within a Virtual Machine. For portability purposes and licensing concerns.
BulkSMS.com is the SMS provider used as it is the most user friendly SMS provider that I have gone through when searching for the needed provider.
This "program" aims to be as easily understood by anyone who is able to use a computer with a least a knowledge in spreadsheets and its basic functions, more espescially conditional statements.
This is a project by Jann Paolo Sy - jpaolosy@gmail.com
Initiated May of 2017, while accomplishing a completely functiong program on June of 2017.
