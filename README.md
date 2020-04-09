# Project Title
PersonalizedEmailSender is a small package which sends personalized emails to a list of recipients.

## Description
This script is able to send personalized emails to a list of recipients. 
It is currently designed to specifically send an overview of financial costs made for certain activities to an organization's members. 
However, it can be extended to send other types of emails as well. 


## Getting Started
- Put the installed file in a new directory
- Add a configuration file (json file) to the same directory
- Add a csv or excel file with financial information, names, and email-addresses to the same directory
- Run the executable

### Installing
- Download the .exe file 
- Run the .exe file

### Prerequisites
See requirements.txt for a full list of requirements in case you want to change the script.

### Folder structure
-Directory
-- file.exe
-- fin.xlsx
-- configuration.json

### json structure
{
"senderemail": "[your_email_goes_here]",
"password": "[your_password_goes_here]",
"xlsxfile"L "[your_xlsx_filepath_goes_here]"
}

