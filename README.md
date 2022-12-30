# 365-python-backup
backup your exchange online mailboxes with python, oauth 2.0 compatible!

make sure to install exchangelib and python-dotenv with pip3 install and fill in the appropiate .env file.

also create an app registration in your tenant: https://ecederstrand.github.io/exchangelib/#oauth-on-office-365

usage:
main.py mailfoldername filenametosaveto.mbox

to do:
- add option to automatically fetch all folders
- maybe figure out less api rights
- look into ms graph api
