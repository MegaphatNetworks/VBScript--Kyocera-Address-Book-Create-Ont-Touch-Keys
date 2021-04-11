# VBScript: Kyocera Address Book Create One-Touch Keys
## by Gabriel Polmar                                      
## Megaphat Networks                                      
## www.megaphat.info                                      

### INTRODUCTION. 
 A customer had over 100 address book entries but only about 10 one-touch keys (front-panel sshortcuts for scanning and faxing to entries in the address book) created.  I did not want to sit there for a long period of time and considering they had 3 copiers, it was more time efficient for me to write this code which I can reuse as often as needed.  
 
 ### ABOUT THE CODE: 
 The original code was written back in August of 2019.  I am just publishing it now so I don't have to look all over my drives trying to find it.  I *was* planning to create an integrated system in order to manage the address book entirely, but never got around to it.
 
 ### Usage:
 The code must be executed at a command line.  Admin permissions are not required.  Execute simply by performing the following:
 > cscript otkey.vbs "c:\mypath\myab.xml"
 
 The cscript is required in order to execute as it uses the cscript.exe interpreter rather than wscript.  The "c:\mypath\myab.xml" is the literal path of your Address Book XML.  It *must be* in quotes.  Without the quotes, the script may fail.  

### Results:
 2 files will be created, a TXT and a LOG file.  The TXT will be the name of your original XML file with the addition of NEW.txt appended to it.  The LOG will be named similarly.
