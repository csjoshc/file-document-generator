
Version 3 - 10-11-18
# file-document-generator (sample in wiki)
Pulls  google sheets data through google Drive API, makes folders based on column of names, moves appropriate files

Newly Implemented features
- more resume and cover letter types (+2)
- pull different resumes for different job classifications, rename them appropriately

To add:
- GUI (may not add for a while)
- more types as needed
- add pdf copying (of transcript) using the same name interation (Chiu Transcript position_name) 
- add in a toggle on cmd: 1. run file generator 2. delete transcript pdfs 
 (since each is 300kb after 30 apps it would accumulate 10mb in my OneDrive)
- add this toggle after the date prompt


Reminder on GitHub usage:
* must create new branch on github before trying to push to it from bash*
* make sure this has no missing files compared to the clone from master or else files will dissapear*
- git fetch
- cd to ~/file-document-generator
- git commit -a -m "Message here"
- git checkout v#
- git push 

Description:
Originally, I was trying to find a way to parse a list on my clipboard into python using pandas, but I didn't go far down that route before I wondered if you could pull the list directly from google sheets. 
This site was helpful as a tutorial: 
https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html
Once I had this set up, it naturally followed that I could generate the folders, move the files and fill in placeholder text from a written letter. Beyond the company name, position name and the date, each cover letter would need to be manually edited further and that is of course well outside of my reach right now. Likewise, the resume can be customized but it probably needs to be manual as well so that it is highly relevant to the job role. The best I can ask the script to do is move an extra resume copy into each subfolder so that I can make the changes myself. 

There is also a 'client_secret.json' in the same directory, but I can't upload that since it gives access to my Google drive