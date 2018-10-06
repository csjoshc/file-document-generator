# file-document-generator (sample in wiki)
Pulls  google sheets data through google Drive API, makes folders based on column of names, moves appropriate files

there is also a 'client_secret.json' in the same directory, but I can't upload that since it gives access to my Google drive

Originally, I was trying to find a way to parse a list on my clipboard into python using pandas, but I didn't go far down that route before I wondered if you could pull the list directly from google sheets. 

This site was helpful as a tutorial: 
https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html

Once I had this set up, it naturally followed that I could generate the folders, move the files and fill in placeholder text from a written letter. Beyond the company name, position name and the date, each cover letter would need to be manually edited further and that is of course well outside of my reach right now. Likewise, the resume can be customized but it probably needs to be manual as well so that it is highly relevant to the job role. The best I can ask the script to do is move an extra resume copy into each subfolder so that I can make the changes myself. 
