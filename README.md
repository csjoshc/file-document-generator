
Version 4 - 10-18-18
# file-document-generator (sample in wiki)
Pulls  google sheets data through google Drive API, makes folders based on column of names, moves appropriate files

Newly Implemented features
-minor update: added a 6th letter and resume types(r)

Immediately planned features (v1.5)
- implement word dictionary (words.txt, this can come from anywhere as long as its reasonably comprehensive)
- Divide code into classes and methods (each letter is an object?) this should massively improve readability and comprehension, even if it adds a little redundancy 
- Pseudocode implementation pasted at bottom
- add changelog, script/class/method headers, LICENSE.txt from https://choosealicense.com/licenses/gpl-3.0/
	#!/usr/bin/env python
	__author__ = "Joshua Chiu"
	__copyright__ = "Copyright 2018, Demo Project - File & Document Generator"
	__credits__ = ["Joshua Chiu"]
	__license__ = "GPL"
	__version__ = "1.5"
			
Several Long-term features:
- add pdf copying (of transcript) using the same name interation (Chiu Transcript position_name) 
- add in a toggle on cmd: 1. run file generator 2. delete transcript pdfs 
 (since each is 500kb after 50 apps it would accumulate 25mb in my OneDrive)
- add this toggle after the date prompt (delete all transcript.pdf in the subfolders of a date folder)
- GUI (may not add for a while)

- web crawler - using Scrapy, CSS selectors, pull job descriptions from list of URLs
- word frequency counter - truncate word endings, common articles (the, it), process pulled job descriptions, plaintext resume and cover letter, looking at top n words
- display results as histogram if the match rate is less than 75%, displays word cloud and histogram if match rate < 50% so that the user pays more attention to revising those documents
- if the cover letter is broken into individual paragraphs or even sentences, we can iterate through combinations to generate a higher match rate (wrt/ job description), similar to fminsearch in matlab (but with a higher threshold)
- this is an automatic feature for job.optimizeLetter() to be implemented in the far future. In the master program we can set a GUI so I can choose to do this for individual jobs
- how would we implement this for resumes?
<pre>
Rough idea for user preferences updating script 
Preferences.py
- add default.txt with default name First Middle Last and greeting as "Dear", and default preference to use First, middle and last names, promptCreate = True
checkPref()
	- if preferences.txt doesn't exist: Do you want to set custom preferences for name and greeting? (y/n)
			no:  promptCreate = check_no_setting, then exit preferences script (and go to date prompt)
			yes:create preferences.txt with promptUpdate (in that file) = True, 
				all other fields empty & create one field at a time; press enter to skip
				then recursively call (checkPref())
	- if preferences DOES exist - promptUpdate (y/n) - default is yes (when created) 
			no:  promptUpdate = check_no_setting, then exit preferences script (and go to date prompt)
			yes: 
				if all preferences are populated, ask the user if they want to update any of them 
					no:  promptUpdate = check_no_setting, then exit preferences script (and go to date prompt)
					yes: --> updatePref()
		 		if one or more setting is missing: Some settings have not been set. Do you want to update them (y/n)
					no: promptUpdate = check_no_setting, then exit preferences script (and go to date prompt)
					yes: --> updatePref()	
check_no_setting(): print out settings, input = ask do you want to save this setting, return True or False depending on input
	
updatePref()			
	Do you want to update any of the settings?
	yes: First name is CustomFirst (f), Middle name is CustomMiddle (m), LastName is CustomLast (l), Greeting is CustomDear (g). 
					press the appropriate letter and prompt for user input
				no: All missing fields have been set, proceed to date prompt 
		
Rough idea for reorganized script with Jobject being the core instance for each job

main script.py
Import appropriate modules
From jobclass import Jobject
t_list = [‘b’, ‘bi’, ‘bme, etc] 
m_list = [‘January’, February’, etc]
dir = “\\\\”
new_words = []
	Initialize connection
	(date, mdy, jt_list) = importCloudData() 
		Ask for user input (mm-dd-yyyy): date
		Split date into list of strings
		month = m_list[date.split()[0]]
		text_date = month + “ “ + date[1] + “, “ + date[2]
		Import column 1: company & jobname
		Import column 3: type
		Check to see if length matches - otherwise throw error 
		Job_and_types  is a list of tuples ((job&company, type), etc), use for loop
		Return (text_date, date, job_and_types)
	(letters, dict) = loadDocuments(type_list = t_list, doc_source_dir = dir)
		Locate Resumes : eg & coverletters
		Using forloop and Typelist to generate file names for checking
		Load cover letters into dictionary {} “cover_letters”
		Using forloop and typelist to generate file names to import from
		Load dictionary (words.txt from EdX or any other dictionary - need formatting)
		Called “valid_words”
		return(cover_letters, valid_words)
	ref_data = [date, dir, letters, dict]
processJobs(jt_list,  ref_data)
	Use prefScript.py()
	For i in jt_list:
		(jc, tp) = (jt_list[i][1], jt_list[i][2])
		letter = letters[tp] #the raw cover letter
		job = Jobject(jc, tp, ref_data[0:1], letter)
		Check if Self.getJob() is in ref_data[3] #this is dict
		job.createFolder()
		job.combineLetterText(name)
		For i in job.spellCheck(ref_data[3]) # check before writing to word doc
			If not i in new_words:
				new_words.append(i)
			Else:
				print(“Re-encountered “ + i)
		job.moveCoverLetter()
		job.writeCoverLetter()
		job.moveResume()
addAllowedWords(new_words) #I may want to implement the dictionary functions as a separate python file 
	If len(new_words) > 0:
		print(“Adding new words: “) 
		print(new_words)
		print(“Updating dictionary not implemented yet”)
	else : 
		print(“no new words encountered”)


Jobject (self, company & job, [t_date, t_dir], text) in separate class file
	Self.folder = company & job
	Self.target directory = t_dir + self.folder
	Self.company = self.folder.split()[0]
	Self.job = self.folder.split()[1]
	Self.date = t_date #text date
	Self.type = j_type
	Self.raw_letter = text 
	Self.done_Letter = “”
	Self.unknownWords = []
	getJob()
		Return self.job
	getLetter() 
		Return self.letter
	spellCheck(word_list):
		Split_done_letter = Self.done_Letter.split()
		For i in split_raw_letter:
			If i in word_list:
				Pass
			Else
				print(“the “ + self.type + “ letter type has a unknown word: “ + i)
				self.unknownWords.append(i)
		Return self.unknownWords
	createFolder(): 
		Create appropriate subfolder - company & job
		Self.folder = subfolder directory
	combineLetterText(user_name):
		self.done_letter
	moveCoverLetter()
		Using self.Folder & master
	writeCoverLetter()
		Using self.letter
		Close file
	moveResume()
		Using self.type
		Using self.Folder & master

</pre>
Other notes:

Reminder on GitHub usage:
* must create new branch on github before trying to push to it from bash*
* make sure this has no missing files compared to the clone from master or else files will dissapear*
- cd to file-document-generator
- git fetch
- git commit -a -m "Message here"
- git checkout v#
- git push 

Required Modules to instal (not exhaustive):
pip install --upgrade oauth2client
pip install python-docx
Pip install gspread
client_secret.json in same dir

Description:
Originally, I was trying to find a way to parse a list on my clipboard into python using pandas, but I didn't go far down that route before I wondered if you could pull the list directly from google sheets. 
This site was helpful as a tutorial: 
https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html
Once I had this set up, it naturally followed that I could generate the folders, move the files and fill in placeholder text from a written letter. Beyond the company name, position name and the date, each cover letter would need to be manually edited further and that is of course well outside of my reach right now. Likewise, the resume can be customized but it probably needs to be manual as well so that it is highly relevant to the job role. The best I can ask the script to do is move an extra resume copy into each subfolder so that I can make the changes myself. 

There is also a 'client_secret.json' in the same directory, but I can't upload that since it gives access to my Google drive