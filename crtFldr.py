# -*- coding: utf-8 -*-
"""
Created on Fri Oct  5 13:38:54 2018
https://gspread.readthedocs.io/en/latest/
@author: joshu
"""
import os
import shutil 
import gspread
import re
from docx import Document 
from docx.shared import Pt
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
month_mapping = {1:'January', 2:'February', 3:'March', 4:'April', 5:'May', 
                 6:'June', 7:'July', 8:'August', 9:'September', 10:'October', 
                 11:'November', 12:'December'}
letters = {'b':'', 'd':'', 's':'', 'bi':'', 'bme':''} 
#contains letter types and will contain letters

sheet = client.open("Job apps")
date = input("Enter the date: ")
data = sheet.worksheet(date)
name_list = data.col_values(1, value_render_option='FORMATTED_VALUE')
type_list = data.col_values(3, value_render_option='FORMATTED_VALUE')
if len(name_list) != len(type_list):
    print("Name and types misaligned or missing. Quitting.")
    input('press enter to exit:')
    exit()
src = "C:/Users/joshu/OneDrive/Documents/2018/Documents"
src2_name = 'Chiu Cover Letter.docx' #empty cover letter with formatting
source2 = os.path.join(src, src2_name)

for x in letters: #reads and stores cover letters
    filename = "CL-" + x + ".txt"
    tempfile = open(os.path.join(src, filename))
    letters[x] = tempfile.read()
    tempfile.close()


os.chdir('C:/Users/joshu/OneDrive/Documents/2018') #script is in github in joshu
temp_dir = None #just the company and position name
dst = None #final complete destination path

def crtFldr(n_list, t_list):
    '''
    this function creates the appropriate Date folder, populated with 
    appropriate subfolders named in 'Company - Position' format

    For each folder, a copy of resume and an empty cover letter docx with 
    formatting is moved into each subfolder as it is created
        
    Then it calls the autofillCoverLetter function once to replace 
    the placeholders (date, company name, position name) from the appropriate
    text file containing raw cover letter text and write to the appropriate 
    Cover Letter docx
    '''
    
    try:
        os.makedirs(date)
    except FileExistsError:
        print(date + ' subfolder already exists!')
        overwrite_perm = input('Do you want to overwrite the ' + date + ' folder? (y/n):')
        if overwrite_perm == 'y':
            pass
        elif overwrite_perm == 'n':
            print('No: Exiting to avoid overwriting ' + date + ' folder')
            exit()
        else:
            print('Unknown response. Exiting to avoid overwriting ' + date + ' folder')
            exit()
        pass
    os.chdir(date)
    count = 0
    for i in n_list:
        temp_dir = i.strip() #remove spaces from end/beginning of string
        l_type = t_list[count] #letter type (single char code) 
        try:
            os.makedirs(temp_dir)
        except FileExistsError:
            print(i + ' subfolder already exists!')
            pass

        src_name = 'R-d.docx' #default resume type
        if l_type in letters:
            src_name = 'R-' + l_type + '.docx' #name of correct resume filename
        source = os.path.join(src, src_name)
        dst = os.path.join(os.getcwd(),temp_dir)
        shutil.copy(source, dst) #resume
        shutil.copy(source2, dst) #copy cover letter
        autoClRe(dst, temp_dir, date, l_type)
        count += 1
def autoClRe(folder, company_and_position, mdy, clt):
    '''
    This function takes the raw cover letter text and replaces 
    placeholder strings with appropriate & unique strings, and then opens, 
    writes to and saves the coverletter docx in the appropriate subfolder
    '''
    os.chdir(company_and_position)
    company_and_position = re.split('-', company_and_position)
    company = company_and_position[0].strip()#remove extra space
    position = company_and_position[1].strip()#remove extra space
    resume_name = 'Chiu Resume - '+position+'.docx'
    cover_name = 'Chiu Cover Letter - '+position+'.docx'
    
    try:
        os.rename(('R-' + clt + '.docx'), resume_name) #rename resume with position
        os.rename(src2_name, cover_name) # rename cover letter with position
    except FileExistsError:
        print("Files for " + company + "'s " + position + " position already exist.")
    mdy = re.split('-',mdy)
    mdy[0] = month_mapping[int(mdy[0])]
    mdy[2] = '20' + str(mdy[2])
    
    if not clt in letters:
        print("Unknown Cl & Re type. Defaulting to 'd' type for " + position 
              + " at " + company + ".")
        clt = 'd'
    
    tcl = letters[clt]#copy the letter text for editing
    tcl = re.sub('_month _date, _year', (mdy[0] + ' ' + mdy[1] + ', ' + mdy[2]), tcl)
    if position[0].lower() in ('a', 'e', 'i', 'o', 'u'):
        print("Changing article a to an for position name " + position)
        tcl = re.sub('a position_name', ('an ' + position), tcl)
    else:
        tcl = re.sub('position_name', position, tcl)
    tcl = re.sub('company_name', company, tcl)
    
    document = Document(cover_name)
    #I like this font/size, plus I can always change it 
    style = document.styles['Normal']
    font = style.font
    font.name = 'Garamond'
    font.size = Pt(11)
    document.add_paragraph(tcl)
    document.save(cover_name)
    os.chdir('..')
crtFldr(name_list, type_list)
input('Finished, press enter to exit:')