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
sheet = client.open("Job apps")
date = input("Enter the date: ")
data = sheet.worksheet(date)
name_list = data.col_values(1, value_render_option='FORMATTED_VALUE')

src = "C:/Users/joshu/OneDrive/Documents/2018/Documents/Reformatted Resume"
src_name = 'Juei-Sheng Joshua Chiu Resume.docx'
source = os.path.join(src, src_name)

src2 = "C:/Users/joshu/OneDrive/Documents/2018/Documents"
src2_name = 'Juei-Sheng Joshua Chiu Cover Letter.docx'
source2 = os.path.join(src2, src2_name)

src2_txt_name = 'Juei-Sheng Joshua Chiu Cover Letter.txt' #raw text of cover letter
source_txt2 = os.path.join(src2, src2_txt_name)
cover_letter_file = open(source_txt2)
cover_letter_txt = cover_letter_file.read()
cover_letter_file.close()

os.chdir('..')#script is in /Documents subfolder for neatness
temp_dir = None #just the company and position name
dst = None #final complete destination path

def crtFldr(list):
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
    for i in list:
        temp_dir = i.strip() #remove spaces from end/beginning of string
    
        try:
            os.makedirs(temp_dir)
        except FileExistsError:
            print(i + ' subfolder already exists!')
            pass
        dst = os.path.join(os.getcwd(),temp_dir)
        shutil.copy(source, dst)
        shutil.copy(source2, dst)
        autofillCoverLetter(dst, temp_dir, date)
       
def autofillCoverLetter(folder, company_and_position, mdy):
    '''
    This function takes the raw cover letter text and replaces 
    placeholder strings with appropriate & unique strings, and then opens, 
    writes to and saves the coverletter docx in the appropriate subfolder
    '''
    file_name = os.path.join(folder, src2_name)
    company_and_position = re.split('-', company_and_position)
    company = company_and_position[0].strip()#remove extra space
    position = company_and_position[1].strip()#remove extra space
    
    mdy = re.split('-',mdy)
    mdy[0] = month_mapping[int(mdy[0])]
    mdy[2] = '20' + str(mdy[2])
    
    tcl = cover_letter_txt
    tcl = re.sub('_month',mdy[0],tcl)
    tcl =  re.sub('_date',mdy[1],tcl)
    tcl = re.sub('_year',mdy[2],tcl)
    tcl = re.sub('position_name', position, tcl)
    tcl = re.sub('company_name', company, tcl)
    
    document = Document(file_name)
    
    #I like this font/size, plus I can always change it 
    style = document.styles['Normal']
    font = style.font
    font.name = 'Garamond'
    font.size = Pt(11)
    document.add_paragraph(tcl)
    document.save(file_name)
crtFldr(name_list)
input('Finished, press enter to exit:')