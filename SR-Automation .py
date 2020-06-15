#! python3

import openpyxl, pprint, os, glob, time, shutil, re, zipfile, PyPDF2, sys
import xml.etree.ElementTree
import pandas as pd
from datetime import datetime, date
from docx import Document
from pathlib import Path
from smtplib import SMTP
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from secrets import jobber_us, jobber_pw, onedrive_us, onedrive_pw

start = time.perf_counter()
os.chdir('C:\\Users\\info\\OneDrive\\1. M2M Administration\\AGED CARE\\Jobber Bot')

#DOB Calculation (to be used in data extraction 
def calculate_age(born):
    today = date.today()
    return today.year - born.year - ((today.month, today.day) < (born.month, born.day))

class DataExtraction():
    def checkFolder(self):
        global document, latest_file, read_pdf, list_of_files, last_file, ext
        #Latest File in Directory - MIGHT NEED TO BE CHNAGED TO THE OLDEST FILE IN DIRECTORY IF RUNNING ON SCHED
        #NEED TO ADD IN FUNCTIONALITY FOR RESTRICTED .DOCX FILES &
        #.PDF versions of the document
        list_of_files = glob.glob(os.getcwd() + '\\*')
        #Ignore debug.log file
        list_of_files.remove('C:\\Users\\info\\OneDrive\\1. M2M Administration\\AGED CARE\\Jobber Bot\\debug.log')
        list_of_files.remove('C:\\Users\\info\\OneDrive\\1. M2M Administration\\AGED CARE\\Jobber Bot\\Jobber-Client.xlsx')
        last_file = max(list_of_files, key=os.path.getctime)
        #Just get the file name
        latest_file = last_file.split('\\')[-1]
        ext = latest_file.split('.')[-1]
        # print(latest_file)
     
    def ExtractDocxData(self):
         # Designed to extact data ONLY FROM May 2020 Service Request form
        # NEED TO SIMPLIFY THIS -- WORK SMARTER NOT HARDER -- NESTED LOOP?
        # These need to be global to pass to other functions
        try:
            global data, info, provider, manager, loc, m_phone, email, name, name_fn, name_ln,\
                   dob, street1, city, state, postcode, number, med_detail, overview,\
                   nok, nok_contact, nok_rel, start, reason, date_of_birth, age, prim_contact,\
                    freq, postcode, docxExtractError, s, i, address, details, headings, all_details, headers, radio, notes, l1, remove, z, c
            ## Politely borrowed from:
            ## https://stackoverflow.com/questions/22756344/how-do-i-extract-data-from-a-doc-docx-file-using-python
            WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
            PARA = WORD_NAMESPACE + 'p'
            TEXT = WORD_NAMESPACE + 't'
            TABLE = WORD_NAMESPACE + 'tbl'
            ROW = WORD_NAMESPACE + 'tr'
            CELL = WORD_NAMESPACE + 'tc'


            with zipfile.ZipFile(latest_file) as docx:
                tree = xml.etree.ElementTree.XML(docx.read('word/document.xml'))

            info = []
            for table in tree.iter(TABLE):
                for row in table.iter(ROW):
                    for cell in row.iter(CELL):
                        info.append(''.join(node.text for node in cell.iter(TEXT)))

            #Remove duplicates
            #Replace empty fields with 'Blank'
            remove = ['Click or tap here to enter text.', 'Click or tap to enter a date', '']
            for detail in info:
                if detail in remove:
                    z = [z for z, n in enumerate(info) if n == detail]
                    for y in z:
                        info[y] = 'Blank'
            #Show Data Extracted
            for detail in enumerate(info):
                print(detail)

            #Find selected radio buttons and add notes if neccessary
            freq = []
            notes = []
            radio = []
            for item in info:
                if u"\u2612" in item:
                    l1 = []
                    l1.append(item)
                    for j in l1:
                        i = j.index(u"\u2612")
                        radio.append(j[i+2])
                if 'female only' in item.lower():
                    notes.append('Female Only')
            # Change radio[0] into full word
            if radio[0] == 'C':
                radio[0] = 'Client'
                notes.append(f'Please Contact {radio[0]}')
                prim_contact = radio[0]
            elif radio[0] == 'N':
                radio[0] = 'NOK'
                notes.append(f'Please Contact {radio[0]}')
                prim_contact = radio[0]
            else:
                prim_contact = ''
            if len(radio) >= 3:
                freq = radio[1] + '(' + radio[2] + ')'
            elif len(radio) >= 2:
                freq = radio[-1]
            else:
                freq = 'Blank'
            freq = str(freq).strip("[|]|''|''")
            #Remove item from headings if not in the document
            # Get Address as BEST AS POSSIBLE
            statereg = re.compile(r'NSW|VIC|TAS|QLD|NT|ACT|WA')
            s = statereg.search(info[24])
            if s == None:
                    if len(re.findall('[0-9]{4}$', info[24])) == 1:
                        postcode = ''.join(re.findall('[0-9]{4}$', info[24]))
                        address = info[24].split(' ')
                        city = address[-2].strip(',|\\')
                        street1 = ' '.join(address[:-2]).strip(',|\\')
                        state = 'Blank'
                    else:
                        street1 = info[24]
                        state = 'Blank'
                        city = 'Blank'
                        postcode = 'Blank'
            # Else use regex to get perfect data from contact
            else:
                state = s.group()
                #Format the data correctly
                i = info[24].split(state)
                x = i[0].split(' ')
                s = x[:-2]
                s2 = ' '.join(s)
                street1 = s2.strip(',')
                c = x[-2]
                city = c.strip(',')
                postcode = i[1]
            # Serv provider info
            provider = info[4]
            manager = info[7]
            loc = info[10]
            m_phone = info[13]
            email = info[16]
            # Get client details
            name = info[19].split(' ')
            name_fn = str(name[:-1]).strip("[|''|''|]")
            name_ln = name[-1]
            number = info[26]
            nok = info[33]
            nok_rel = info[37]
            nok_contact = info[35]
            dob = info[21]
            med_detail = info[28]
            overview = info[30]
            start = info[52].strip('**mark today for ASAP')
            reason = info[54]
            freq = str(freq).strip("[|]|''|''")
            # Use function to get Age
            date_of_birth = datetime.strptime(dob, "%d/%m/%Y")
            #Add start date to notes
            if start != 'Blank':
                notes.append(f'Please begin services on/after {start}')
            notes = ' - '.join(notes)
            print(notes)
            def calculate_age(born):
                today = date.today()
                return today.year - born.year - ((today.month, today.day) < (born.month, born.day))
            age = calculate_age(date_of_birth)

        except Exception as e:
            docxExtractError = e
            print(f'Error Extracting Data From {latest_file}')
    def ExtractPdfData(self):
        global page, page2, provider, manager, loc, m_phone, email, name, name_fn, name_ln,\
               dob, address, street1, city, state, number, med_detail, overview,\
               nok, nok_contact, nok_rel, start, reason, date_of_birth, age, notes, \
               prim_contact, freq, postcode, pdfExtractError
        #remove unecessary elements and empty spaces
        read_pdf = PyPDF2.PdfFileReader(last_file)
        page = read_pdf.getPage(0).extractText()
        page2 = page.split('\n')
        while(' ' in page2):
                page2.remove(' ')
        ##page2 = list(filter(lambda a: a != '\n', page))
        try:
            #Assign page elements to their right position
            #Company/Case Manager Details
            provider = page2[12] 
            manager = page2[14]
            loc = page2[16]
            m_phone = page2[18]
            email = page2[20]

            # Get Client Details 
            name = page2[30].split()
            name_fn = name[0]
            name_ln = name[1]
            dob = page2[34]
            #PDF requires different method to extract address
##            address = page2[36].split(' ')
##            street1 = street1 = ' '.join([str(elem) for elem in address[:3]])
##            city = address[3]
##            state = re.sub('[\W_]+', '', address[-2])
##            postcode = address[-1]
            #Using Regex to find address and SPLIT it up based on state
            #Hopefully somewhat correct 
            statereg = re.compile(r'NSW|VIC|TAS|QLD|NT|ACT|WA')
            s = statereg.search(page2[36])
            state = s.group()

            #Format the data correctly
            i = page2[36].split(state)
            x = i[0].split(' ')
            s = x[:-2]
            s2 = ' '.join(s)
            street1 = s2.strip(',')
            c = x[-2]
            city = c.strip(',')
            postcode = i[1]
            
            number = page2[41]
            med_detail = page2[49]
            overview = page2[53]
            nok = page2[59]
            nok_contact = page2[65]
            nok_rel = page2[62]
            start = page2[70]
            reason = page2[101]
            # Use function to get Age
            date_of_birth = datetime.strptime(dob, "%d/%m/%Y")
            def calculate_age(born):
                today = date.today()
                return today.year - born.year - ((today.month, today.day) < (born.month, born.day))
            age = calculate_age(date_of_birth)

            # Work out how to get radio buttons, find state & PC
            prim_contact = ''
            freq = ''
            notes = ''
        except Exception as e:
            pdfExtractError = e
            print(f'There was an error extracting data from {latest_file}\n\n\{pdfExtractError}')
                
    def ExtractBilling(self):
        ## SEARCH BILLING INFORMATION SPREADSHEET
        ## WORK OUT HOW TO CHECK IF THE CLIENT IS NOT IN SPREADSHEET AT ALL
        global billing, res1
        df = pd.read_excel('C:\\Users\\info\\OneDrive\\1. M2M Administration\\PROPOSALS\\Jarrods Folder\\Automation\\Accounts_Contacts.xlsx')
        res1 = df[df['Company'].str.contains(provider, na=False)]
        billing = []
        if len(res1.index) > 1:
            billing = (df[df['Company'].str.contains(loc, na=False)])
        else:
            billing = res1

class FileManip():
    def FileMove(self):
        global new_folder
        new_folder = 'C:\\Users\\info\\OneDrive\\1. M2M Administration\\AGED CARE\\AGED CARE CLIENT PROVIDERS\\'+ \
          state + '\\' + provider + '\\' + name_fn + ' ' + name_ln
        Path(new_folder).mkdir(parents=True, exist_ok=True)
        #Rename file and move to directory under the right name
        os.rename(latest_file, new_folder + '\\' + name_fn + ' ' + name_ln + '- Service Request ' +\
                date.today().strftime("%d-%m-%Y") + '.' + ext)

    def WriteToExcel(self):
        global details, sheet, wb
        ##Write name to cells -- 
        ## CHECK IF ENTRY IS DUPLICATED -- DO NOT LET DUPLICATED DATA PASS ON TO .XLSX FILE AND CRM
        wb = openpyxl.load_workbook('C:\\Users\\info\\OneDrive\\1. M2M Administration\\AGED CARE\\Jobber Bot\\Jobber-Client.xlsx')
        sheet = wb['Imports']
        ## For Loop to insert new data below existing data
        details = [[datetime.now().strftime("%d/%m/%Y %H:%M:%S"), name_fn, name_ln, provider, number,\
                    street1, city, state, postcode, dob, age, prim_contact, freq, med_detail, overview, nok, nok_contact,\
                    nok_rel, start, reason, notes, billing['Billing Street 1'].values[0], billing['Billing Street 2'].values[0],\
                    billing['Billing City'].values[0], billing['Billing State'].values[0], str(billing['Postcode'].values[0])]]
        for detail in details:
            sheet.append(detail)
        wb.save('C:\\Users\\info\\OneDrive\\1. M2M Administration\\AGED CARE\\Jobber Bot\\Jobber-Client.xlsx')
        print('Data written to Jobber-client.xlsx')

class JobberBot():
    def __init__(self):
        self.driver = webdriver.Chrome()

    def login(self):
        self.driver.get('https://secure.getjobber.com/login')
        time.sleep(1) # Let the user actually see something!
        #LOGGING INTO JOBBER
        email_box = self.driver.find_element_by_name('user_session[login]')
        email_box.send_keys(jobber_us)
        password_box = self.driver.find_element_by_name('user_session[password]')
        password_box.send_keys(jobber_pw)
        self.driver.find_element_by_name('commit').click()
        #Create 'New Client' button
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[1]/div/div[2]/div[2]/div/div/div/div[1]/a').click()
        time.sleep(1.5)
    
    def FieldPopulation(self):
        # Dropdown the 'Additional Client details' menu
        self.driver.find_element_by_xpath('//*[@id="new_client"]/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-header/div/div[2]/div/div').click()
        time.sleep(1.5)
        # FIELDS TO FILL
        #Name
        time.sleep(1.5)
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/div[2]/label/div[1]/div/div/div[1]/div[2]/placeholder-field/input').send_keys(name_fn)
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/div[2]/label/div[1]/div/div/div[1]/div[3]/placeholder-field/input').send_keys(name_ln)
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/div[2]/label/div[1]/div/div/div[2]/div/placeholder-field/input').send_keys(provider)
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/div[2]/div[2]/ul/li/div/div[2]/div/div/div[2]/div/div[1]/placeholder-field/input').send_keys(number)
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/div[2]/div[3]/ul/li/div/div[2]/div/div/div[2]/div/div[1]/placeholder-field/input').send_keys(billing['Email'].values[0])
        #Uncheck 'Billing address is the same...' checkbox
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[1]/label/sg-icon').click()
        time.sleep(1.5)
        #Address
        # Dropdown opens - SEND 'ESCAPE' KEY TO EXIT'
        time.sleep(2)
        self.driver.find_element_by_xpath('//*[@id="new_client"]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/div[1]/div/div[1]/div/placeholder-field/input').send_keys(street1)
        self.driver.find_element_by_xpath('//*[@id="new_client"]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/div[1]/div/div[1]/div/placeholder-field/input').send_keys(webdriver.common.keys.Keys.ESCAPE)
        time.sleep(1.5)
        #City
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[3]/div[2]/div/div[1]/div/div[3]/div[1]/placeholder-field/input').send_keys(city)
        #State
        self.driver.find_element_by_xpath('//*[@id="new_client"]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/div[1]/div/div[3]/div[2]/placeholder-field/input').send_keys(state)
        #Postcode
        self.driver.find_element_by_xpath('//*[@id="new_client"]/div/div/div/div[1]/div[2]/div[3]/div[2]/div/div[1]/div/div[4]/div[1]/placeholder-field/input').send_keys(postcode)
        #Contact details
        #driver.find_element_by_xpath('//*[@id="client_phones_attributes_1587200583889_number"]').sendkeys('9844 3360')
        #Case Manager
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-body/div[1]/placeholder-field[1]/input').send_keys(manager)
        #Case Manager Contact
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-body/div[1]/placeholder-field[3]/input').send_keys(m_phone)
        #DOB
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-body/div[1]/placeholder-field[7]/input').send_keys(dob)
        #Age
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-body/div[1]/placeholder-field[8]/input').send_keys(age)
        #Home/Mobile
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-body/div[1]/placeholder-field[9]/input').send_keys(number)
        #Next of Kin
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-body/div[1]/placeholder-field[10]/input').send_keys(nok)
        #Next of Kin Contact
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-body/div[1]/placeholder-field[11]/input').send_keys(nok_contact)
        #Availability
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-body/div[1]/placeholder-field[12]/input').send_keys(start)
        #Medical Notes
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-body/div[1]/placeholder-field[13]/input').send_keys(med_detail)
        time.sleep(0.5)
        #Additional Notes
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-body/div[1]/placeholder-field[14]/input').send_keys(overview)
        #Massage details
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-body/div[1]/placeholder-field[15]/input').send_keys(reason)
        time.sleep(2)
        #Enter Billing Details 
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[2]/div/div[1]/div/placeholder-field/input').send_keys(billing['Billing Street 1'].values[0])
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[2]/div/div[2]/div/placeholder-field/input').send_keys(billing['Billing Street 2'].values[0])
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[2]/div/div[3]/div[1]/placeholder-field/input').send_keys(billing['Billing City'].values[0])
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[2]/div/div[3]/div[2]/placeholder-field/input').send_keys(billing['Billing State'].values[0])
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[2]/div/div[4]/div[1]/placeholder-field/input').send_keys(str(billing['Postcode'].values[0]))
        #Select Australia from dropdown
        select_fr = Select(self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[2]/div/div[4]/div[2]/div/select'))
        select_fr.select_by_index(0)
        # CREATE CLIENT
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[2]/div[2]/button').click()
        print('Successfully created client ' + str(name_fn) + str(name_ln) + 'in Jobber')
        time.sleep(2)

        ### Need to get past possible 'duplicate' pop up:
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[3]/div/div[2]/div/div[2]/form/input[96]').click()
        time.sleep(2)
        ## Add AgedCare Tag
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/div[2]/div[2]/div[1]/div[2]/a').click()
        time.sleep(1.5)
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/div[2]/div[2]/div[2]/form/div/div[1]/placeholder-field/input').send_keys('AgedCare')
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/div[2]/div[2]/div[2]/form/div/div[2]/a').click()
        print('Added AgedCare Tag')
        time.sleep(1)
        ## CREATE JOB
        #Open Dropdown
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/div[1]/div[4]/div[1]/div/div/button').click()
        time.sleep(1)
        #Select 'job'
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/div[1]/div[4]/div[1]/div/div/div[1]/nav/a[3]').click()
        #Add freq to job title
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form[1]/div/div/div/div[1]/div/div[2]/div[1]/div[2]/div[1]/div/placeholder-field/input').send_keys(freq)
        #Add notes to job
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form[1]/div/div[1]/div/div[1]/div/div[2]/div[1]/div[2]/div[2]/div/placeholder-field/textarea').send_keys(notes)
        #Check 'Schedule Later'
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form[1]/div/div[1]/div/div[2]/div[2]/div[1]/div/div[1]/div/div[2]/div/div[4]/div/label/sg-icon').click()
        #Save Job
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form[1]/div/div[1]/div/div[2]/div[5]/div[2]/div/div/a').click()
        print(f'Successfully created an unassigned job for {name_fn} {name_ln} in Jobber.')
        time.sleep(1)
        self.driver.close()
        self.driver.quit()

class ConfEmail():
    def SuccessEmail(self):
        ## SEND AN EMAIL TO NOTIFY STAFF
        debuglevel = 1

        smtp = SMTP()
        smtp.set_debuglevel(debuglevel)
        smtp.connect('awcp027.server-cpanel.com', 587)
        smtp.login('donotreply@massage2motivate.com.au', 'magnoliA24@')

        from_addr = "Automated Bot <donotreply@massage2motivate.com.au>"
        to_addr = "agedcare@massage2motivate.com.au"

        subj = F'SYNC COMPLETED FOR {name_fn} {name_ln}'
        date = datetime.now().strftime( "%d/%m/%Y %H:%M" )
        message_text = f'Hello,\n\n\
        A service request for {name_fn} {name_ln} has been recieved, processed and moved to the following file location:\n\n\
        {new_folder}\n\n\n\
        The following billing details were entered:\n\
        Company: {billing["Company"].values[0]}\n\
        Address: {billing["Billing Street 1"].values[0]}, {billing["Billing Street 2"].values[0]}, {billing["Billing City"].values[0]}, {billing["Billing State"].values[0]} {str(billing["Postcode"].values[0])}\n\n\
        This task was completed at {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}\n\n\n\
        Kind regards,\n\n\
        Your Friendly Automated Bot.'

        msg = "From: %s\nTo: %s\nSubject: %s\nDate: %s\n\n%s" \
            % ( from_addr, to_addr, subj, date, message_text )

        smtp.sendmail(from_addr, to_addr, msg)
        smtp.quit()   

    def ErrorEmail(self):
        ## SEND AN EMAIL TO NOTIFY STAFF
        debuglevel = 1

        smtp = SMTP()
        smtp.set_debuglevel(debuglevel)
        smtp.connect('awcp027.server-cpanel.com', 587)
        smtp.login('donotreply@massage2motivate.com.au', 'magnoliA24@')

        from_addr = "Automated Bot <donotreply@massage2motivate.com.au>"
        to_addr = "agedcare@massage2motivate.com.au"

        subj = F'ERROR COMPLETING SYNC FOR {name_fn} {name_ln}'
        date = datetime.now().strftime( "%d/%m/%Y %H:%M" )
        message_text = f'Hello,\n\n\
        An error has occured while processing the service request for {name_fn} {name_ln}. \n\n\
        The Service Request has been moved to the following location:\n\n\
        {new_folder}\n\n\n\
        Error message:\n\
        {error}\n\n\n\
        This task was completed at {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}\n\n\n\
        Kind regards,\n\n\
        Your Friendly Automated Bot.'

        msg = "From: %s\nTo: %s\nSubject: %s\nDate: %s\n\n%s" \
            % ( from_addr, to_addr, subj, date, message_text )

        smtp.sendmail(from_addr, to_addr, msg)
        smtp.quit()
        
    def WrongVersion(self):
        ## SEND AN EMAIL TO NOTIFY STAFF ThAT THE DOCUMENT IS NOT THE CORRECT FORMAT
        debuglevel = 1

        smtp = SMTP()
        smtp.set_debuglevel(debuglevel)
        smtp.connect('awcp027.server-cpanel.com', 587)
        smtp.login('donotreply@massage2motivate.com.au', 'magnoliA24@')

        from_addr = "Automated Bot <donotreply@massage2motivate.com.au>"
        to_addr = "agedcare@massage2motivate.com.au"

        subj = F'ERROR PARSING {latest_file}'
        date = datetime.now().strftime( "%d/%m/%Y %H:%M" )
        message_text = f'Hello,\n\n\
        An error has occured while parsing a the following .docx file: {latest_file}\n\n\n\
        Please only place Service Requests with "Latest Form Issued May 2020" under the M2M logo in the Jobber Bot folder\n\n\n\n\n\
        This task was completed at {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}\n\n\n\
        Kind regards,\n\n\
        Your Friendly Automated Bot.'

        msg = "From: %s\nTo: %s\nSubject: %s\nDate: %s\n\n%s" \
            % ( from_addr, to_addr, subj, date, message_text )

        smtp.sendmail(from_addr, to_addr, msg)
        smtp.quit()

#Initialize prog to check folder
#Start building a structured conditional list of actions for program
#i.e. if x failes, program has done y so far - passed on into en email
DataExt = DataExtraction()
Conf = ConfEmail()
manip = FileManip()
#Check folder for new files
try:
    DataExt.checkFolder()
except Exception:
    os._exit(0)
missing_billing = ['Blank', 'blank@blank.com.au', 'Blank', 'Blank', 'Blank', 'Blank', 'Blank', 'Blank']
if '.pdf' in latest_file:
    print('Latest file is a .pdf')
    # Try to run program and log error if error along the way
    try:
        DataExt.ExtractPdfData()
    except Exception as error:
        Conf.ErrorEmail()
        os._exit(0)
    try:
        DataExt.ExtractBilling()
        if len(billing) == 0:
            df_length = len(billing)
            billing.loc[df_length] = missing_billing
    except Exception as error:
        Conf.ErrorEmail()
        os._exit(0)
    try:
        manip.FileMove()
    except Exception as error:
        Conf.ErrorEmail()
        os._exit(0)
    try:
        manip.WriteToExcel()
    except Exception as error:
        Conf.ErrorEmail()
        os._exit(0)
    ## Start Bot
    try:
        bot = JobberBot()
        bot.login()
        try:
            bot.FieldPopulation()
        except Exception as error:
            Conf.ErrorEmail()
            os._exit(0)
    except Exception as error:
        Conf.ErrorEmail()
        os._exit(0)
    # If Program makes it to here: success email will be sent!!
    try:
        Conf.SuccessEmail()
    except Exception as error:
        Conf.ErrorEmail()
        os._exit(0)
elif '.docx' in latest_file:
    document = Document(latest_file)
    print('Latest file is a .docx')
    # Try to extract data else send us an error message
    # Only extract data if document is May 2020 version -- else send error message
    if document.sections[0].header.paragraphs[1].text == 'Latest Form Issued May 2020':
        print('Document is correct version issued in May 2020')
        try:
            DataExt.ExtractDocxData()
        except Exception as error:
            Conf.ErrorEmail()
            os._exit(0)
        try:
            DataExt.ExtractBilling()
            if len(billing) == 0:
                df_length = len(billing)
                billing.loc[df_length] = missing_billing
        except Exception as error:
            Conf.ErrorEmail()
            os._exit(0)
        try:
            manip.FileMove()
        except Exception as error:
            Conf.ErrorEmail()
            os._exit(0)
        try:
            manip.WriteToExcel()
        except Exception as error:
            Conf.ErrorEmail()
            os._exit(0)
        ## Start Bot
        try:
            bot = JobberBot()
            bot.login()
            try:
                bot.FieldPopulation()
            except Exception as error:
                Conf.ErrorEmail()
                os._exit(0)
        except Exception as error:
            Conf.ErrorEmail()
            os._exit(0)
    ##    # If Program makes it to here: success email will be sent!!
        try:
            Conf.SuccessEmail()
        except Exception as error:
            Conf.ErrorEmail()
            os._exit(0)
    else:
        Conf.WrongVersion()
        os._exit(0)
else:
    print('The latest document is not a .pdf or .docx')
    os._exit(0)
