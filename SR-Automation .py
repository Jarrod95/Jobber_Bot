import openpyxl, pprint, os, glob, time, shutil, re
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
os.chdir('Set to desired folder')

#DOB Calculation (to be used in data extraction )
def calculate_age(born):
    today = date.today()
    return today.year - born.year - ((today.month, today.day) < (born.month, born.day))

class DataExtraction():
    def checkFolder(self):
        #Latest File in Directory - MIGHT NEED TO BE CHNAGED TO THE OLDEST FILE IN DIRECTORY IF RUNNING ON SCHED
        #NEED TO ADD IN FUNCTIONALITY FOR RESTRICTED .DOCX FILES &
        #.PDF versions of the document 
        list_of_files = glob.glob(os.getcwd() + '\\*') 
        last_file = max(list_of_files, key=os.path.getctime)
        #Just get the file name
        global document, latest_file, read_pdf
        latest_file = last_file.split('\\')[-1]
        # print(latest_file)
     
    def ExtractDocxData(self):
        # NEED TO SIMPLIFY THIS -- WORK SMARTER NOT HARDER -- NESTED LOOP?
        # Table 1
        global provider, manager, loc, m_phone, email, name, name_fn, name_ln,\
               dob, address, street1, city, state, number, med_detail, overview,\
               nok, nok_contact, nok_rel, start, reason, date_of_birth, age,\
               prim_contact, type1, freq, postcode
        table = document.tables[0]
        table1 = []
        for row in table.rows:
            for cell in row.cells:
                table1.append(cell.text)

        # Table 2
        table = document.tables[1]
        table2 = []
        for row in table.rows:
            for cell in row.cells:
                table2.append(cell.text)
                
        # Table 3
        table = document.tables[2]
        table3 = []
        for row in table.rows:
            for cell in row.cells:
                table3.append(cell.text)
                
        # Table 4
        table = document.tables[3]
        table4 = []
        for row in table.rows:
            for cell in row.cells:
                table3.append(cell.text)

        # Company/ Case Manager details
        provider = table1[5]
        manager = table1[8]
        loc = table1[11]
        m_phone = table1[14]
        email = table1[-1]

        # Get Client Details 
        name = table2[3].split()
        name_fn = name[0]
        name_ln = name[1]
        dob = table2[5]
        address = table2[7].split()
        street1 = street1 = ' '.join([str(elem) for elem in address[:3]])
        city = address[3]
        state = re.sub('[\W_]+', '', address[-2])
        number = table2[9]
        med_detail = table2[11]
        overview = table2[13]
        nok = table3[5]
        nok_contact = table3[11]
        nok_rel = table3[13]
        start = table3[-6]
        reason = table3[-1]
        # Use function to get Age
        date_of_birth = datetime.strptime(dob, "%d/%m/%Y")
        # def calculate_age(born):
        #     today = date.today()
        #     return today.year - born.year - ((today.month, today.day) < (born.month, born.day))
        age = calculate_age(date_of_birth)

        # Work out how to get radio buttons, find state & PC
        prim_contact = ''
        type1 = ''
        freq = ''
        postcode = ''

    def ExtractPdfData(self):
        global provider, manager, loc, m_phone, email, name, name_fn, name_ln,\
               dob, address, street1, city, state, number, med_detail, overview,\
               nok, nok_contact, nok_rel, start, reason, date_of_birth, age,\
               prim_contact, type1, freq, postcode
        page = read_pdf.getPage(0).extractText()
        #remove unecessary elements and empty spaces
        page2 = page.split('\n')
        while(' ' in page2):
                page2.remove(' ')
        ##page2 = list(filter(lambda a: a != '\n', page))
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
        address = page2[36].split(' ')
        street1 = street1 = ' '.join([str(elem) for elem in address[:3]])
        city = address[3]
        state = re.sub('[\W_]+', '', address[-2])
        postcode = address[-1]
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

        # Work out how to get radio buttons, find state & postcode if missing
        prim_contact = ''
        type1 = ''
        freq = ''
                
    def ExtractBilling(self):
        ## SEARCH BILLING INFORMATION SPREADSHEET
        ## WORK OUT HOW TO CHECK IF THE CLIENT IS NOT IN SPREADSHEET AT ALL
        global billing
        df = pd.read_excel('Pull billing details from assigned spreadsheet')
        res1 = df[df['Company'].str.contains(provider, na=False)]
        billing = []
        if len(res1.index) > 1:
            billing.append(df[df['Company'].str.contains(loc, na=False)])
        else:
                billing = res1

class FileManip():
    def FileMove():
        global new_folder
        new_folder = 'C:\\Users\\info\\OneDrive\\1. M2M Administration\\AGED CARE\\AGED CARE CLIENT PROVIDERS\\'+ \
          state + '\\' + provider + '\\' + name_fn + ' ' + name_ln
        Path(new_folder).mkdir(parents=True, exist_ok=True)
        #Rename file and move to directory under the right name
        os.rename(last_file, new_folder + '\\' + name_fn + ' ' + name_ln + ' Service Request - ' +\
                date.today().strftime("%d-%m-%Y") + '.docx')

    def WriteToExcel():
        ##Write name to cells -- 
        ## CHECK IF ENTRY IS DUPLICATED -- DO NOT LET DUPLICATED DATA PASS ON TO .XLSX FILE AND CRM
        wb = openpyxl.load_workbook('SPECIFIED BACKEND .XLSX FILE')
        sheet = wb['Sheet1']
        ## For Loop to insert new data below existing data
        details = [[datetime.now().strftime("%d/%m/%Y %H:%M:%S"), name_fn, name_ln, provider, number,\
                    street1, city, state, postcode, dob, age, prim_contact, type1, freq, med_detail, overview, nok, nok_contact,\
                    nok_rel, start, reason]]
        for detail in details:
            sheet.append(detail)
        wb.save('SPECIFIED BACKEND .XLSX FILE')

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

    def FieldPopulation(self):
        ## FIELDS TO FILL
        #Name
        self.driver.find_element_by_name('client[first_name]').send_keys(name_fn)
        self.driver.find_element_by_name('client[last_name]').send_keys(name_ln)
        self.driver.find_element_by_xpath('//*[@id="client_company_name"]').send_keys(provider)
        #Uncheck 'Billing address is the same...' checkbox
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[1]/label/sg-icon').click()
        time.sleep(1)
        #Enter Billing Details 
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[2]/div/div[1]/div/placeholder-field/input').send_keys(str(billing['Email']))
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[2]/div/div[2]/div/placeholder-field/input').send_keys(str(billing['Billing Street 1']))
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[2]/div/div[2]/div/placeholder-field/input').send_keys(str(billing['Billing Street 2']))
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[2]/div/div[3]/div[1]/placeholder-field/input').send_keys(str(billing['Billing City']))
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[2]/div/div[3]/div[2]/placeholder-field/input').send_keys(str(billing['Postcode']))
        #Select Australia from dropdown
        select_fr = Select(self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[2]/div[4]/div/div[2]/div/div[4]/div[2]/div/select'))
        select_fr.select_by_index(0)
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
        ## Dropdown the 'Additional Client details' menu
        self.driver.find_element_by_xpath('//*[@id="new_client"]/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-header/div/div[2]/div/div').click()
        time.sleep(1.5)
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
        #Additional Notes
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form/div/div/div/div[1]/div[1]/sg-accordion/sg-accordion-section[2]/sg-accordion-section-body/div[1]/placeholder-field[14]/input').send_keys(overview)
        time.sleep(2)
        # CREATE CLIENT
        self.driver.find_element_by_xpath('//*[@id="new_client"]/div/div/div/div[2]/div/div/div[2]/button').click()
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
        ## NEED TO WORK OUT HOW TO GET FREQ FROM WORD
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form[1]/div/div/div/div[1]/div/div[2]/div[1]/div[2]/div[1]/div/placeholder-field/input').send_keys('M')
        #Check 'Schedule Later
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form[1]/div/div/div/div[2]/div[2]/div/div[1]/div/div[1]/div/div[1]/div/div/label/sg-icon').click()
        #Save Job
        self.driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[2]/div[1]/div[2]/div/form[1]/div/div/div/div[2]/div[5]/div/button').click()
        print(f'Successfully created an unassigned job for {name_fn} {name_ln} in Jobber.')

class ConfEmail():
    def SuccessEmail(self):
        ## SEND AN EMAIL TO NOTIFY STAFF
        debuglevel = 0

        smtp = SMTP()
        smtp.set_debuglevel(debuglevel)
        smtp.connect('SERVER', 587)
        smtp.login('EMAIL ADDRESS', 'PASSWORD')

        from_addr = "Automated Bot <EMAIL ADDRESS>"
        to_addr = "EMAIL ADDRESS"

        subj = F'SYNC COMPLETED FOR {name_fn} {name_ln}'
        date = datetime.now().strftime( "%d/%m/%Y %H:%M" )
#need to add {new folder} back into conf email when ready 
        message_text = f'Hello,\n\n\
        A service request for {name_fn} {name_ln} has been recieved, processed\
        and moved to the following file location:\n\n\
        {new_folder}\n\n\n\
        This task was completed at {date}\n\n\n\
        Kind regards,\n\n\
        Your Friendly Automated Bot.'

        msg = "From: %s\nTo: %s\nSubject: %s\nDate: %s\n\n%s" \
            % ( from_addr, to_addr, subj, date, message_text )

        smtp.sendmail(from_addr, to_addr, msg)
        smtp.quit()    

#Initialize prog to check folder
#Start building a structured conditional list of actions for program
#i.e. if x fails, program has done y so far - passed on into an email
DataExt = DataExtraction()
Conf = ConfEmail()
DataExt.checkFolder()
if latest_file == None:
    print('There are no new documents in the folder. Program will not be run.')
    exit()
elif '.pdf' in latest_file:
    read_pdf = read_pdf = PyPDF2.PdfFileReader(latest_file)
    DataExt.ExtractPdfData()
    print('Latest file is a .pdf')
elif '.docx' in latest_file:
    document = Document(latest_file)
    DataExt.ExtractDocxData()
    print('Latest file is a .docx')
else:
    print('The latest document is not a .pdf or .docx')
    
DataExt.ExtractBilling()

##bot = JobberBot()
##bot.login()