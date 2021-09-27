# ABCDI
ABCDI: Automated BOT for Covid test results Data entry into ICMR portal

Developed by: Dr. Chetan A C, Scientist B, ICMR National Institue of Traditional Medicine
Email: chetan.ac.icmr@gmail.com

ABCDI is an desktop bot for automatic entry of COVID RTPCR test results into ICMR portal, https://cvstatus.icmr.gov.in, which is mandatory for COVID tests done in India as per Government of India and State Government orders. The bot is based on Selenium-Python framework. The program is designed to read the "login.xlsx" for login credentials & configuraiton details, and the data source in a preset format for data source. Program triggers the web drivers of Firefox or Chrome as per the configuraiton, opens the COVID data portal of ICMR, logs in with the credentials provided, searches the srf ID as provided, starting from the last entry in the data source excel sheet, and updates information as provided, followed by submission, before moving on to the next entry. 

The bot runs locally from the desktop computer, without sharing data elsewhere apart from the ICMR covid portal. 

Instructions for use:
Typically, the data of samples is received along with samples or generated from NIC's SRF app, which is filled during sample collection.
The collated information should be then updated after the samples are processed and tested for RT-PCR.
An excel sheet of the data should be prepared in the format as in "example data feed.xlsx".
Update the "logins.xlsx" file with login credentials of the lab. 
ABCDI is distributed freely under GNU AGPL license.

Update the appropriate drivers in /src/drivers/ as per the version of the installed browser. Firefox is recommended, for which gecko drivers are included, but may need to be updated as required in the host system.

Contents of "login_file.xlsx" in sheet login_cred
1. URL	https://cvstatus.icmr.gov.in
2. User_Name
3. Password	Login
4. prefix	Local_Lab
5. ProceedWithRATfollowup	NO
6. Unique Text for the lab - already fetched alert	"Lab Name:- Contact Person:- Email:- "
7. Geckodriver path	C:\\entryBot\\geckodriver.exe
8. ChromeDriver Path	C:\\entryBot\\chromedriver.exe
9. Default Driver	firefox

Contact: 
Dr. Chetan A C
Email: chetan.ac.icmr@gmail.com

Acknowledgements:
1. ICMR NITM COVID testing lab, especially the data entry staff for providing crucial inputs to develop this software.
2. Nodal officer for COVID testing lab at ICMR NITM, Dr. Banappa Unger, Scientist E
3. Dr. Debprasad Chattopadhyay, Director, ICMR NITM for his support.

