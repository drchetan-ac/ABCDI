from selenium import webdriver
from selenium.webdriver.support.select import Select
import openpyxl
from selenium.common.exceptions import *
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import NoAlertPresentException
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.common.exceptions import NoSuchElementException
from datetime import date
from datetime import time
from datetime import datetime
from selenium.common.exceptions import TimeoutException
import xlrd
from json import dumps
from selenium.common.exceptions import StaleElementReferenceException
from easygui import *
from src.config import getConfig   # ModuleNotFoundError: No module named 'config fixed'

from src.config import AppConfiguration   # ModuleNotFoundError: No module named 'config fixed
from src.auth import clickonlogin, askforlogincred
import logging
#import error for sheetactivity fixed
# from src.sheetActivity import sheetActivity
from src.sheetActivity import * # opensourcefile, opensrfid, fetchrowvalues, openBysearch, entersampleid, entersamplety, enterdates, entertestkit, enterResults, finalsubmit, enterPatientid
import os
import configparser

cwd = os.getcwd()
#config_Path= cwd + "\\Config.txt"
#configuration = configparser.ConfigParser(
#if path.exists(config_Path):

appConfig = AppConfiguration()
config = appConfig.getConfig()



def entryIntoSheet(driver,age, age_in, Gender, phone, sampleId, sampleTy,
                   sampleRDate, sampleTDate, sampleCDate,testKit, Egen,
                   SRFid, CtEgene, ORF1a, CtORF, RDRP_SGene, CtRDRP, fiResult,
                   ArogyaSetuDownload, address, pincode, State, Dist, nationality,
                   hospitalization, sampleCollectedFrom, symptomstatus, Mobilebelongsto,
                   patientname, occupation, patientId ):
    if not "followup_entries" in driver.current_url:
        stateDist(driver, State, Dist, 0) # page refreshes if district does not load, hence this should be in the first
        enterNationality(driver, nationality)
        checkNenterage_in(driver, age_in)
        enterAge(driver, age)
        enterGender(driver, Gender)
        enterPhonenum(driver, phone)
        enterMobilebelongsto(driver, Mobilebelongsto)
        enterpatientName(driver, patientname)
        aarogyasetuApp(driver, ArogyaSetuDownload)
        enterAddress(driver, address)
        enterPincode(driver, pincode)
        enterPatientid(driver, patientId, SRFid)
        patientCategoryselect(driver)

    ptOccupation(driver,occupation)
    modeOftransport(driver)
    sampleCollectedfrom(driver, sampleCollectedFrom)
    hospitalizationEntry(driver, hospitalization)
    symptomStatus(driver, symptomstatus)
    entersampleid(driver, sampleId)
    entersamplety(driver, sampleTy)
    enterdates(driver, sampleRDate, sampleTDate, sampleCDate)
    entertestkit(driver, testKit)
    enterResults(driver, Egen, SRFid, CtEgene, ORF1a, CtORF, RDRP_SGene, CtRDRP, fiResult)



def main_run():
    getConfig()
    #driver = login()
    currentrow, sheet = opensourcefile()
    counter = 0
    sampleCollectedFrom = "Point of entry"
    symptomstatus = "Asymptomatic"
    Mobilebelongsto = "Patient"
    while currentrow > 1:
        #if not counter == 0:
        if ("driver" in locals()) or ("driver" in globals()):
            if driver:
                driver.close()
        driver = login()
        counter = 0
        while counter < 300 and currentrow > 1:
            #SRFid = sheet.cell(row=currentrow, column=2).value
            #Check if the SRF ID is blank and move to next row, else continue
            if sheet.cell(row=currentrow, column=2).value is None:
                currentrow = currentrow-1
                continue

            srfidopening, textalert, SRFid = opensrfid(sheet, currentrow, driver)

            if srfidopening == "no SRF ID":
                currentrow = currentrow - 1
                continue
            elif srfidopening == "Alert!!":
                print(SRFid, "Check if this has to be entered by search")

                if config["ProceedWithRATfollowup"] == 'Yes':
                    if not config["labUniqueAlert"] in textalert:
                        # + patientname
                        if boolbox('SRF ID is fetched by another lab, Do you want to search and enter?', 'SRF ID is fetched by another lab, Do you want to search and enter? Click Yes or No',
                                   ('[<F1>]Yes', '[<F2>]No'), None, '[<F1>]Yes', '[<F2>]No', ):
                            logging.info(str(
                                SRFid) + " is already taken into ICMR Portal, Continuing with followup entry by search" + '\n')
                            print(str(
                                SRFid) + " is already taken into ICMR Portal, Continuing with followup entry by search" + '\n')
                            #fetchrowvalues(sheet, currentrow)
                            sampleTy, sampleId, sampleRDate, sampleTDate, testKit, Egen, ORF1a, RDRP_SGene, fiResult, CtEgene, CtORF, CtRDRP,\
                            labId, phone, ICMRHqID, State, Dist, patientname, patientId, hospitalization, sampleCDate, \
                            nationality, age, age_in, Gender, ArogyaSetuDownload, address, pincode, category, occupation, Mobilebelongsto,\
                            symptomstatus, sampleCollectedFrom, fathersName = fetchrowvalues(sheet, currentrow)
                            #searchopenok = openBysearch(driver)
                            searchopenok =openBysearch(driver, State, Dist, ICMRHqID, patientId, phone, patientname, Gender, age,
                                         age_in)
                            if not searchopenok:
                                logging.info(str(
                                    SRFid) + " is not searched & Selected, Continuing with next entry" + '\n')
                                print(str(
                                    SRFid) + " is not searched & Selected, Continuing with next entry" + '\n')
                                #currentrow = currentrow - 1
                                #counter = counter + 1
                                #continue
                            else:
                                entryIntoSheet(driver, age, age_in, Gender, phone, sampleId, sampleTy,
                                               sampleRDate, sampleTDate, sampleCDate, testKit, Egen,
                                               SRFid, CtEgene, ORF1a, CtORF, RDRP_SGene, CtRDRP, fiResult,
                                               ArogyaSetuDownload, address, pincode, State, Dist, nationality,
                                               hospitalization, sampleCollectedFrom, symptomstatus, Mobilebelongsto,
                                               patientname, occupation)
                                '''
                                checkNenterage_in(driver, age_in)
                                enterAge(driver, age)
                                enterGender(driver, Gender)
                                enterPhonenum(driver, phone)
                                ptOccupation(driver)
                                modeOftransport(driver)
                                sampleCollectedfrom(driver)
                                entersampleid(driver, sampleId)
                                entersamplety(driver, sampleTy)
                                enterdates(driver, sampleRDate, sampleTDate, sampleCDate)
                                entertestkit(driver, testKit)
                                enterResults(driver, Egen, SRFid, CtEgene, ORF1a,CtORF, RDRP_SGene, CtRDRP, fiResult)
                                '''
                                textalert1 = None
                                submitstatus,textalert1= finalsubmit(driver, fiResult, SRFid)
                                if submitstatus:
                                    print(SRFid, "Submitted sucessfully", '\n') #, '...... end of')
                                    logging.info("Submitted sucessfully") # + '...... End of  ' + str(SRFid) + "}")
                                    #currentrow = currentrow - 1
                                    # continue
                                else:
                                    print("SRF ID", SRFid, " has an alert")
                                    logging.info(str(SRFid) + " Has an alert" + '\n' + textalert1 + '\n')
                                    textbox(msg=str(SRFid) + " Has an alert" + '\n' + textalert1 + '\n' + '\n' + '\n' + "Skipping over to the next SRF ID",
                                            title='Alert!', text='', codebox=False, callback=None, run=True)

                            #currentrow = currentrow - 1
                            #counter = counter + 1
                            #continue
                            #print(str(SRFid), "...... end of")
                            #logging.info('...... End of  ' + str(SRFid) + "}")

                        else:
                            print(
                                SRFid, "is taken up by another lab, moving on to next entry")
                            logging.info(str(
                                SRFid) + " is taken up by another lab, moving on to next entry")
                            #logging.info('...... End of  ' + str(SRFid) + "}")
                            #currentrow = currentrow - 1
                            #counter = counter + 1
                            #continue

                    else:
                        logging.info(
                            str(SRFid) + " is already entered into ICMR Portal by this lab, skipping to next row ")
                        logging.info(textalert)
                        print(
                            SRFid, " is already entered into ICMR Portal by this lab, skipping to next row")
                        #logging.info(
                        #    '...... End of  ' + str(SRFid) + "}")
                        #currentrow = currentrow - 1
                        #counter = counter + 1
                        #continue
                else:
                    logging.info(
                        str(SRFid) + " is already taken into ICMR Portal, please check if this is followup of RAT testing, skipping to next row ")
                    logging.info(textalert)
                    print(
                        SRFid, " is already taken into ICMR Portal, please check if this is followup of RAT testing, skipping to next row ")

                #print(SRFid, "...... end of")
                #logging.info('...... End of  ' + str(SRFid) + "}")
                #currentrow = currentrow - 1
                #counter = counter + 1
                #continue
            else:
                sampleTy, sampleId, sampleRDate, sampleTDate, testKit, Egen, ORF1a, RDRP_SGene, fiResult, CtEgene, CtORF, CtRDRP, \
                labId, phone, ICMRHqID, State, Dist, patientname, patientId, hospitalization, sampleCDate, \
                nationality, age, age_in, Gender, ArogyaSetuDownload, address, pincode, category, occupation, Mobilebelongsto, \
                symptomstatus, sampleCollectedFrom, fathersName = fetchrowvalues(sheet, currentrow)
                entryIntoSheet(driver, age, age_in, Gender, phone, sampleId, sampleTy,
                               sampleRDate, sampleTDate, sampleCDate, testKit, Egen,
                               SRFid, CtEgene, ORF1a, CtORF, RDRP_SGene, CtRDRP, fiResult,
                               ArogyaSetuDownload, address, pincode, State, Dist, nationality,
                               hospitalization, sampleCollectedFrom, symptomstatus, Mobilebelongsto,
                               patientname, occupation, patientId)
                '''
                checkNenterage_in(driver, age_in)
                enterAge(driver, age)]
                enterGender(driver, Gender)
                enterPhonenum(driver, phone)
                ptOccupation(driver)
                modeOftransport(driver)
                sampleCollectedfrom(driver)
                # mode_of_transport
                entersampleid(driver, sampleId)
                entersamplety(driver, sampleTy)
                enterdates(driver, sampleRDate, sampleTDate, sampleCDate)
                entertestkit(driver, testKit)
                enterResults(driver, Egen, SRFid, CtEgene, ORF1a, CtORF, RDRP_SGene, CtRDRP, fiResult)
                '''

                # click on the Submit button, check for any alerts and print values
                textalert1 = None
                #textalert1 = ''
                submitstatus, textalert1 = finalsubmit(driver, fiResult, SRFid)
                if submitstatus:
                    print(SRFid, "Submitted sucessfully", '\n')    #, '...... end of')
                    logging.info("Submitted sucessfully" + '\n')   #+                           '...... End of  ' + str(SRFid) + "}")
                    #continue
                else:
                    if textalert1 != None:
                        print("SRF ID", SRFid, " has an alert")
                        logging.info(
                            str(SRFid) + "Has an alert" + '\n' + textalert1 + '\n')
                        textbox(msg=str(
                            SRFid) + " Has an alert" + '\n' + textalert1 + '\n' + '\n' + '\n' + "Skipping over to the next SRF ID",
                                title='Alert! - Submit text box', text='', codebox=False,
                                callback=None, run=True)
                        #continue
                    else:
                        logging.info(
                            str(SRFid) + " May have failed, please check" + '\n' + '\n')
                        textbox(msg=str(SRFid) + " May have failed, please Check" + '\n' + "Skipping over to the next SRF ID",
                                title='Alert! - Submit text box', text='', codebox=False,
                                callback=None, run=True)
                #submitpage = finalsubmit(driver)
            #update the counters and close the SRF, move to next SRF
            currentrow = currentrow - 1
            counter = counter + 1
            print('...... End of  ' + str(SRFid) + "}")
            logging.info('...... End of  ' + str(SRFid) + "}")
            continue

    if driver:
        driver.close()
    logging.info("End time of execution:" +
                 datetime.now().strftime('%d_%m_%Y__%H_%M_%S_%f') + '\n')


if __name__ == "__main__":
    main_run()