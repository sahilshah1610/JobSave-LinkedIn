from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, NoSuchAttributeException
from selenium.common.exceptions import *
from selenium.webdriver.common.action_chains import ActionChains
import time
import re
import json
import pandas as pd
import random
import string
import os

class JobAlertsSaving:

    def __init__(self, data):

        self.email = data["email"]
        self.password = data["password"]
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("start-maximized")
        self.options.add_argument("disable-infobars")
        self.options.add_argument("--disable-extensions")
        self.driver = webdriver.Chrome(chrome_options=self.options,executable_path=data["driver_path"])
        self.driver.maximize_window()
        self.keyword = data['keyword']
        self.location = data['location']
        self.companyNameList = list()
        self.companyStrengthList = list()
        self.companyTypeList = list()
        self.companyPositionList = list()
        self.companyLocationList = list()
        self.companyDatePostedList = list()
        self.companyURLList = list()


    def createExcel(self):
		# Creating excel file for all the data collected and storing it in /excels folder
        dataFile = pd.DataFrame({'Company Name':self.companyNameList,
                                 'Job Position':self.companyPositionList,
                                 'Company Strength':self.companyStrengthList,
                                 'Location':self.companyLocationList,
                                 'Company Type':self.companyTypeList,
                                 'Job URL':self.companyURLList,
                                 'Date posted':self.companyDatePostedList})
        dataFile['Status'] = ''
        ranString = self.getRandomString()
        fileName = self.keyword + '_' + self.location + '_' +ranString+'.xlsx'
        writer = pd.ExcelWriter(os.path.abspath('excels/'+fileName),engine='xlsxwriter')
        dataFile.to_excel(writer, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:C', 30)
        worksheet.set_column('D:D', 30)
        worksheet.set_column('E:E', 20)
        worksheet.set_column('F:F', 28)
        worksheet.set_column('G:G', 20)
        worksheet.set_column('H:H', 20)
        worksheet.set_column('I:I', 20)
        writer.save()
        print("{} file is being created".format(fileName))
        self.driver.close()



    def getRandomString(self):
        letters = string.ascii_lowercase
        result = ''.join(random.choice(letters) for i in range(5))
        return result

    def login(self):
        # login into linkedin.com and entering email and pass
        self.driver.get("https://www.linkedin.com/login")
        loginEmail = self.driver.find_element_by_name("session_key")
        loginEmail.clear()
        loginEmail.send_keys(self.email)
        loginPass = self.driver.find_element_by_name("session_password")
        loginPass.clear()
        loginPass.send_keys(self.password)
        loginPass.send_keys(Keys.RETURN)
        print("Logged in to LINKEDIN")

    def jobSearch(self):
        # clicking jobs tab and entering keyword and location for jobs
        jobLink = self.driver.find_element_by_link_text("Jobs")
        jobLink.click()
        time.sleep(5)
        print("Jobs Tab Opened")

        searchKeywords = self.driver.find_element_by_xpath('//input[starts-with(@id,"jobs-search-box-keyword")]')
        searchKeywords.clear()
        searchKeywords.send_keys(self.keyword)
        time.sleep(8)

        searchLocation = self.driver.find_element_by_xpath('//input[starts-with(@id,"jobs-search-box-location")]')
        searchLocation.clear()
        searchLocation.send_keys(self.location)
        searchLocation.send_keys(Keys.RETURN)
        print("keyword and location entered")


    def applyingFilter(self):
        # Applying filters
		
		# Selecting all filters button
        allFiltersButton = self.driver.find_element_by_xpath('//button[@data-control-name="all_filters"]')
        allFiltersButton.click()
        time.sleep(5)

		# Selecting entry level button
        experienceLevelButton = self.driver.find_element_by_xpath('//label[@for="experience-2"]')
        experienceLevelButton.click()
        time.sleep(2)

		# Selecting sort by as most recent
        sortByButton = self.driver.find_element_by_xpath('//label[@for="sortBy-DD"]')
        sortByButton.click()
        time.sleep(2)
       
		# Clicking on Apply button to apply all the filters
        applyButtonClick = self.driver.find_element_by_class_name('search-advanced-facets__button--apply.ml4.mr2.artdeco-button.artdeco-button--3.artdeco-button--primary.ember-view')
        applyButtonClick.click()
        time.sleep(8)


    def availableJobs(self):

        totalOffers = self.driver.find_element_by_class_name('display-flex.t-12.t-black--light.t-normal')
        totalOffersCount = int(totalOffers.text.split(' ', 1)[0].replace(",",""))
        print("Total jobs as per filters -> ",totalOffersCount)
        time.sleep(7)
        currentURL = self.driver.current_url
        results = self.driver.find_elements_by_class_name("jobs-search-results__list-item.occludable-update.p0.relative.ember-view")
        for result in results:
            hoverAction = ActionChains(self.driver).move_to_element(result)
            hoverAction.perform()
            companyTitles = result.find_elements_by_class_name("full-width.artdeco-entity-lockup__title.ember-view")
            for title in companyTitles:
                title.click()
                self.retrieveInfo(title)

        if totalOffersCount > 24:
            time.sleep(4)

            findTotalPages = self.driver.find_elements_by_class_name('artdeco-pagination__indicator.artdeco-pagination__indicator--number.ember-view')[-1]
            totalPages = findTotalPages.text
            totalPagesCount = int(re.sub(r"[^\d.]", "", totalPages))
            print("total pages are ",totalPagesCount)
            getLastPage = self.driver.find_element_by_xpath("//button[@aria-label='Page "+str(totalPagesCount)+"']")
            getLastPage.send_keys(Keys.RETURN)
            time.sleep(4)
            lastPageURL = self.driver.current_url
            totalJobs = int(lastPageURL.split('start=',1)[1])

            for pageNumber in range(25, totalJobs+25, 25):
                self.driver.get(currentURL + '&start=' +str(pageNumber))
                time.sleep(4)
                resultsExt = self.driver.find_elements_by_class_name("jobs-search-results__list-item.occludable-update.p0.relative.ember-view")
                for eachResult in resultsExt:
                    hoverAction = ActionChains(self.driver).move_to_element(eachResult)
                    hoverAction.perform()
                    companyTitles = eachResult.find_elements_by_class_name(
                        "full-width.artdeco-entity-lockup__title.ember-view")
                    for title in companyTitles:
                        title.click()
                        self.retrieveInfo(title)
        self.createExcel()

    def retrieveInfo(self, jobClicked):

        time.sleep(5)

		# Extracting Company industry type and strength
        try:
            companyDiv = self.driver.find_element_by_class_name('artdeco-list__item.jobs-details-job-summary__section.jobs-details-job-summary__section--center')
            companyULTag = companyDiv.find_element_by_tag_name('ul')
            companyDetails = companyULTag.find_elements_by_class_name('jobs-details-job-summary__text--ellipsis')
            companyIndustry = companyDetails[1].text
            companyStrength = companyDetails[0].text

        except :
            companyIndustry = 'No data'
            companyStrength = 'No data '

        if companyIndustry == "":
            companyIndustry = 'No data'
        if companyStrength == "":
            companyStrength = 'No data'

        self.companyTypeList.append(companyIndustry)
        self.companyStrengthList.append(companyStrength)

		# Extracting company Name and position
        try:
            companyNameDiv = self.driver.find_element_by_class_name('jobs-details-top-card__content-container')
            companyNamePositionTag = companyNameDiv.find_element_by_tag_name('h2')
            companyPositionText = companyNamePositionTag.text
            companyNameTagA = companyNameDiv.find_elements_by_tag_name('a')[1]
            companyNameText = companyNameTagA.text
        except :
            companyPositionText = ' '
            companyNameText = ' '


        if companyPositionText == "":
            companyPositionText = 'No data'
        if companyNameText == "":
            companyNameText = 'No data'

        self.companyPositionList.append(companyPositionText)
        self.companyNameList.append(companyNameText)

		# Extracting company location
        companyLocationDiv = self.driver.find_element_by_class_name('jobs-details-top-card__content-container')
        companyLocationClass = companyLocationDiv.find_element_by_class_name('jobs-details-top-card__bullet')
        companyLocationText = companyLocationClass.text
        if companyLocationText == "":
            companyLocationDiv = self.driver.find_element_by_class_name(
                'jobs-details-top-card__company-info.t-14.t-black--light.t-normal.mt1')
            companyLocationClass = companyLocationDiv.find_element_by_class_name(
                "jobs-details-top-card__exact-location.t-black--light.link-without-visited-state")
            companyLocationText = companyLocationClass.text

        self.companyLocationList.append(companyLocationText)

        # Extracting date posted value [Working on it]
        companyDatePostedText = 'No data'
        self.companyDatePostedList.append(companyDatePostedText)
        currentJobUrl = self.driver.current_url
        self.companyURLList.append(str(currentJobUrl))
        


if __name__ == "__main__":
    with open("credentials.json") as config:
        data = json.load(config)

    objJobAlertsSaving = JobAlertsSaving(data)
    objJobAlertsSaving.login()
    time.sleep(8)
    objJobAlertsSaving.jobSearch()
    time.sleep(8)
    objJobAlertsSaving.applyingFilter()
    time.sleep(5)
    objJobAlertsSaving.availableJobs()
