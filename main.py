from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas
import time

#Open webdriver
driver = webdriver.Chrome()  # Use the appropriate driver for your browser (e.g., Chrome, Firefox)
driver.get("https://egov.uscis.gov/")

#function to get status by entering the number and clicking the submit button
def getStatus(ticketNo):
    form = driver.find_element(By.ID, "receipt_number").send_keys(ticketNo)
    submit_button = driver.find_element(By.NAME,"initCaseSearch")
    submit_button.click()
    wait = WebDriverWait(driver, 20)  # Adjust the timeout as needed
    div_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "caseStatusSection")))
    time.sleep(2)
    status = div_element.find_element(By.ID, "landing-page-header").text
    return(status)

#function to extract the number from the ticket ID string
def ticketGenerator(ticket):
    ticketno = int(ticket[3:])
    return ticketno

#parameter input
firstTicket = "WAC2390054340"
increment = 5
numberOfRecords = 10

output = pandas.DataFrame(columns = ['Number','Status'])

#iterate through the list of tickets and retrieve status, save to dataframe
for i in range(0,numberOfRecords):
    caseNumber = "WAC" + str(ticketGenerator(firstTicket)+i*increment)
    status = getStatus(caseNumber)
    new_row = {"Number": caseNumber, "Status": status}
    output = pandas.concat([output, pandas.DataFrame([new_row])], ignore_index=True)
    print(i, caseNumber,status )

#output dataframe to excel
output.to_excel("output.xlsx", index=False)