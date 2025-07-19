import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains

from random_user_agent.user_agent import UserAgent
from random_user_agent.params import SoftwareName, OperatingSystem

import numpy as np
import pandas as pd
import time
import calendar
import yagmail
import datetime
import random
import os

class testCase(object):
	"""A drivetest test case"""
	def __init__(self, name,**kwargs):

		self.name 			  = name
		self.testType         = "G2"
		self.locations_gta    =["Toronto Etobicoke", "Oshawa", "Brampton", "Newmarket", "Burlington", "Mississauga", "Oakville",
								 "Toronto Downsview", "Toronto Metro East",
								"Toronto Port Union"]
		self.location         = self.locations_gta
		self.headless         = False
		self.incognito        = False
		self.sendEmail        = True
		self.test_html_dir    = None

	def _save_page_source(self, driver, file_path):
		"""Saves the page source to a file."""
		try:
			with open(file_path, 'w', encoding='utf-8') as f:
				f.write(driver.page_source)
		except Exception as e:
			print(f"Error saving page source to {file_path}: {e}")

	def sendEmail(self):
		print("Sending Email.......")

		receiver = self.emailAddress
		subject = "New timeslots found on drivetest.ca website"
		body = """\
		Hello

		A new timeslot is found on drivetest.ca website.
		Please book it as soons as possible.

		Best,
		Book.Road.Test.Online team

		"""
		yag = yagmail.SMTP(self.cloundEmailAddress)
		yag.send(
		    to=receiver,
		    subject=subject,
		    contents=body, 
		    attachments='./open_timeslots.csv'
		)
		print("Sending Email.......DONE")

	def bookARoadTest(self):
		# Create a timestamped directory for the test run
		timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
		self.test_html_dir = os.path.join("test_html", timestamp)
		os.makedirs(self.test_html_dir, exist_ok=True)

		self.url = "https://drivetest.ca/dtbookingngapp/registration/signup"

		# Let undetected-chromedriver handle the options
		driver = uc.Chrome()
		driver.get(self.url)
		self._save_page_source(driver, os.path.join(self.test_html_dir, "01_signup_page.html"))
		time.sleep(random.uniform(2, 4))

		# wait maximum 10 seconds for elements
		ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
		wait = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions)
		actions = ActionChains(driver)

		# Sign in
		time.sleep(random.uniform(2, 4))
		wait.until(EC.element_to_be_clickable((By.LINK_TEXT, 'Sign in'))).click()
		self._save_page_source(driver, os.path.join(self.test_html_dir, "02_confirm_registration_page.html"))
		time.sleep(random.uniform(2, 4))

		# Enter credentials
		time.sleep(random.uniform(1, 2))
		wait.until(EC.element_to_be_clickable((By.ID, 'driverLicenceNumber'))).send_keys(self.licenceNumber)
		time.sleep(random.uniform(0.5, 1.5))
		wait.until(EC.element_to_be_clickable((By.ID, 'driverLicenceExpiry'))).send_keys(self.expiryDate)
		time.sleep(random.uniform(0.5, 1.5))
		submit_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(., "Submit")]')))
		driver.execute_script("arguments[0].click();", submit_button)
		self._save_page_source(driver, os.path.join(self.test_html_dir, "03_view_bookings_page.html"))
		time.sleep(random.uniform(3, 5))

		# Book a new road test
		time.sleep(random.uniform(2, 4))
		book_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(., "Book a New Road Test")]')))
		driver.execute_script("arguments[0].click();", book_button)
		self._save_page_source(driver, os.path.join(self.test_html_dir, "04_select_licence_class_page.html"))
		time.sleep(random.uniform(3, 5))

		# Select license class
		time.sleep(random.uniform(1, 3))
		if self.testType == 'G2':
			wait.until(EC.element_to_be_clickable((By.XPATH, '//label[@for="licence_class_G2"]'))).click()
		elif self.testType == 'G':
			wait.until(EC.element_to_be_clickable((By.XPATH, '//label[@for="licence_class_G"]'))).click()
		time.sleep(random.uniform(0.5, 1.5))
		continue_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(., "CONTINUE")]')))
		driver.execute_script("arguments[0].click();", continue_button)
		self._save_page_source(driver, os.path.join(self.test_html_dir, "05_drivetest_locations_page.html"))
		time.sleep(random.uniform(3, 5))

		open_timeslots_data = []

		# Select locations
		locations_dir = os.path.join(self.test_html_dir, "Locations")
		os.makedirs(locations_dir, exist_ok=True)
		for location in self.location:
			print(f"Searching in location: {location}")
			location_dir = os.path.join(locations_dir, location)
			os.makedirs(location_dir, exist_ok=True)
			try:
				time.sleep(random.uniform(2, 4))
				# Use a more precise XPath to match the exact location name
				location_btn = wait.until(EC.presence_of_element_located((By.XPATH, f'//button[.//span[@class="locationName" and normalize-space()="{location}"]]')))
				actions.double_click(location_btn).perform()
				time.sleep(random.uniform(3, 5)) # Wait for page to load

				# Scrape calendar
				current_year = datetime.datetime.now().year
				for month in self.months:
					month_abbr = month[:3]
					month_number = list(calendar.month_abbr).index(month_abbr)
					month_dir = os.path.join(location_dir, month)
					os.makedirs(month_dir, exist_ok=True)
					
					# Navigate to the correct month
					while True:
						month_year_text = wait.until(EC.presence_of_element_located((By.ID, 'calendarMonthDiv'))).text
						displayed_month = month_year_text.split(' ')[0]
						if month.upper() == displayed_month.upper():
							self._save_page_source(driver, os.path.join(location_dir, f"{month}.html"))
							break
						wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-labelledby="next-label"]'))).click()
						time.sleep(random.uniform(0.5, 1.5)) # Brief pause for calendar to update

					# Find available dates
					available_dates_count = len(driver.find_elements(By.XPATH, '//button[contains(@class, "date-available")]'))
					for i in range(available_dates_count):
						# Re-find the elements on each iteration to avoid stale references
						available_dates = driver.find_elements(By.XPATH, '//button[contains(@class, "date-available")]')
						date_btn = available_dates[i]
						
						day = int(date_btn.text)
						date_btn.click()
						time.sleep(random.uniform(2, 4))
						calendar_continue_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[contains(., "CONTINUE")]')))
						driver.execute_script("arguments[0].click();", calendar_continue_button)
						
						# Scrape times
						self._save_page_source(driver, os.path.join(month_dir, f"{day}.html"))
						time.sleep(random.uniform(2, 4)) # Wait for times to load
						time_btns = driver.find_elements(By.XPATH, '//button[contains(@class, "widget")]')
						for time_btn in time_btns:
							hour = time_btn.text
							full_date = f"{current_year}-{month_number:02d}-{day:02d} {hour}"
							print(f"Found an open slot: {location}, {full_date}")
							open_timeslots_data.append({'Location': location, 'DateTime': pd.to_datetime(full_date)})
						
						driver.back()
						time.sleep(random.uniform(2, 4))
						wait.until(EC.presence_of_element_located((By.ID, 'calendarMonthDiv'))) # Wait for calendar to be back

			except Exception as e:
				print(f"Could not process location {location}: {e}")
			
			# Go back to location selection
			driver.get("https://drivetest.ca/dtbookingngapp/driveTestLocations")
			time.sleep(random.uniform(2, 4))


		# Process results
		if open_timeslots_data:
			open_timeslots = pd.DataFrame(open_timeslots_data)
			open_timeslots.sort_values(by='DateTime', inplace=True)
			open_timeslots['DateTime'] = open_timeslots['DateTime'].dt.strftime('%Y-%m-%d %I:%M %p')
			open_timeslots.to_csv('./open_timeslots.csv', index=False)

			print("\n--- All Available Bookings ---")
			print(open_timeslots)

			print("\n--- Soonest Available Bookings ---")
			print(open_timeslots.head(3))

			if self.sendEmail:
				self.sendEmail()
		else:
			print("No open timeslots found.")
			
		print(f"\nSearch completed at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

		# close the chrome driver window
		driver.quit()
