# local library
from classes import testCase
import config


"""
The process:
- Enter into website
- Input relevant information
- Select test centers to check (may be multiple)
    - Parse to find free date
        - if found store relevant information
- Send an email about the dates
"""

"""
crontab command can schedule the scraping in unix-based systems. for instance:
$ crontab -e
and adding the following line
*/30 * * * * /usr/bin/python path/to/driveTestscraper.py
to the crontab file will run the driveTestscraper.py script every 30 minutes.
"""

tc = testCase('myTest')

tc.licenceNumber       = config.licenceNumber
tc.expiryDate          = config.expiryDate
tc.emailAddress        = config.emailAddress
tc.testType            = "G"
tc.location    		   = ['Renfrew','Smiths Falls','Pembroke','Brockville','Ottawa Walkley','Ottawa Canotek','Cornwall','Hawkesbury','Kingston']
tc.months              = ['July','August','September']

# options
tc.write        	   = True   # write available dates into a csv file?
tc.headless 		   = False	# launch headless chrome driver? 
tc.verify	 		   = False   # for the first time booking, set to false. for reschedeling set to True.
tc.sendEmail           = False   # send the available dates to the applicant's email. if True, a cloud email address is needed to send the email. 
tc.cloundEmailAddress  = "cloud-email@gmail.com"


tc.bookARoadTest()
