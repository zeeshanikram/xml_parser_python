import xlrd
import subprocess
import random
import urllib, urllib2, base64
import datetime, time

###	VPN Credentials and vpn iplist
iplist = 'iplist.txt'
username = ''
password = ''

########################################################
###########	 Change the DATES below 	################

NEXT_PAY_DATE_CURRENT = '09/11/2012'
NEXT_PAY_DATE_BIWEEKLY = '09/25/2012'
NEXT_PAY_DATE_WEEKLY = '01/30/2011'
NEXT_PAY_DATE_MONTHLY = '01/25/2011'
NEXT_PAY_DATE_SEMI_MONTHLY = '01/15/2011'
TRACKING_ID = 'yum20'		## Put your track ID here
SLEEP_TIME	= 5

#############################################################

current_month_list = ['02', '03', '04']
current_year_list = ['2006', '2007']

def get_random_ip(iplist):
	random_line = random.choice(open(iplist).readlines())
	return random_line.strip()

def connect_vpn(username, password):
	while True:
		ip = get_random_ip(iplist)
		print "Dialing VPN: " + ip
		try:
			result = subprocess.check_output('rasdial hmavpn ' + username + ' ' + password +' /phone:' + ip, stderr=subprocess.STDOUT, shell = True)
			print result
			if result.find('Successfully connected') != -1:
				break
			if result.find('already connected') != -1:
				disconnect_vpn()
		except subprocess.CalledProcessError, e:
			print "subproces CalledProcessError.output = " + e.output 
		


def disconnect_vpn():
	print "Disconnecting VPN"
	subprocess.call('rasdial /disconnect')

def minimalist_xldate_as_datetime(xldate, datemode):
    # datemode: 0 for 1900-based, 1 for 1904-based
    return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=xldate + 1462 * datemode))

wb = xlrd.open_workbook('file.xls')
#Check the sheet names

wb.sheet_names()
#Get the first sheet either by index or by name

sh = wb.sheet_by_index(0)
sh = wb.sheet_by_name(u'Sheet1')
#Iterate through rows, returning each as a list that you can index:


for rownum in range(sh.nrows):
	if rownum == 0: continue
	while True:
		####### 	Connect VPN   ###########
		#connect_vpn(username, password)
		time.sleep(SLEEP_TIME)

		###############################################
		loan_count = '0'	# Always zero
		loan_amount = str(round(sh.cell(rownum,0).value)).rstrip('0').rstrip('.')
		loan_amount = loan_amount.strip()
		firstname = sh.cell(rownum,1).value
		lastname = sh.cell(rownum,2).value
		email = sh.cell(rownum,3).value
		email = email.strip()
		address = sh.cell(rownum,4).value
		city = sh.cell(rownum,5).value
		state = sh.cell(rownum,6).value
		
		zip = str(round(sh.cell(rownum,7).value)).rstrip('0').rstrip('.')
		zip = zip.strip()
		if len(zip) == 4: zip = '0' + zip

		rentown = 'own'		# Always own
		homeLength02 = random.choice(current_month_list)
		homeLength03 = random.choice(current_year_list)

		home_phone = str(round(sh.cell(rownum,8).value)).rstrip('0').rstrip('.')
		home_phone = home_phone.strip()
		homePhone01 = home_phone[:3]
		homePhone02 = home_phone[3:6]
		homePhone03 = home_phone[-4:]

		birth_date = list(xlrd.xldate_as_tuple(sh.cell(rownum,9).value, 0))
		
		birthDate01 = str(birth_date[1])
		if len(birthDate01) < 2: birthDate01 = '0' + birthDate01

		birthDate02 = str(birth_date[2])
		if len(birthDate02) < 2: birthDate02 = '0' + birthDate02

		birthDate03 = birth_date[0]

		alterPhone01 = homePhone01
		alterPhone02 = homePhone02
		alterPhone03 = homePhone03

		licensestate = str(sh.cell(rownum,11).value)
		license = str(sh.cell(rownum,10).value)

		gender = 'M'

		ssn = str(round(sh.cell(rownum,12).value)).rstrip('0').rstrip('.')
		if len(ssn) == 7: ssn = '00' + ssn
		if len(ssn) == 8: ssn = '0' + ssn
		ssnPart01 = ssn[:3]
		ssnPart02 = ssn[3:5]
		ssnPart03 = ssn[5:]

		bankName = sh.cell(rownum,13).value
		bankaccNumber = str(round(sh.cell(rownum,14).value)).rstrip('0').rstrip('.')
		
		bankabaRouting = str(round(sh.cell(rownum,15).value)).rstrip('0').rstrip('.')
		if len(bankabaRouting) == 7: bankabaRouting = '00' + bankabaRouting
		if len(bankabaRouting) == 8: bankabaRouting = '0' + bankabaRouting

		bankLength01 = '01'	# Hidden
		bankLength02 = '0' + str(int(homeLength02)+1)
		bankLength03 = str(int(homeLength03)+1)

		activechecking = 'yes'

		reference_name1 = sh.cell(rownum,16).value
		reference_name2 = sh.cell(rownum,17).value

		reference_relationship1 = 'friend'
		reference_relationship2 = 'friend'

		refPhoneOne01 = homePhone01
		refPhoneOne02 = random.randint(100, 999)
		refPhoneOne03 = random.randint(1000, 9999)

		refPhoneTwo01 = homePhone01
		refPhoneTwo02 = random.randint(100, 999)
		refPhoneTwo03 = random.randint(1000, 9999)

		currentlyemployed = 'yes'
		shifthours = '6'
		shift = 'day'
		companyname = sh.cell(rownum,18).value
		jobtitle = 'worker'
		supervisor_name = sh.cell(rownum,19).value

		work_phone = str(int(sh.cell(rownum,20).value))
		
		workPhone01 = work_phone[:3]
		workPhone02 = work_phone[3:6]
		workPhone03 = work_phone[-4:]

		payPeriod = sh.cell(rownum,21).value
		payPeriod = payPeriod.strip()
		nextPayDate = NEXT_PAY_DATE_BIWEEKLY

		if payPeriod == 'biweekly':
			payPeriod = 'Bi_Weekly'
		elif payPeriod == 'weekly':
			payPeriod = 'Weekly'
			nextPayDate = NEXT_PAY_DATE_WEEKLY
		elif payPeriod == 'monthly':
			payPeriod = 'Monthly'
			nextPayDate = NEXT_PAY_DATE_MONTHLY
		elif payPeriod == 'twicemonthly':
			payPeriod = 'Semi_Monthly'
			nextPayDate = NEXT_PAY_DATE_SEMI_MONTHLY


		takehomepay = str(round(sh.cell(rownum,22).value)).rstrip('0').rstrip('.')
		checkDeposit = 'Check'

		mainIncome = 'Job'

		nextPayday01 = NEXT_PAY_DATE_CURRENT.split('/')[0]
		nextPayday02 = NEXT_PAY_DATE_CURRENT.split('/')[1]
		nextPayday03 = NEXT_PAY_DATE_CURRENT.split('/')[2]

		secondPayday01 = nextPayDate.split('/')[0]
		secondPayday02 = nextPayDate.split('/')[1]
		secondPayday03 = nextPayDate.split('/')[2]

		dateHired01 = '0' + str(int(homeLength02)+1)
		dateHired02 = dateHired01
		dateHired03 = bankLength03

		clientIP = str(sh.cell(rownum,25).value)

		military = '0'
		citizen = 'yes'
		submit = 'true'

		query_args = {'loan_count' : loan_count, 'loan_amount': loan_amount, 'firstname': firstname, 'lastname': lastname, 'email': email,
						'address': address, 'city': city, 'state': state, 'zip': zip, 'rentown': rentown, 'homeLength02': homeLength02,
						'homeLength03': homeLength03, 'homePhone01': homePhone01, 'homePhone02': homePhone02, 'homePhone03': homePhone03,
						'birthDate01': birthDate01, 'birthDate02': birthDate02, 'birthDate03': birthDate03, 'alterPhone01': alterPhone01,
						'alterPhone02': alterPhone02, 'alterPhone03': alterPhone03, 'licensestate': licensestate, 'license': license,
						'gender': gender, 'ssnPart01': ssnPart01, 'ssnPart02': ssnPart02, 'ssnPart03': ssnPart03, 'bankName': bankName,
						'bankaccNumber': bankaccNumber, 'bankabaRouting': bankabaRouting, 'bankLength01': bankLength01, 'bankLength02': bankLength02,
						'bankLength03': bankLength03, 'activechecking': activechecking, 'reference_name1': reference_name1, 'reference_name2': reference_name2,
						'reference_relationship1': reference_relationship1, 'reference_relationship2': reference_relationship2, 'refPhoneOne01': refPhoneOne01,
						'refPhoneOne02': refPhoneOne02, 'refPhoneOne03': refPhoneOne03, 'refPhoneTwo01': refPhoneTwo01, 'refPhoneTwo02': refPhoneTwo02, 
						'refPhoneTwo03': refPhoneTwo03, 'currentlyemployed': currentlyemployed,
						'shifthours': shifthours, 'shift': shift, 'companyname': companyname, 'jobtitle': jobtitle, 'supervisor_name': supervisor_name,
						'workPhone01': workPhone01, 'workPhone02': workPhone02, 'workPhone03': workPhone03, 'payPeriod': payPeriod,
						'takehomepay': takehomepay, 'checkDeposit': checkDeposit, 'mainIncome': mainIncome, 'nextPayday01': nextPayday01,
						'nextPayday02': nextPayday02, 'nextPayday03': nextPayday03, 'secondPayday01': secondPayday01, 'secondPayday02': secondPayday02,
						'secondPayday03': secondPayday03, 'clientIP': clientIP, 'military': military, 'citizen': citizen, 'submit': submit, 
						'dateHired01': dateHired01, 'dateHired02': dateHired02, 'dateHired03': dateHired03}

		encoded_args = urllib.urlencode(query_args)

		request = urllib2.Request("", encoded_args)
		base64string = base64.encodestring('%s:%s' % ('', '')).replace('\n', '')
		request.add_header("Authorization", "Basic %s" % base64string)   
		request.add_header('User-Agent', 'User-Agent=Mozilla/5.0 (Windows NT 6.1; WOW64; rv:14.0) Gecko/20100101 Firefox/14.0.1')
		request.add_header('Accept','text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8')
		request.add_header('Accept-Charset', 'ISO-8859-1,utf-8;q=0.7,*;q=0.3')
		request.add_header('Accept-Encoding', 'gzip,deflate,sdch')
		request.add_header('Accept-Language', 'en-US,en;q=0.8')
		request.add_header('Cache-Control', 'max-age=0')
		request.add_header('Connection', 'keep-alive')
		request.add_header('Content-Type', 'application/x-www-form-urlencoded')
		request.add_header('Host', '')
		request.add_header('Referer', '')

		#############################################################################
		###############  Change cookies ##############################################
		request.add_header('Cookie', '__utma=170014704.1778539061.1346716205.1346716205.1346716205.1; __utmz=170014704.1346716205.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none)')

		try:
			result = urllib2.urlopen(request)
			resp_data = result.read()
			if resp_data.find('Service not available in your region') != -1:
				print resp_data
				continue
		
			f = open("dump/" + str(firstname) + "_output.txt", "w")
			print resp_data
			f.write(resp_data)
			f.close()
		except urllib2.URLError, e:
			continue

		#print f.write(result.read())
		#print request.data

		########	Disconnect VPN 	########
		disconnect_vpn()
		break


