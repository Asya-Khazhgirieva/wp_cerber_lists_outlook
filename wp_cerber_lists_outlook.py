import re
import datetime
from win32com.client import Dispatch

print(datetime.datetime.strftime(datetime.datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Outlook>: Создание COM-объекта')
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

print(datetime.datetime.strftime(datetime.datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Входящие>: Поиск папки Отчеты WordPress')
inbox = outlook.Folders["email@domen.ru"].Folders["Входящие"].Folders["Отчеты WordPress"]

print(datetime.datetime.strftime(datetime.datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Отчеты WordPress>: Получение писем')
messages = inbox.Items
unread_messages = []

# Определяем ключевую фразу в теме писем
subject_keyword = "Уведомление WP Cerber: Число блокировок увеличилось"

# Поиск по теме писем
subject_filter = "@SQL=" + "urn:schemas:httpmail:subject LIKE '%" + subject_keyword + "%'"

# При необходимости добавляем доп. фильтры и тут их совмещаем
filter_criteria = subject_filter

# Выполните поиск писем в папке "Отчеты" с заданными фильтрами
filtered_items = messages.Restrict(filter_criteria)

# Фильтрация непрочитанных писем
for item in filtered_items:
    if item.UnRead:
        unread_messages.append(item)

#Список запрещенных имен пользователей
block_list = [
	"admin", "Administrator", "ds_admin", "admin123", 
	"Administrador", "poweruser", "admin9", "hostadmin", "websiteteam", "ademin", "admina", "admincms",
	"SixtyOne_Admin", "admin1", "administrator", "http", "webmaster",
	"editor", "root", "user", "wordpress", "superadmin", "admin2", "Admin",
	"administratoir", "itsme", "wadminw", "wpadmin", "wadmin", "admin1982", "adminlin",
	"admin_user", "test", "theadmin", "weboost_admin", "ad-ministrador",
	"admin0909", "adminbg", "administratoirr", "admin_vio", "betta_wp", "crisma-admin",
	"scwp-admin", "Senseadmin", "siteadmin", "suadmin", "systemwpadmin", "webmasterx",
	"wpadminns", "admin_rl", "1gridadmin", "cron_2fa48d5bef7cd1420d73ca40f3023354"
	]

# Словарь для хранения IP адресов и соответствующих им имен пользователей
ip_username_dict = {}

for item in unread_messages:
	# Извлекаем тело письма
	body = item.Body
	
	# Разбиваем тело письма на строки
	lines = [line.strip() for line in body.splitlines() if line.strip()]
	
	# Поиск строки с IP адресом и именем пользователя
	ip_address = None
	username = None
	username_match = None
	for line in lines:
		ip_match = re.search(r'для IP (\d+\.\d+\.\d+\.\d+)', line)
		if ip_match:
			ip_address = ip_match.group(1)
		if "Причина: Попытка входа с запрещенным именем" in line:
			username_match = re.search(r'Причина: Попытка входа с запрещенным именем: ([\w-]+)', line)
		if "Причина: Попытка войти с несуществующим именем" in line:
			username_match = re.search(r'Причина: Попытка войти с несуществующим именем пользователя: ([\w-]+)', line)

		if username_match:
			username = username_match.group(1)
			if username not in block_list:
				username = None
			else:
				item.UnRead = False  # Пометить как прочитанное

	if ip_address and username:
		ip_username_dict[ip_address] = username  # Добавить в словарь

# Выведите словарь с результатами
for ip, username in ip_username_dict.items():
	print(f"IP адрес: {ip}, Имя пользователя: {username}")

#Добавляем полученные IP адреса в чёрный список доступа через админку сайта

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

#указание параметров окна браузера
options = Options()
options.add_argument("--start-maximized") #полноэкранный режим окна
print(datetime.datetime.strftime(datetime.datetime.now(), "%d.%m.%Y %H:%M:%S") + ' <Chrome>: Открытие браузера')
chromedriver = Service("C:\\Users\user\chromedriver\chromedriver.exe")
driver = webdriver.Chrome(service=chromedriver, options=options)

#Заходим в админку
driver.get('URL')

login = driver.find_element(By.NAME, 'log')
password = driver.find_element(By.NAME, 'pwd')
sign = driver.find_element(By.XPATH, '//*[@id="wp-submit"]')
login.send_keys('login')
password.send_keys('password')
sign.click()

wp_cerber = driver.find_element(By.XPATH, '//*[@id="toplevel_page_cerber-security"]/a/div[3]')
wp_cerber.click()

wp_cerber_lists = driver.find_element(By.XPATH, '//*[@id="crb-admin"]/h2/a[6]')
wp_cerber_lists.click()


for ip, username in ip_username_dict.items():
	add_acl_comment = driver.find_elements(By.NAME, 'add_acl_comment')
	add_acl = driver.find_elements(By.NAME, 'add_acl')
	ip_add = add_acl[1]
	name = add_acl_comment[1]
	send = driver.find_element(By.XPATH, '//*[@id="crb-admin"]/div[2]/div[1]/div[2]/form/table/tbody/tr[1]/td[2]/input')

	ip_add.send_keys(ip)
	name.send_keys(username)
	send.click()

#Конец кода
driver.quit()
