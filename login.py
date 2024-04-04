from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
print("")
# Открываем сайт
driver.get("https://funpay.com/orders/trade")

# Ожидаем, пока пользователь не введет логин и пароль вручную
input("Пожалуйста, войдите в ваш аккаунт на сайте и нажмите Enter, когда закончите...")

# Удаляем предыдущий файл cookies.txt, если он существует
if os.path.exists("cookies.txt"):
    os.remove("cookies.txt")

# Сохраняем сессию, чтобы при следующем запуске авторизация была сохранена
cookies = driver.get_cookies()
with open("cookies.txt", "w") as file:
    file.write(str(cookies))

print("Файлы куки сохранены")

# Закрываем браузер
driver.quit()
