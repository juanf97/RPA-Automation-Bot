from selenium  import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import Excel

user = "margaritav"
password = "M1234#25"
company = "IGGA"
#Ruta de accceso al driver 
#"C:\Users\GMIT.DESKTOP-SLAHD3V\Ingenieria y Gestion Administrativa S.A.S\EQUIPO_GMIT - Documentos\Coordinadora de programación\RPA Officetrack\chromedriver_win32\chromedriver.exe"
Service = Service(executable_path="C:\\Users\\GMIT.DESKTOP-SLAHD3V\\Ingenieria y Gestion Administrativa S.A.S\\EQUIPO_GMIT - Documentos\\Coordinadora de programación\\RPA Officetrack\\chromedriver_win32\\chromedriver.exe")
#"C:\Users\JuanVC\OneDrive - Ingenieria y Gestion Administrativa S.A.S\Documentos\drivers\chromedriver.exe"
dirver = webdriver.Chrome(service=Service)
dirver.get("https://latam.officetrack.com/")
print("el link de esta pagina es: ",dirver.current_url)

input_element = dirver.find_element(By.ID,"txtUserName")
input_element.send_keys(user) #input_element.send_keys(Keys.ENTER)

input_Password = dirver.find_element(By.ID,"txtPassword")
input_Password.send_keys(password) #input_element.send_keys(Keys.ENTER)

input_Company = dirver.find_element(By.ID,"txtCompany")
input_Company.send_keys(company)
input_Company.send_keys(Keys.ENTER)

#abrir el link en una pestaña nueva
time.sleep(2)
dirver.execute_script("window.open('about:blank', '_blank')")
handles = dirver.window_handles
dirver.switch_to.window(handles[-1])
dirver.get("https://latam.officetrack.com/Tasks/Imports/ImportTasks.aspx")
time.sleep(2)
print("el link de esta pagina es 2: ",dirver.current_url)


time.sleep(3)
#cargar el archivo
#C:\Users\GMIT.DESKTOP-SLAHD3V\Ingenieria y Gestion Administrativa S.A.S\EQUIPO_GMIT - Documentos\Coordinadora de programación\RPA Officetrack\Excel_officeTrack\Cambios
input_file = dirver.find_element(By.ID,"RadUpload1file0").send_keys(r"C:\Users\GMIT.DESKTOP-SLAHD3V\Ingenieria y Gestion Administrativa S.A.S\EQUIPO_GMIT - Documentos\Coordinadora de programación\RPA Officetrack\Excel_officeTrack\Cambios\Importe.xlsx")
# time.sleep(3)
btn_import = dirver.find_element(By.XPATH,'//div/div/ul/li[1]/a/span/span/span/span[contains(@class,"rtbText")]')
btn_import.click()
print("Texto del elemento:",btn_import.text)

time.sleep(20)
dirver.quit()
print("se  han importado los datos correctamente" )
