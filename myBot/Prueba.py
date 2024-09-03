import pyautogui as robot
import time
import pyperclip as Rb




# Abrir Microsoft Word
robot.hotkey('win', 'r')  # Abrir el cuadro de diálogo Ejecutar
time.sleep(0.5)
robot.write('winword')  # Escribir 'winword' (el comando para abrir Word en Windows)
robot.press('enter')  # Presionar Enter para abrir Word
time.sleep(0.5)  # Esperar a que Word se abra (ajusta este tiempo según sea necesario)
# Esperar un momento para que Word se abra completamente
time.sleep(0.5)

time.sleep(0.5)
# Abrir un documento nuevo
robot.hotkey('ctrl', 'n')  # Atajo de teclado para abrir un documento nuevo en Word
for _ in range(10):
    Rb.copy('El desarrollo web del lado del servidor incluye las funciones complejas de backend que los sitios web llevan a cabo para mostrar información al usuario. Por ejemplo, los sitios web deben interactuar con las bases de datos, comunicarse con otros sitios web y proteger los datos cuando se los envía a través de la red.')
    time.sleep(.8)
    robot.hotkey("ctrl","v")
    robot.press('enter')
    

robot.press('enter') 
robot.write('Muchas Gracias le deseo un Feliz ')


# for _ in range(3):
#     time.sleep(0.5)
#     print("Posición actual del mouse:", robot.position())
#     robot.moveTo(870, 110, duration=0.5)
#     robot.click()
#     robot.write('Automatización GUI')
#     robot.press('enter')
#     time.sleep(1)

# entrar a word
# time.sleep(5)
# print("Posición actual del mouse:", robot.position())
# robot.moveTo(812, 741, duration=0.5)
# robot.click()
# # robot.moveTo(378, 262, duration=0.5)
# # robot.click()
# for _ in range(9):
#     Rb.copy('El desarrollo web del lado del servidor incluye las funciones complejas de backend que los sitios web llevan a cabo para mostrar información al usuario. Por ejemplo, los sitios web deben interactuar con las bases de datos, comunicarse con otros sitios web y proteger los datos cuando se los envía a través de la red.')
#     robot.hotkey("ctrl","v")
#     robot.press('enter')
    
# for _ in range(1):
#     robot.press('enter') 
# robot.write('Muchas Gracias le deseo un Feliz ')