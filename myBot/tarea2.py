import numpy as np
import matplotlib.pyplot as plt

# Defino la función de interpolación lineal 
def Trazado1(x, x0, x1, f_x0, f_x1):
    return f_x1 + (f_x1 - f_x0)/(x1 - x0) * (x - x1)


def Trazado2(x, x2, x1, f_x2, f_x1):
    return f_x2 + (f_x2 - f_x1)/(x2 - x1) * (x - x2)

def Trazado3(x, x3, x2, f_x3, f_x2):
    return f_x3 + (f_x3 - f_x2)/(x3 - x2) * (x - x3)

# Defino los valores dados
x0 = 1
x1 = 4
x2 = 5
x3 = 6
f_x0 = 0
f_x1 = 1.3862
f_x2 = 1.6094
f_x3 = 1.7917

# Calculo f(1.5) y f(5.7) usando el primer trazado
x_values_trazado1 = [1.5, 5.7]
y_values_trazado1 = [Trazado1(x, x0, x1, f_x0, f_x1) for x in x_values_trazado1]

# Calculo f(1.5) y f(5.7) usando el segundo trazado
x_values_trazado2 = [1.5, 5.7]
y_values_trazado2 = [Trazado2(x, x2, x1, f_x2, f_x1) for x in x_values_trazado2]

# Calculo f(1.5) y f(5.7) usando el tercer trazado
x_values_trazado3 = [1.5, 5.7]
y_values_trazado3 = [Trazado3(x, x3, x2, f_x3, f_x2) for x in x_values_trazado3]

# Imprimo los resultados
print("Trazado 1:")
print("funcion(1.5) =", y_values_trazado1[0])
print("funcion(5.7) =", y_values_trazado1[1])
print("\nTrazado 2:")
print("funcion(1.5) =", y_values_trazado2[0])
print("funcion(5.7) =", y_values_trazado2[1])
print("\nTrazado 3:")
print("funcion(1.5) =", y_values_trazado3[0])
print("funcion(5.7) =", y_values_trazado3[1])

# Creo un array de valores de x para la gráfica
x_range = np.linspace(0, 7, 100)

# Calculo los valores correspondientes de y para la gráfica usando el primer trazado
y_range_trazado1 = Trazado1(x_range, x0, x1, f_x0, f_x1)

y_range_trazado2 = Trazado2(x_range, x2, x1, f_x2, f_x1)

y_range_trazado3 = Trazado3(x_range, x3, x2, f_x3, f_x2)

# Grafico las funciones resultantes
plt.plot(x_range, y_range_trazado1, label='Trazado 1')
plt.plot(x_range, y_range_trazado2, label='Trazado 2')
plt.plot(x_range, y_range_trazado3, label='Trazado 3')

# Etiquetas de ejes y título
plt.xlabel('x')
plt.ylabel('f(x)')
plt.title('Gráfico de las funciones de interpolación lineal')
# Mostramos la leyenda
plt.legend()
# Mostramos la gráfica
plt.grid(True)
plt.show()


# se define una nueva función interpolante que utiliza los tres trazados
# Esta función Interpolante realiza la interpolación utilizando los tres trazados proporcionados (Trazado1, Trazado2, Trazado3) dependiendo del valor de x
def Interpolante(x, x0, x1, x2, x3, f_x0, f_x1, f_x2, f_x3):
    if x <= x1:
        return Trazado1(x, x0, x1, f_x0, f_x1)
    elif x <= x2:
        return Trazado2(x, x2, x1, f_x2, f_x1)
    else:
        return Trazado3(x, x3, x2, f_x3, f_x2)

# Calculamos f(1.5) y f(5.7) usando la nueva función interpolante
x_values = [1.5, 5.7]
y_values = [Interpolante(x, x0, x1, x2, x3, f_x0, f_x1, f_x2, f_x3) for x in x_values]

# Imprimimos los resultados
print("\ninterpolante:")
print("f(1.5) =", y_values[0])
print("f(5.7) =", y_values[1])

# Graficamos la nueva función interpolante junto con los puntos de intersección
y_range = [Interpolante(x, x0, x1, x2, x3, f_x0, f_x1, f_x2, f_x3) for x in x_range]

plt.plot(x_range, y_range, label='Interpolante')
plt.plot(x_values, y_values, 'ro')#Especifica que los puntos se trazarán como círculos rojos. La 'r' indica que los puntos serán rojos

# Etiquetamos los puntos de intersección
for i in range(len(x_values)):
    plt.text(x_values[i], y_values[i], f'({x_values[i]}, {y_values[i]:.4f})') #verticalalignment='bottom', horizontalalignment='right' Esto coloca el texto de modo que la parte inferior del texto esté alineada con las coordenadas (x_values[i], y_values[i]).

# ejes y título
plt.xlabel('x')
plt.ylabel('f(x)')
plt.title('Gráfico de la función interpolante')
# agregar una leyenda a la gráfica
plt.legend()
# Mostramos la gráfica
plt.grid(True)
plt.show()
