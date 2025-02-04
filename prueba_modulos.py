from aplicacion_modular.calculador_datos import fcn_cal

try:
    mensaje=fcn_cal(6,2,6)

except TypeError:
    mensaje="Faltan datos, corrobore: \n¿Ingresó corriente?\n¿Ingresó potencia de UPS? \n¿Ingresó número UPS?"
print(mensaje)