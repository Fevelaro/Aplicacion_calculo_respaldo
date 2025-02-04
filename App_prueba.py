#!/usr/bin/env python3
import numpy as np
#from scipy.interpolate import interp1d
#from scipy import interpolate
from openpyxl import Workbook 
from scipy.interpolate import InterpolatedUnivariateSpline
from PySide6.QtWidgets import QFormLayout, QWidget, QLineEdit, QPushButton, QSpinBox, QDoubleSpinBox, QLabel, QApplication
from PySide6.QtWidgets import QButtonGroup, QCheckBox, QGridLayout,QVBoxLayout

potenciaKw=np.array([1,2,3,4,5,6,7,8,9,10])
Pack_1_6=np.array([60,25,16,10,6])
Pack_1_10=np.array([60,25,16,10,6])

class my_widget(QWidget):
    def __init__(self):
        super().__init__()
        #Label
        self.nombre_label= QLabel()
        self.nombre_label.setText("Nombre y número de local: ")
        self.numero_local=QSpinBox()
        self.numero_local.setMaximum(999)

        self.potencia_label= QLabel()
        self.potencia_label.setText("Potencia de UPS: ")
        self.P1=QCheckBox("6 kW")
        self.P2=QCheckBox("10 kW")

        self.P1.stateChanged.connect(lambda: self.deseleccionar_otros(self.P1))
        self.P2.stateChanged.connect(lambda: self.deseleccionar_otros(self.P2))

        self.Corriente_consumida = QDoubleSpinBox()
        self.Corriente_consumida.setSuffix(" A")
        self.corriente_label= QLabel()
        self.corriente_label.setText("Corriente consumida: ")

        self.mensaje_label = QLabel()
        self.mensaje_label.setText("Mensaje: ")

        self.mensaje = QLabel("")

        self.prueba_2_minutos_label = QLabel()
        self.prueba_2_minutos_label.setText("Con carga inicial completa. \nDespués de 2 minutos debe tener: ")
        self.prueba_2_minutos = QLabel()

        self.boton_calcular=QPushButton("Calcular")

        self.tiempo_de_respaldo = QLabel()
        self.tiempo_de_respaldo.setText("El respaldo mínimo esperado es: ")
        
        self.respaldo = QLabel("")
        
        #LineEdit
        self.nombre_LE=QLineEdit()

        #Layout
        self.layout=QGridLayout()

        self.layout.addWidget(self.nombre_label,0,0)
        self.layout.addWidget(self.nombre_LE,0,1)
        self.layout.addWidget(self.numero_local,0,2)

        self.layout.addWidget(self.potencia_label,1,0)
        self.layout.addWidget(self.P1,1,1)
        self.layout.addWidget(self.P2,1,2)

        self.layout.addWidget(self.corriente_label,2,0)
        self.layout.addWidget(self.Corriente_consumida,2,1)

        self.layout.addWidget(self.boton_calcular,4,1)


        self.layout.addWidget(self.tiempo_de_respaldo,5,0)
        self.layout.addWidget(self.respaldo,5,1)
        self.layout.addWidget(self.mensaje_label,7,0)
        self.layout.addWidget(self.mensaje,7,1)

        self.layout.addWidget(self.prueba_2_minutos_label,6,0)
        self.layout.addWidget(self.prueba_2_minutos,6,1)

        self.setLayout(self.layout)
        self.boton_calcular.clicked.connect(self.fcn_cal)

    def deseleccionar_otros(self, checkbox_actual):
        # Deseleccionar todos los demás checkboxes
        for checkbox in [self.P1, self.P2]:
            if checkbox != checkbox_actual:
                checkbox.setChecked(False)


    def fcn_cal(self):
        self.respaldo.setText("")
        self.mensaje.setText("")
        self.prueba_2_minutos.setText("")
        carga=float()
        carga=float(220) * self.Corriente_consumida.value()
        if self.P1.isChecked():
            pack=Pack_1_6
            potencia=potenciaKw[0:len(Pack_1_6)]
        elif self.P2.isChecked():
            pack=Pack_1_10
            potencia=potenciaKw[0:len(Pack_1_10)]
        else:
            self.mensaje.setText("Debe seleccionar una UPS")

        try: 
            potencia=potencia*1000
            if carga > potencia[len(potencia)-1]:
                self.mensaje.setText("La UPS está sobre cargada")
                self.respaldo2 = 0
            else :
                self.mensaje.setText("")
                self.respaldo2 = 1
            f=InterpolatedUnivariateSpline(potencia,pack)

            valor=f(carga)
            valor=round(float(valor))
            valor_esperado=round(float(valor*0.8))

            prueba_dos_min=(2*100)/valor_esperado
#            print("A los dos minutos debiese bajar un "+str(prueba_dos_min)+" por ciento")

            if self.respaldo2 == 1:
                self.respaldo.setText(str(valor_esperado)+" minutos")
                self.prueba_2_minutos.setText(str(round(100-prueba_dos_min))+" por ciento")
            else :
                self.respaldo.setText("Sin respaldo adecuado")

#            self.respaldo.setText(str(valor_esperado)+" minutos")
        except:
            self.mensaje.setText("Debe seleccionar una potencia de UPS")

        return


app=QApplication()
Mainwidget=my_widget()
Mainwidget.show()
app.exec()