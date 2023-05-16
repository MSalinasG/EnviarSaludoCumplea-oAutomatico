import pywhatkit
import pandas as pd
import time
import datetime
import tkinter as tk
from tkinter import messagebox

data = pd.read_excel("D:\Envio_msj_cumpleaños_wsp\datos.xlsx")  

def obtener_mensaje_x_tipo(tipo):
    if tipo == 1:
        return "¡Feliz cumpleaños {0}! \U0001F973\U0001F973 Espero que este día esté lleno de alegría, amor y bendiciones. Eres una persona especial y te mereces los mejor hoy. \U0001F381 Que este nuevo año de vida esté lleno de éxitos, salud y felicidad. \U0001F917 ¡Disfruta al máximo tu día y que todos tus sueños se hagan realidad! \U0001F382 ¡Feliz cumpleaños! \U0001F389\U0001F389"
    elif tipo == 2:
        return "¡Feliz cumpleaños, {0}! \U0001F973 Que este día especial te llene de alegría y bendiciones. \U0001F389\U0001F389 Espero que encuentres felicidad y éxito en todos tus proyectos. \U0001F382 ¡Que tengas un año maravilloso! \U0001F973\U0001F973"


def enviar_mensaje(destinatario, mensaje, hora, minutos):     
    pywhatkit.sendwhatmsg(destinatario, mensaje, hora, minutos, 10, True, 2)

for i in range(len(data)):
    try :
        #Enviando mensaje
        celular = "+" + data.loc[i,"Celular"].astype(str)
        nacimiento = data.loc[i,"FechaNacimiento"]
        tipo = data.loc[i,"Tipo"]        
        apodo = data.at[i,"Apodo"]

        # Obtén la fecha actual
        fecha_actual = datetime.datetime.now().strftime("%d/%m")
       
        if fecha_actual == nacimiento:
            if pd.isna(apodo) or apodo.strip() == "":
                nombre_final = data.loc[i,"Nombre"]
            else:
                nombre_final = data.loc[i,"Apodo"]

            mensaje = obtener_mensaje_x_tipo(tipo).format(nombre_final)

            minutos_a_sumar = 1  # Número de minutos a sumar
            hora_actual = datetime.datetime.now().time()
            nueva_hora = datetime.datetime.combine(datetime.date.today(), hora_actual) + datetime.timedelta(minutes=minutos_a_sumar)
            nueva_hora = nueva_hora.time()
            nueva_hora_en_horas = nueva_hora.hour
            nueva_hora_en_minutos = nueva_hora.minute          
        
            enviar_mensaje(celular, mensaje, nueva_hora_en_horas, nueva_hora_en_minutos)

    except:

        print("No se pudo encontrar")

# Crea una ventana oculta para el mensaje emergente
root = tk.Tk()
root.withdraw()

# Muestra un mensaje emergente
messagebox.showinfo("Mensaje", "Se enviaron todos los mensajes de hoy")

# Cierra la ventana oculta
root.destroy()