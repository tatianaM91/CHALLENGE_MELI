import pandas as pd
import xlrd                                              #Importar datos en formato excel
import smtplib                                           #Conexión con el servidor de correo
from email.mime.multipart import MIMEMultipart           #Creación del cuerpo del correo electrónico 1
from email.mime.application import MIMEApplication       #Creación del cuerpo del correo electrónico 2            
from email.mime.text import MIMEText                     #Creación del cuerpo del correo electrónico 3

def importar_excel():
    global Bd
    #Importar excel
    wb = xlrd.open_workbook('user_manager.xlsx')   
    hoja = wb.sheet_by_index(0) 
    Fila=[]
    Bd=[]
    #Crear una lista de listas
    for i in range(hoja.nrows):
        for j in range (hoja.ncols):
            Fila.append(hoja.cell_value(i, j))
        Bd.append(Fila)
        Fila=[]
        
    print (Bd)

def importar_json():
    global Bd
    global Bd2    
    global Lista_aux
    #Leer json y crear una variable de tipo diccionario
    json_f = pd.read_json ('dblist.json')
    db_lista=dict(json_f['db_list'])
    Lista_aux=[]
    Fila=[]
    Bd2=[]
    F=0
    #Cruce de la base de datos (json y excel)
    for i in db_lista:    
        Fila.append(db_lista[i]['dn_name'])
        try:
            Fila.append(db_lista[i]['owner']['email'])
            Lista_aux.append(db_lista[i]['owner']['uid'])
            for j in Bd:
                if(j[1]==db_lista[i]['owner']['uid']):
                    F=j[3]
                    break
                else:
                    F='Null'
            Fila.append(F)
            F=0
        except:
            Fila.append('Null')
            Fila.append('Null')
        Fila.append(db_lista[i]['classification'])
        Bd2.append(Fila)
        Fila=[]
        
    print (Bd2)

def Conectar_servidor(dn, e1, e2, cla, pos):
    global Lista_aux
    # Crear el objeto mensaje
    mensaje = MIMEMultipart()             
    mensaje['de']     = 'agropru1@gmail.com'       #Correo de prueba para enviar algo desde la página
    mensaje['para']   = 'jtmartinm@gmail.com'      #Correo funcionario a cargo            
    #Cuerpo del mensaje
    msn = ('Este mensaje fue enviado por: Grupo de seguridad Mercado Libre'+'\n'+
           '(Nombre BD = '+dn+', Correo owner = '+e1+', Correo manager = '+e2+')'+'\n'+
           'Solicitamos de su colaboracion como owner de la cuenta con ID ' + Lista_aux[pos]+
           ', ya que tiene clasificacion de '+ str(cla) +', por lo que solicitamos de su aprobación.')
    mensaje.attach(MIMEText(msn, 'plain'))
    # Datos de acceso a la cuenta de usuario
    usuario   ='agropru1'
    contrasena='Agrosavia123'          
    #Interfaz de conexión con el servidor de gmail
    servidor = smtplib.SMTP('smtp.gmail.com:587')
    servidor.starttls()
    servidor.login(usuario, contrasena)
    servidor.sendmail(mensaje['de'], mensaje['para'], mensaje.as_string())
    servidor.quit()  
            
def enviar_correo():
    global Bd2
    Enviar_correo=[]
    conta=0
    for i in Bd2:        
        if(i[3]['confidentiality']=='high' or i[3]['integrity']=='high' or i[3]['availability']=='high'):
            Enviar_correo.append("SI")
            Conectar_servidor(i[0], i[1], i[2], i[3], conta)
        else:
            Enviar_correo.append("NO")
        conta=conta+1
    df = pd.DataFrame(Bd2, columns=['Nombre base de datos', 'Email owner', 'Email manager', 'Clasificacion'])
    df.to_excel('Basedatos.xlsx')

if __name__=='__main__':
    importar_excel()
    importar_json()
    enviar_correo()