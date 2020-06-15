import logging
from itertools import cycle

logging.basicConfig(level=logging.DEBUG)
from spyne import Application, rpc, ServiceBase, Integer, Unicode
from spyne import Iterable
from spyne.protocol.http import HttpRpc
from spyne.protocol.json import JsonDocument
from spyne.server.wsgi import WsgiApplication
from spyne.protocol.soap import Soap11
from spyne.model.primitive import String

from openpyxl import Workbook
import base64

def ordenar(carrera):
    tope=len(carrera)
    izquierda=[]
    derecha=[]
    centro=[]
    if(tope>1):
        pivote=carrera[0][1]
        for i in carrera:
            if(i[1]>pivote):
                izquierda.append(i)
            elif(i[1]==pivote):
                centro.append(i)
            elif(i[1]<pivote):
                derecha.append(i)
        return ordenar(izquierda)+centro+ordenar(derecha)
    else:
        return carrera

def almacenar(carreras, datos, max_ing):
    sw=0
    n=len(carreras)
    if(n==max_ing):
        if(datos[1]<carreras[max_ing-1][1]):
            return carreras
        for posicion in range (0,max_ing):
            if(sw==0):
                if(carreras[posicion][1]<=c1):
                    aux=carreras[posicion]
                    carreras[posicion]=datos
                    sw=1
            else:
                if(posicion<21):
                    aux2=carreras[posicion]
                    carreras[posicion]=aux
                    aux=aux2
                elif(posicion==21):
                    carreras[posicion]=aux
        return carreras
    elif(n>=0 and n<(max_ing-1)):
        carreras.append(datos)
        return carreras
    elif(n==(max_ing-1)):
        carreras.append(datos)
        return ordenar(carreras)

def entregarCarrera(indice):
    if(indice==0):
        return "21089"
    elif(indice==1):
        return "21002"
    elif(indice==2):
        return "21012"
    elif(indice==3):
        return "21048"
    elif(indice==4):
        return "21015"
    elif(indice==5):
        return "21081"
    elif(indice==6):
        return "21082"
    elif(indice==7):
        return "21047"
    elif(indice==8):
        return "21074"
    elif(indice==9):
        return "21032"
    elif(indice==10):
        return "21087"
    elif(indice==11):
        return "21073"
    elif(indice==12):
        return "21039"
    elif(indice==13):
        return "21080"
    elif(indice==14):
        return "21083"
    elif(indice==15):
        return "21024"
    elif(indice==16):
        return "21023"
    elif(indice==17):
        return "21043"
    elif(indice==18):
        return "21046"
    elif(indice==19):
        return "21071"
    elif(indice==20):
        return "21041"
    elif(indice==21):
        return "21076"
    elif(indice==22):
        return "21049"
    elif(indice==23):
        return "21075"
    elif(indice==24):
        return "21096"
    elif(indice==25):
        return "21031"
    elif(indice==26):
        return "21030"
    elif(indice==27):
        return "21045"   


def insertar(carreras):
    excel = Workbook()
    for carrera in carreras:
        indice = carreras.index(carrera)
        hoja = excel.create_sheet(entregarCarrera(indice),indice)
        fila = 1
        hoja['A'+str(fila)]='INDICE'
        hoja['B'+str(fila)]='RUT'
        hoja['C'+str(fila)]='PUNTAJE'
        for dato in carrera:
            fila+=1
            hoja['A'+str(fila)] = (carrera.index(dato)+1)
            hoja['B'+str(fila)] = dato[0]
            hoja['C'+str(fila)] = dato[1]
    del excel['Sheet']
    nombre="Admision UTEM.xlsx"
    excel.save(nombre)


class psuService(ServiceBase):
    @rpc(Unicode, Unicode, Unicode, _returns = Iterable(Unicode))
    def separacion(ctx, nombre_archivo, mime, dato_64):
        listado=[]
        for i in range(0,22):
            listado.append([i,i])
        listado=ordenar(listado)
        base64_bytes = dato_64.encode('ascii')
        message_bytes = base64.b64decode(base64_bytes)
        message = message_bytes.decode('ascii')
        message=message.split("\n")
        for linea in message: 
            if(len(linea)!=0):
                linea=linea.split(";")
                rut=linea[0]
                nem=int(linea[1])
                ranking=int(linea[2])
                lenguaje=int(linea[3])
                matematicas=int(linea[4])
                ciencias=int(linea[5])
                historia=int(linea[6])

                c1=float(nem*0.15+ranking*0.2+lenguaje*0.3+matematicas*0.25)
                c2=float(nem*0.2+ranking*0.2+lenguaje*0.4+matematicas*0.1)
                c3=float(nem*0.2+ranking*0.2+lenguaje*0.3+matematicas*0.15)
                c4_7=float(nem*0.1+ranking*0.2+lenguaje*0.3+matematicas*0.3)
                c8=float(nem*0.15+ranking*0.25+lenguaje*0.2+matematicas*0.2) 
                c9_10=float(nem*0.2+ranking*0.2+lenguaje*0.15+matematicas*0.35)
                c11=float(nem*0.15+ranking*0.35+lenguaje*0.2+matematicas*0.2)
                c12_13=float(nem*0.15+ranking*0.25+lenguaje*0.2+matematicas*0.3)
                c14_15=float(nem*0.1+ranking*0.25+lenguaje*0.15+matematicas*0.3)
                c16_17=float(nem*0.1+ranking*0.4+lenguaje*0.3+matematicas*0.1)
                c18=float(nem*0.2+ranking*0.3+lenguaje*0.2+matematicas*0.1)
                c19_28=float(nem*0.1+ranking*0.25+lenguaje*0.2+matematicas*0.35)    
                if(historia>=ciencias):
                    c1=c1+float(historia*0.1)
                    c2=c2+float(historia*0.1)
                    c3=c3+float(historia*0.15)
                    c4_7=c4_7+float(historia*0.1)
                    c8=c8+float(historia*0.2)
                    c9_10=c9_10+float(historia*0.1)
                    c11=c11+float(historia*0.1)
                    c12_13=c12_13+float(historia*0.1)
                    c14_15=c14_15+float(historia*0.2)
                    c16_17=c16_17+float(historia*0.1)
                    c18=c18+float(historia*0.2)
                    c19_28=c19_28+float(historia*0.1)
                else:
                    c1=c1+float(ciencias*0.1)
                    c2=c2+float(ciencias*0.1)
                    c3=c3+float(ciencias*0.1)
                    c4_7=c4_7+float(ciencias*0.1)
                    c8=c8+float(ciencias*0.2)
                    c9_10=c9_10+float(ciencias*0.1)
                    c11=c11+float(ciencias*0.1)
                    c12_13=c12_13+float(ciencias*0.1)
                    c14_15=c14_15+float(ciencias*0.2)
                    c16_17=c16_17+float(ciencias*0.1)
                    c18=c18+float(ciencias*0.2)
                    c19_28=c19_28+float(ciencias*0.1)
                    
                todos[0]=almacenar(todos[0], [rut,c1], 2100)
                todos[1]=almacenar(todos[1], [rut,c2], 2100)
                todos[2]=almacenar(todos[2], [rut,c3], 2100)
                todos[3]=almacenar(todos[3], [rut,c4_7], 2100)
                todos[4]=almacenar(todos[4], [rut,c8], 2100)
                todos[5]=almacenar(todos[5], [rut,c9_10], 2100)
                todos[6]=almacenar(todos[6], [rut,c11], 2100)
                todos[7]=almacenar(todos[7], [rut,c12_13], 2100)
                todos[8]=almacenar(todos[8], [rut,c14_15], 2100)
                todos[9]=almacenar(todos[9], [rut,c16_17], 2100)
                todos[10]=almacenar(todos[10], [rut,c18], 2100)
                todos[11]=almacenar(todos[11], [rut,c19_28], 2100)
            else:
                pass
        for i in range(0,12):
            posicion=0
            posiciones_par=[0,0]
            posicion_tetra=[0,0,0,0]
            posicion_ing=[0,0,0,0,0,0,0,0,0,0]
            if(i==0):
                while(len(carreras[0])<22):
                    carreras[0]=almacenar(carreras[0],todos[i][posicion], 22)
                    matriculados.append(todos[i][posicion][0])
                    posicion=posicion+1
            elif(i==1):
                while(len(carreras[1])<22):
                    if(todos[i][posicion][0] in matriculados):
                        posicion=posicion+1
                    else:
                        carreras[1]=almacenar(carreras[1],todos[i][posicion], 22)
                        matriculados.append(todos[i][posicion][0])
                        posicion=posicion+1
            elif(i==2):
                while(len(carreras[2])<22):
                    if(todos[i][posicion][0] in matriculados):
                        posicion=posicion+1
                    else:
                        carreras[2]=almacenar(carreras[2],todos[i][posicion], 22)
                        matriculados.append(todos[i][posicion][0])
                        posicion=posicion+1
            elif(i==3):
                while((len(carreras[3])+len(carreras[4])+len(carreras[5])+len(carreras[6]))<88):
                    if(n[0]==0):
                        cant_actual=len(carreras[3])
                        while(cant_actual==len(carreras[3])):
                            if(todos[i][posicion_tetra[0]][0] in matriculados):
                                posicion_tetra[0]=posicion_tetra[0]+1
                            else:
                                carreras[3]=almacenar(carreras[3],todos[i][posicion_tetra[0]], 22)
                                matriculados.append(todos[i][posicion_tetra[0]][0])
                                posicion_tetra[0]=posicion_tetra[0]+1
                        n[0]=1
                    elif(n[0]==1):
                        cant_actual=len(carreras[4])
                        while(cant_actual==len(carreras[4])):
                            if(todos[i][posicion_tetra[1]][0] in matriculados):
                                posicion_tetra[1]=posicion_tetra[1]+1
                            else:
                                carreras[4]=almacenar(carreras[4],todos[i][posicion_tetra[1]], 22)
                                matriculados.append(todos[i][posicion_tetra[1]][0])
                                posicion_tetra[1]=posicion_tetra[1]+1
                        n[0]=2
                    elif(n[0]==2):
                        cant_actual=len(carreras[5])
                        while(cant_actual==len(carreras[5])):
                            if(todos[i][posicion_tetra[2]][0] in matriculados):
                                posicion_tetra[2]=posicion_tetra[2]+1
                            else:
                                carreras[5]=almacenar(carreras[5],todos[i][posicion_tetra[2]], 22)
                                matriculados.append(todos[i][posicion_tetra[2]][0])
                                posicion_tetra[2]=posicion_tetra[2]+1
                        n[0]=3
                    elif(n[0]==3):
                        cant_actual=len(carreras[6])
                        while(cant_actual==len(carreras[6])):
                            if(todos[i][posicion_tetra[3]][0] in matriculados):
                                posicion_tetra[3]=posicion_tetra[3]+1
                            else:
                                carreras[6]=almacenar(carreras[6],todos[i][posicion_tetra[3]], 22)
                                matriculados.append(todos[i][posicion_tetra[3]][0])
                                posicion_tetra[3]=posicion_tetra[3]+1
                        n[0]=0
            elif(i==4):
                while(len(carreras[7])<22):
                    if(todos[i][posicion][0] in matriculados):
                        posicion=posicion+1
                    else:
                        carreras[7]=almacenar(carreras[7],todos[i][posicion], 22)
                        matriculados.append(todos[i][posicion][0])
                        posicion=posicion+1
            elif(i==5):
                while((len(carreras[8])+len(carreras[9]))<44):
                    if(n[1]==0):
                        cant_actual=len(carreras[8])
                        while(cant_actual==len(carreras[8])):
                            if(todos[i][posiciones_par[0]][0] in matriculados):
                                posiciones_par[0]=posiciones_par[0]+1
                            else:
                                carreras[8]=almacenar(carreras[8],todos[i][posiciones_par[0]], 22)
                                matriculados.append(todos[i][posiciones_par[0]][0])
                                posiciones_par[0]=posiciones_par[0]+1
                        n[1]=1
                    elif(n[1]==1):
                        cant_actual=len(carreras[9])
                        while(cant_actual==len(carreras[9])):
                            if(todos[i][posiciones_par[1]][0] in matriculados):
                                posiciones_par[1]=posiciones_par[1]+1
                            else:
                                carreras[9]=almacenar(carreras[9],todos[i][posiciones_par[1]], 22)
                                matriculados.append(todos[i][posiciones_par[1]][0])
                                posiciones_par[1]=posiciones_par[1]+1
                        n[1]=0
            elif(i==6):
                while(len(carreras[10])<22):
                    if(todos[i][posicion][0] in matriculados):
                        posicion=posicion+1
                    else:
                        carreras[10]=almacenar(carreras[10],todos[i][posicion], 22)
                        matriculados.append(todos[i][posicion][0])
                        posicion=posicion+1
            elif(i==7):
                while((len(carreras[11])+len(carreras[12]))<44):
                    if(n[2]==0):
                        cant_actual=len(carreras[11])
                        while(cant_actual==len(carreras[11])):
                            if(todos[i][posiciones_par[0]][0] in matriculados):
                                posiciones_par[0]=posiciones_par[0]+1
                            else:
                                carreras[11]=almacenar(carreras[11],todos[i][posiciones_par[0]], 22)
                                matriculados.append(todos[i][posiciones_par[0]][0])
                                posiciones_par[0]=posiciones_par[0]+1
                        n[2]=1
                    elif(n[2]==1):
                        cant_actual=len(carreras[12])
                        while(cant_actual==len(carreras[12])):
                            if(todos[i][posiciones_par[1]][0] in matriculados):
                                posiciones_par[1]=posiciones_par[1]+1
                            else:
                                carreras[12]=almacenar(carreras[12],todos[i][posiciones_par[1]], 22)
                                matriculados.append(todos[i][posiciones_par[1]][0])
                                posiciones_par[1]=posiciones_par[1]+1
                        n[2]=0
            elif(i==8):
                while((len(carreras[13])+len(carreras[14]))<44):
                    if(n[3]==0):
                        cant_actual=len(carreras[13])
                        while(cant_actual==len(carreras[13])):
                            if(todos[i][posiciones_par[0]][0] in matriculados):
                                posiciones_par[0]=posiciones_par[0]+1
                            else:
                                carreras[13]=almacenar(carreras[13],todos[i][posiciones_par[0]], 22)
                                matriculados.append(todos[i][posiciones_par[0]][0])
                                posiciones_par[0]=posiciones_par[0]+1
                        n[3]=1
                    elif(n[3]==1):
                        cant_actual=len(carreras[14])
                        while(cant_actual==len(carreras[14])):
                            if(todos[i][posiciones_par[1]][0] in matriculados):
                                posiciones_par[1]=posiciones_par[1]+1
                            else:
                                carreras[14]=almacenar(carreras[14],todos[i][posiciones_par[1]], 22)
                                matriculados.append(todos[i][posiciones_par[1]][0])
                                posiciones_par[1]=posiciones_par[1]+1
                        n[3]=0
            elif(i==9):
                while((len(carreras[15])+len(carreras[16]))<44):
                    if(n[4]==0):
                        cant_actual=len(carreras[15])
                        while(cant_actual==len(carreras[15])):
                            if(todos[i][posiciones_par[0]][0] in matriculados):
                                posiciones_par[0]=posiciones_par[0]+1
                            else:
                                carreras[15]=almacenar(carreras[15],todos[i][posiciones_par[0]], 22)
                                matriculados.append(todos[i][posiciones_par[0]][0])
                                posiciones_par[0]=posiciones_par[0]+1
                        n[4]=1
                    elif(n[4]==1):
                        cant_actual=len(carreras[16])
                        while(cant_actual==len(carreras[16])):
                            if(todos[i][posiciones_par[1]][0] in matriculados):
                                posiciones_par[1]=posiciones_par[1]+1
                            else:
                                carreras[16]=almacenar(carreras[16],todos[i][posiciones_par[1]], 22)
                                matriculados.append(todos[i][posiciones_par[1]][0])
                                posiciones_par[1]=posiciones_par[1]+1
                        n[4]=0
            elif(i==10):
                while(len(carreras[17])<22):
                    if(todos[i][posicion][0] in matriculados):
                        posicion=posicion+1
                    else:
                        carreras[17]=almacenar(carreras[17],todos[i][posicion], 22)
                        matriculados.append(todos[i][posicion][0])
                        posicion=posicion+1
            elif(i==11):
                while((len(carreras[18])+len(carreras[19])+len(carreras[20])+len(carreras[21])+len(carreras[22])+len(carreras[23])+len(carreras[24])+len(carreras[25])+len(carreras[26])+len(carreras[27]))<220):
                    if(n[5]==0):
                        cant_actual=len(carreras[18])
                        while(cant_actual==len(carreras[18])):
                            if(todos[i][posicion_ing[0]][0] in matriculados):
                                posicion_ing[0]=posicion_ing[0]+1
                            else:
                                carreras[18]=almacenar(carreras[18],todos[i][posicion_ing[0]], 22)
                                matriculados.append(todos[i][posicion_ing[0]][0])
                                posicion_ing[0]=posicion_ing[0]+1
                        n[5]=1
                    elif(n[5]==1):
                        cant_actual=len(carreras[19])
                        while(cant_actual==len(carreras[19])):
                            if(todos[i][posicion_ing[1]][0] in matriculados):
                                posicion_ing[1]=posicion_ing[1]+1
                            else:
                                carreras[19]=almacenar(carreras[19],todos[i][posicion_ing[1]], 22)
                                matriculados.append(todos[i][posicion_ing[1]][0])
                                posicion_ing[1]=posicion_ing[1]+1
                        n[5]=2
                    elif(n[5]==2):
                        cant_actual=len(carreras[20])
                        while(cant_actual==len(carreras[20])):
                            if(todos[i][posicion_ing[2]][0] in matriculados):
                                posicion_ing[2]=posicion_ing[2]+1
                            else:
                                carreras[20]=almacenar(carreras[20],todos[i][posicion_ing[2]], 22)
                                matriculados.append(todos[i][posicion_ing[2]][0])
                                posicion_ing[2]=posicion_ing[2]+1
                        n[5]=3
                    elif(n[5]==3):
                        cant_actual=len(carreras[21])
                        while(cant_actual==len(carreras[21])):
                            if(todos[i][posicion_ing[3]][0] in matriculados):
                                posicion_ing[3]=posicion_ing[3]+1
                            else:
                                carreras[21]=almacenar(carreras[21],todos[i][posicion_ing[3]], 22)
                                matriculados.append(todos[i][posicion_ing[3]][0])
                                posicion_ing[3]=posicion_ing[3]+1
                        n[5]=4
                    elif(n[5]==4):
                        cant_actual=len(carreras[22])
                        while(cant_actual==len(carreras[22])):
                            if(todos[i][posicion_ing[2]][0] in matriculados):
                                posicion_ing[2]=posicion_ing[2]+1
                            else:
                                carreras[22]=almacenar(carreras[22],todos[i][posicion_ing[2]], 22)
                                matriculados.append(todos[i][posicion_ing[2]][0])
                                posicion_ing[2]=posicion_ing[2]+1
                        n[5]=5
                    elif(n[5]==5):
                        cant_actual=len(carreras[23])
                        while(cant_actual==len(carreras[23])):
                            if(todos[i][posicion_ing[2]][0] in matriculados):
                                posicion_ing[2]=posicion_ing[2]+1
                            else:
                                carreras[23]=almacenar(carreras[23],todos[i][posicion_ing[2]], 22)
                                matriculados.append(todos[i][posicion_ing[2]][0])
                                posicion_ing[2]=posicion_ing[2]+1
                        n[5]=6
                    elif(n[5]==6):
                        cant_actual=len(carreras[24])
                        while(cant_actual==len(carreras[24])):
                            if(todos[i][posicion_ing[2]][0] in matriculados):
                                posicion_ing[2]=posicion_ing[2]+1
                            else:
                                carreras[24]=almacenar(carreras[24],todos[i][posicion_ing[2]], 22)
                                matriculados.append(todos[i][posicion_ing[2]][0])
                                posicion_ing[2]=posicion_ing[2]+1
                        n[5]=7
                    elif(n[5]==7):
                        cant_actual=len(carreras[25])
                        while(cant_actual==len(carreras[25])):
                            if(todos[i][posicion_ing[2]][0] in matriculados):
                                posicion_ing[2]=posicion_ing[2]+1
                            else:
                                carreras[25]=almacenar(carreras[25],todos[i][posicion_ing[2]], 22)
                                matriculados.append(todos[i][posicion_ing[2]][0])
                                posicion_ing[2]=posicion_ing[2]+1
                        n[5]=8
                    elif(n[5]==8):
                        cant_actual=len(carreras[26])
                        while(cant_actual==len(carreras[26])):
                            if(todos[i][posicion_ing[2]][0] in matriculados):
                                posicion_ing[2]=posicion_ing[2]+1
                            else:
                                carreras[26]=almacenar(carreras[26],todos[i][posicion_ing[2]], 22)
                                matriculados.append(todos[i][posicion_ing[2]][0])
                                posicion_ing[2]=posicion_ing[2]+1
                        n[5]=9
                    elif(n[5]==9):
                        cant_actual=len(carreras[27])
                        while(cant_actual==len(carreras[27])):
                            if(todos[i][posicion_ing[2]][0] in matriculados):
                                posicion_ing[2]=posicion_ing[2]+1
                            else:
                                carreras[27]=almacenar(carreras[27],todos[i][posicion_ing[2]], 22)
                                matriculados.append(todos[i][posicion_ing[2]][0])
                                posicion_ing[2]=posicion_ing[2]+1
                        n[5]=0
        insertar(carreras)
        archivo=open("Admision UTEM.xlsx", "r")
        todo=archivo.read()
        exc=todo.encode('ascii')
        base64_bytes=base64.b64encode(exc)
        exc_64=base64_bytes.decode('ascii')
        yield("Admision UTEM.xlsx")
        yield(mime_exc)
        yield(exc_64)
        """
        mime=mime.upper()
        if(mime!="CSV" and mime!="TEXT" and mime!="TXT"):
            return ("Tipo MIME especificado invalido; favor enviar especificacion como CSV, TEXT, TXT")

        n=[0,0,0,0,0,0]
        carreras=[]
        for i in range(0,28):
            carreras.append([])
        base64_bytes=dato_64.encode('ascii')
        message_bytes=base64.b64decode(base64_bytes)
        message=message_bytes.decode('ascii')

        message=message.split('\n')
        for i in range(0,len(message)):
            sw=0
            message[i]=message[i].split(';')
            rut=message[i][0]
            nem=int(message[i][1])
            ranking=int(message[i][2])
            lenguaje=int(message[i][3])
            matematicas=int(message[i][4])
            ciencias=int(message[i][5])
            historia=int(message[i][6])

            c1=float(nem*0.15+ranking*0.2+lenguaje*0.3+matematicas*0.25)
            c2=float(nem*0.2+ranking*0.2+lenguaje*0.4+matematicas*0.1)
            c3=float(nem*0.2+ranking*0.2+lenguaje*0.3+matematicas*0.15)
            c4_7=float(nem*0.1+ranking*0.2+lenguaje*0.3+matematicas*0.3)
            c8=float(nem*0.15+ranking*0.25+lenguaje*0.2+matematicas*0.2) 
            c9_10=float(nem*0.2+ranking*0.2+lenguaje*0.15+matematicas*0.35)
            c11=float(nem*0.15+ranking*0.35+lenguaje*0.2+matematicas*0.2)
            c12_13=float(nem*0.15+ranking*0.25+lenguaje*0.2+matematicas*0.3)
            c14_15=float(nem*0.1+ranking*0.25+lenguaje*0.15+matematicas*0.3)
            c16_17=float(nem*0.1+ranking*0.4+lenguaje*0.3+matematicas*0.1)
            c18=float(nem*0.2+ranking*0.3+lenguaje*0.2+matematicas*0.1)
            c19_28=float(nem*0.1+ranking*0.25+lenguaje*0.2+matematicas*0.35)    
            if(historia>=ciencias):
                c1=c1+float(historia*0.1)
                c2=c2+float(historia*0.1)
                c3=c3+float(historia*0.15)
                c4_7=c4_7+float(historia*0.1)
                c8=c8+float(historia*0.2)
                c9_10=c9_10+float(historia*0.1)
                c11=c11+float(historia*0.1)
                c12_13=c12_12+float(historia*0.1)
                c14_15=c14_15+float(historia*0.2)
                c16_17=c16_17+float(historia*0.1)
                c18=c18+float(historia*0.2)
                c19_28=c19_28+float(historia*0.1)
            else:
                c1=c1+float(ciencias*0.1)
                c2=c2+float(ciencias*0.1)
                c3=c3+float(ciencias*0.1)
                c4_7=c4_7+float(ciencias*0.1)
                c8=c8+float(ciencias*0.2)
                c9_10=c9_10+float(ciencias*0.1)
                c11=c11+float(ciencias*0.1)
                c12_13=c12_12+float(ciencias*0.1)
                c14_15=c14_15+float(ciencias*0.2)
                c16_17=c16_17+float(ciencias*0.1)
                c18=c18+float(ciencias*0.2)
                c19_28=c19_28+float(ciencias*0.1)
            mayor=mayor_p(c1,c2,c3,c4_7,c8,c9_10,c11,c12_13,c14_15,c16_17,c18,c19_28)
            if(mayor==0):
                datos=[rut,c1]
                carreras[0]=almacenar(carreras[0], datos, 22)
            elif(mayor==1):
                datos=[rut,c2]
                carreras[1]=almacenar(carreras[1], datos, 22)
            elif(mayor==2):
                datos=[rut,c3]
                carreras[2]=almacenar(carreras[2], datos, 22)
            elif(mayor==3):
                datos=[rut,c4_7]
                if(n[0]==0):
                    carreras[3]=almacenar(carreras[3], datos, 22)
                    n[0]=1
                elif(n[0]==1):
                    carreras[4]=almacenar(carreras[4], datos, 22)
                    n[0]=2
                elif(n[0]==2):
                    carreras[5]=almacenar(carreras[5], datos, 22)
                    n[0]=3
                elif(n[0]==3):
                    carreras[6]=almacenar(carreras[6], datos, 22)
                    n[0]=0
            elif(mayor==4):
                datos=[rut,c8]
                carreras[7]=almacenar(carreras[7], datos, 22)
            elif(mayor==5):
                datos=[rut,c9_10]
                if(n[1]==0):
                    carreras[8]=almacenar(carreras[8], datos, 22)
                    n[1]=1
                elif(n[1]==1):
                    carreras[9]=almacenar(carreras[9], datos, 22)
                    n[1]=0
            elif(mayor==6):
                datos=[rut,c11]
                carreras[10]=almacenar(carreras[10], datos, 22)
            elif(mayor==7):
                datos=[rut,c12_13]
                if(n[2]==0):
                    carreras[11]=almacenar(carreras[11], datos, 22)
                    n[2]=1
                elif(n[2]==1):
                    carreras[12]=almacenar(carreras[12], datos, 22)
                    n[2]=0
            elif(mayor==8):
                datos=[rut,c4_15]
                if(n[3]==0):
                    carreras[13]=almacenar(carreras[13], datos, 22)
                    n[3]=1
                elif(n[3]==1):
                    carreras[14]=almacenar(carreras[14], datos, 22)
                    n[3]=0
            elif(mayor==9):
                datos=[rut,c16_17]
                if(n[4]==0):
                    carreras[15]=almacenar(carreras[15], datos, 22)
                    n[4]=1
                elif(n[4]==1):
                    carreras[16]=almacenar(carreras[16], datos, 22)
                    n[4]=0
            elif(mayor==11):
                datos=[rut,c18]
                carreras[17]=almacenar(carreras[17], datos, 22)
            elif(mayor==12):
                datos=[rut,c19_28]
                if(n[5]==0):
                    carreras[18]=almacenar(carreras[18], datos, 22)
                    n[5]=1
                elif(n[5]==1):
                    carreras[19]=almacenar(carreras[19], datos, 22)
                    n[5]=2
                elif(n[5]==2):
                    carreras[20]=almacenar(carreras[20], datos, 22)
                    n[5]=3
                elif(n[5]==3):
                    carreras[21]=almacenar(carreras[21], datos, 22)
                    n[5]=4
                elif(n[5]==4):
                    carreras[22]=almacenar(carreras[22], datos, 22)
                    n[5]=5
                elif(n[5]==5):
                    carreras[23]=almacenar(carreras[23], datos, 22)
                    n[5]=6
                elif(n[5]==6):
                    carreras[24]=almacenar(carreras[24], datos, 22)
                    n[5]=7
                elif(n[5]==7):
                    carreras[25]=almacenar(carreras[25], datos, 22)
                    n[5]=8
                elif(n[5]==8):
                    carreras[26]=almacenar(carreras[26], datos, 22)
                    n[5]=9
                elif(n[5]==9):
                    carreras[27]=almacenar(carreras[27], datos, 22)
                    n[5]=0

        wb=Workbook()
        ws=wb.active
        ws.title="C. 1"
        ws1 = wb.create_sheet("C. 2")
        ws2 = wb.create_sheet("C. 3")
        ws3 = wb.create_sheet("C. 4")
        ws4 = wb.create_sheet("C. 5")
        ws5 = wb.create_sheet("C. 6")
        ws6 = wb.create_sheet("C. 7")
        ws7 = wb.create_sheet("C. 8")
        ws8 = wb.create_sheet("C. 9")
        ws9 = wb.create_sheet("C. 10")
        ws10 = wb.create_sheet("C. 11")
        ws11 = wb.create_sheet("C. 12")
        ws12 = wb.create_sheet("C. 13")
        ws13 = wb.create_sheet("C. 14")
        ws14 = wb.create_sheet("C. 15")
        ws15 = wb.create_sheet("C. 16")
        ws16 = wb.create_sheet("C. 17")
        ws17 = wb.create_sheet("C. 18")
        ws18 = wb.create_sheet("C. 19")
        ws19 = wb.create_sheet("C. 20")
        ws20 = wb.create_sheet("C. 21")
        ws21 = wb.create_sheet("C. 22")
        ws22 = wb.create_sheet("C. 23")
        ws23 = wb.create_sheet("C. 24")
        ws24 = wb.create_sheet("C. 25")
        ws25 = wb.create_sheet("C. 26")
        ws26 = wb.create_sheet("C. 27")
        ws27 = wb.create_sheet("C. 28")

        for i in range(0,28):
            for j in range(0, len(carreras[i])):
                if(i==0):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==1):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==2):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==3):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==4):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==5):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==6):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==7):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==8):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==9):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==10):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==11):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==12):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==13):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==14):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==15):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==16):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==17):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==18):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==19):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==20):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==21):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==22):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==23):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==24):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==25):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==26):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==27):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                elif(i==28):
                    ws['A'+str(j+1)]=carrera[i][j][0]
                    ws['B'+str(j+1)]=carrera[i][j][1]
                
        wb.save(filename = 'admitidos.xlsx')
        yield("admitidos")
        yield("spreadsheet")
        yield(excel_base64)
        """

application = Application(
    [
        psuService
    ],
    tns = 'spyne.examples.hello.soap',
    in_protocol = Soap11(),
    out_protocol = Soap11()
)

if __name__ == '__main__':
    # You can use any Wsgi server. Here, we chose
    # Python's built-in wsgi server but you're not
    # supposed to use it in production.
    from wsgiref.simple_server import make_server
    wsgi_app = WsgiApplication(application)
    server = make_server('127.0.0.1', 8000, wsgi_app)
    print("\n\tServidor en Linea")
    server.serve_forever()
