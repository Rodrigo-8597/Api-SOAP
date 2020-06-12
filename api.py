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


import openpyxl
import base64

def mayor_p(c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11,c12):
    lista=[c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11,c12]
    m=0
    p=0
    for i in range(0,12):
        if(ista[i]>m):
            m=lista[i]
            p=i
    return p

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
    elif(n==0):
        carreras.append(datos)
        return carreras
    else:
        carreras.append(datos)
    return ordenar(carreras)

class psuService(ServiceBase):
    @rpc(Unicode, Unicode, Unicode, _returns = Iterable(Unicode))
    def separacion(ctx, nombre_archivo, mime, dato_64):
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

class digitoService(ServiceBase):
    @rpc(Unicode, Unicode, _returns = Iterable(Unicode))
    def digito_verificador(ctx, rut, times):
        n_rut = rut.split('-')
        reversed_digits = map(int, reversed(str(n_rut[0])))
        factors = cycle(range(2, 8))
        s = sum(d * f for d, f in zip(reversed_digits, factors))
        mod = (-s) % 11
        if (mod == 10):
            mod = 'k'
        if (mod == 11):
            mod = 0
        if (str(mod) == n_rut[1]):
            yield ('Para el rut ' + str(rut) + ' ' + 'el digito verificador es '+ str(mod))
        else:
            yield('dv ingresado '+ str(n_rut[1]) + ' el dv correcto es '+ str(mod))


class nompropService(ServiceBase):
    @rpc(Unicode, Unicode, Unicode, Unicode,_returns = Iterable(Unicode))
    def generar_saludo(ctx, nom, pat, mat, sexo):
        nombreCompleto = nom + ' ' + pat + ' ' + mat + ' '
        nomComProp = nombreCompleto.title()
        if (int(sexo) == 1):
            sex = 'Sra. '
        else:
            sex = 'Sr. '
        yield ("Holi")
        yield (sex + ' ' + nomComProp )




application = Application(
    [
        digitoService,
        nompropService
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
