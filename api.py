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
        if(lista[i]>m):
            m=lista[i]
            p=i
    return p
    

def funcion():
    pass

class psuService(ServiceBase):
    @rpc(Unicode, Unicode, Unicode, _returns = Iterable(Unicode))
    def separacion(ctx, nombre_archivo, mime, dato_64):
        mime=mime.upper()
        if(mime!="CSV" and mime!="TEXT" and mime!="TXT"):
            return ("Tipo MIME especificado invalido; favor enviar especificacion como CSV, TEXT, TXT")
        carreras=[]
        for i in range(0,28):
            carreras.append([])

        base64_bytes=dato_64.encode('ascii')
        message_bytes=base64.b64decode(base64_bytes)
        message=message_bytes.decode('ascii')


        message=message.split('\n')
        for i in range(0,len(message)):
            message[i]=message[i].split(';')
            nem=int(message[i][1])
            ranking=int(message[i][2])
            lenguaje=int(message[i][3])
            matematicas=int(message[i][4])
            ciencias=int(message[i][5])
            historia=int(message[i][6])

            c1=float(nem*0.1+ranking*0.1+lenguaje*0.1+matematicas*0.1)
            c2=float(nem*0.1+ranking*0.1+lenguaje*0.1+matematicas*0.1)
            c3=float(nem*0.1+ranking*0.1+lenguaje*0.1+matematicas*0.1)
            c4_7=float(nem*0.1+ranking*0.1+lenguaje*0.1+matematicas*0.1)
            c8=float(nem*0.1+ranking*0.1+lenguaje*0.1+matematicas*0.1)
            c9_10=float(nem*0.1+ranking*0.1+lenguaje*0.1+matematicas*0.1)
            c11=float(nem*0.1+ranking*0.1+lenguaje*0.1+matematicas*0.1)
            c12_13=float(nem*0.1+ranking*0.1+lenguaje*0.1+matematicas*0.1)
            c14_15=float(nem*0.1+ranking*0.1+lenguaje*0.1+matematicas*0.1)
            c16_17=float(nem*0.1+ranking*0.1+lenguaje*0.1+matematicas*0.1)
            c18=float(nem*0.1+ranking*0.1+lenguaje*0.1+matematicas*0.1)
            c19_28=float(nem*0.1+ranking*0.1+lenguaje*0.1+matematicas*0.1)           
            
            if(historia>=ciencias):
                c1=c1+float(historia*0.1)
                c2=c2+float(historia*0.1)
                c3=c3+float(historia*0.1)
                c4_7=c4_7+float(historia*0.1)
                c8=c8+float(historia*0.1)
                c9_10=c9_10+float(historia*0.1)
                c11=c11+float(historia*0.1)
                c12_13=c12_12+float(historia*0.1)
                c14_15=c14_15+float(historia*0.1)
                c16_17=c16_17+float(historia*0.1)
                c18=c18+float(historia*0.1)
                c19_28=c19_28+float(historia*0.1)
            else:
                c1=c1+float(ciencias*0.1)
                c2=c2+float(ciencias*0.1)
                c3=c3+float(ciencias*0.1)
                c4_7=c4_7+float(ciencias*0.1)
                c8=c8+float(ciencias*0.1)
                c9_10=c9_10+float(ciencias*0.1)
                c11=c11+float(ciencias*0.1)
                c12_13=c12_12+float(ciencias*0.1)
                c14_15=c14_15+float(ciencias*0.1)
                c16_17=c16_17+float(ciencias*0.1)
                c18=c18+float(ciencias*0.1)
                c19_28=c19_28+float(ciencias*0.1)
            mayor=mayor_p(c1,c2,c3,c4_7,c8,c9_10,c11,c12_13,c14_15,c16_17,c18,c19_28)
            if(mayor==0):
                pass
            elif(mayor==1):
                pass
            elif(mayor==2):
                pass
            elif(mayor==3):
                pass
            elif(mayor==4):
                pass
            elif(mayor==5):
                pass
            elif(mayor==6):
                pass
            elif(mayor==7):
                pass
            elif(mayor==8):
                pass
            elif(mayor==9):
                pass
            elif(mayor==11):
                pass
            elif(mayor==12):
                pass
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
