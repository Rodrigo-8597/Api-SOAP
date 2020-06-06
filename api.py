import logging
from itertools import cycle

logging.basicConfig(level=logging.DEBUG)
from spyne import Application, rpc, ServiceBase, \
    Integer, Unicode
from spyne import Iterable
from spyne.protocol.http import HttpRpc
from spyne.protocol.json import JsonDocument
from spyne.server.wsgi import WsgiApplication
from spyne.protocol.soap import Soap11
from spyne.model.primitive import String

import openpyxl
import Base64

class psuService(ServiceBase):
    @rpc(Unicode, Unicode, _returns = Iterable(Unicode))
    def separacion(ctx, dato_64):
        dato_real=dato_64.decode64()



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
    server.serve_forever()
