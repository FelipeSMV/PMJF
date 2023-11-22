from controlador import Controlador
from vista import Vista

if __name__ == "__main__":
    vista = Vista()
    controlador = Controlador(vista)
    vista.establecer_controlador(controlador)
    controlador.iniciar_aplicacion()
