from modelo import Modelo
from vista import Vista
from controlador import Controlador

def main():
    modelo = Modelo()
    controlador = Controlador(modelo, None)  # Pasa None temporalmente
    vista = Vista(modelo, controlador)  # Ahora pasa la instancia de controlador

if __name__ == "__main__":
    main()
