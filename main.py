# main.py

import logging
from evaluator import setup_logging, gui

def main():
    """
    Función principal que configura el logging e inicia la interfaz gráfica.
    """
    try:
        # 1. Configurar el sistema de registro de actividad
        setup_logging()

        # 2. Crear una instancia de la aplicación principal y ejecutarla
        app = gui.MainApp()
        app.mainloop()

    except Exception as e:
        logging.critical("Ha ocurrido un error fatal al iniciar la aplicación.", exc_info=True)
        try:
            from tkinter import messagebox
            messagebox.showerror("Error Fatal", f"No se pudo iniciar la aplicación:\n\n{e}")
        except:
            print(f"ERROR FATAL: No se pudo iniciar la aplicación: {e}")

if __name__ == "__main__":
    main()