# main.py

import logging
from evaluator import setup_logging, gui

def main():
    """
    Función principal que configura el logging e inicia la interfaz gráfica.
    """
    try:
        setup_logging()
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