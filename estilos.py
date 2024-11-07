# estilos.py

from PyQt6.QtCore import QFile, QTextStream

def aplicar_estilos(widget, archivo_css="estilos.css"):
    """
    Aplica estilos CSS desde un archivo al widget proporcionado.
    :param widget: El widget al que se le aplicarán los estilos.
    :param archivo_css: El archivo CSS que contiene los estilos (por defecto 'estilos.css').
    """
    archivo_css = QFile(archivo_css)
    if archivo_css.open(QFile.OpenModeFlag.ReadOnly):
        archivo = QTextStream(archivo_css)
        widget.setStyleSheet(archivo.readAll())
        archivo_css.close()
