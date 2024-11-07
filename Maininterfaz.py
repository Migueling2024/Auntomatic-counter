import sys
import os
import pandas as pd
import re
import random  # Para clasificación aleatoria
import shutil #mover y copiar archivos
import matplotlib.pyplot as plt #para realizar la grafica
import difflib  # Para calcular la similitud entre nombres de archivos
from PyQt6.QtGui import QIcon #interfaz
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, 
    QMessageBox, QMainWindow, QFileDialog, QDialog, QScrollArea,QInputDialog
)
from PyQt6.QtCore import Qt, QFile, QTextStream
from estilos import aplicar_estilos #importar la función de aplicar estilos

class Validador:
    def __init__(self, archivo_excel):
        self.archivo_excel = archivo_excel
        self.datos = self.cargar_datos()
        self.usuario_actual = None

    def cargar_datos(self):
        try:
            df = pd.read_excel(self.archivo_excel)
            return df
        except FileNotFoundError:
            print("Archivo no encontrado.")
            return pd.DataFrame()

    def validar_credenciales(self, correo, contraseña):
        self.datos['correo'] = self.datos['correo'].astype(str)
        self.datos['contraseña'] = self.datos['contraseña'].astype(str)
        usuario = self.datos[(self.datos['correo'] == correo) & (self.datos['contraseña'] == contraseña)]
        if not usuario.empty:
            self.usuario_actual = {
                'correo': usuario.iloc[0]['correo'],
                'contraseña': usuario.iloc[0]['contraseña'],
                'documento': usuario.iloc[0]['documento identidad']
            }
            return True
        return False

class VentanaInicioSesion(QMainWindow):
    def __init__(self, validador):
        super().__init__()
        self.validador = validador
        self.init_ui()
        self.setObjectName("ventanaPrincipal")
        self.setStyleSheet("""
            QMainWindow#ventanaPrincipal {
                background-image: url(fondopoo.jpg);
            }
        """)

    def init_ui(self):
        self.setWindowTitle('Inicio de Sesión')
        self.setGeometry(100, 100, 400, 250)
        self.setWindowIcon(QIcon('Logo.jpg'))
        # Crear un QWidget central que contenga el layout
        widget_central = QWidget(self)
        widget_central.setObjectName("widgetCentral")  # Asigna un nombre para estilos específicos
        # Aplicar el CSS desde el archivo externo utilizando la función importada
        aplicar_estilos(widget_central)
        aplicar_estilos(self)
        layout = QVBoxLayout(widget_central)

        # Añadir los elementos al layout
        self.label_correo = QLabel('Correo:')
        self.label_correo.setObjectName("labelCorreo")  # Para estilos específicos de este QLabel
        layout.addWidget(self.label_correo)

        self.input_correo = QLineEdit()
        layout.addWidget(self.input_correo)

        self.label_contraseña = QLabel('Contraseña:')
        self.label_contraseña.setObjectName("labelContraseña")
        layout.addWidget(self.label_contraseña)

        self.input_contraseña = QLineEdit()
        self.input_contraseña.setEchoMode(QLineEdit.EchoMode.Password)
        layout.addWidget(self.input_contraseña)

        self.boton_iniciar_sesion = QPushButton('Iniciar Sesión')
        self.boton_iniciar_sesion.clicked.connect(self.iniciar_sesion)
        layout.addWidget(self.boton_iniciar_sesion)

        # Establecer el widget central
        self.setCentralWidget(widget_central)

    def iniciar_sesion(self):
        correo = self.input_correo.text()
        contraseña = self.input_contraseña.text()

        if self.validador.validar_credenciales(correo, contraseña):
            documento = self.validador.usuario_actual['documento']
            QMessageBox.information(self, 'Éxito', f'Bienvenido, documento: {documento}')
            self.close()
            self.abrir_dashboard()
        else:
            QMessageBox.warning(self, 'Error', 'Correo o contraseña incorrectos.')

    def abrir_dashboard(self):
        self.dashboard = VentanaDashboard(self.validador.usuario_actual)
        self.dashboard.show()

class VentanaDashboard(QMainWindow):
    def __init__(self, usuario):
        super().__init__()
        self.usuario = usuario
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Dashboard')
        self.setGeometry(50, 50, 1200, 800)

        layout_principal = QVBoxLayout()

        # Aplicar el CSS desde el archivo externo utilizando la función importada
        aplicar_estilos(self)
        #Poner el logo
        self.setWindowIcon(QIcon('Logo.jpg'))

        self.label_documento = QLabel(f"Documento: {self.usuario['documento']}")
        self.label_documento.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_principal.addWidget(self.label_documento)

        layout_central = QVBoxLayout()
        #boton escanear factura
        self.boton_escanear_factura = QPushButton("Escanear factura")
        self.boton_escanear_factura.setFixedSize(200, 50)
        self.boton_escanear_factura.clicked.connect(self.escanear_factura)
        layout_central.addWidget(self.boton_escanear_factura)
        #boton administrar archivos
        self.boton_administrar_archivos = QPushButton("Administrar Archivos")
        self.boton_administrar_archivos.setFixedSize(200, 50)
        self.boton_administrar_archivos.clicked.connect(self.abrir_administrar_archivos)
        layout_central.addWidget(self.boton_administrar_archivos)
        #boton facturas sin pagar
        self.boton_añadir_factura_sin_pagar = QPushButton("Facturas sin Pagar")
        self.boton_añadir_factura_sin_pagar.setFixedSize(200, 50)
        self.boton_añadir_factura_sin_pagar.clicked.connect(self.abrir_administrar_facturas_sin_pagar)
        layout_central.addWidget(self.boton_añadir_factura_sin_pagar)

        # Botón "Informe Personalizado"
        self.boton_informe_personalizado = QPushButton("Informe Personalizado")
        self.boton_informe_personalizado.setFixedSize(200, 50)
        self.boton_informe_personalizado.clicked.connect(self.abrir_informe_personalizado)
        layout_central.addWidget(self.boton_informe_personalizado)

        layout_central.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout_principal.addLayout(layout_central)

        widget_central = QWidget()
        widget_central.setLayout(layout_principal)
        self.setCentralWidget(widget_central)

    def abrir_administrar_archivos(self):
        self.ventana_administrar = VentanaAdministrarArchivos()
        self.ventana_administrar.exec()
    
    def abrir_administrar_facturas_sin_pagar(self):
        print("oe")
        self.ventana_administrar_facturas = AdministrarFacturasSinPagar()
        self.ventana_administrar_facturas.exec()

    def abrir_informe_personalizado(self):
        self.ventana_informe = VentanaInformePersonalizado()
        self.ventana_informe.exec()
    


    def escanear_factura(self):
        archivo, _ = QFileDialog.getOpenFileName(self, "Seleccionar Factura", "", "Archivos de Excel (*.xlsx *.xls)")
        if archivo:
            tipo_factura, subgrupo, celdas_vacias = self.analizar_factura(archivo)
            if tipo_factura:
                if celdas_vacias:
                    self.mostrar_celdas_vacias(celdas_vacias)
                else:
                    m=self.guardar_factura(archivo, tipo_factura, subgrupo)
                    if m=="moved":
                        QMessageBox.information(self, "Factura Escaneada", f"La factura ha sido guardada en '{tipo_factura}/{subgrupo}'.")

            else:
                self.guardar_factura(archivo, "requiere revisión", None)
                QMessageBox.warning(self, "Requiere Revisión", "El archivo fue guardado en 'requiere revisión' y requiere verificación manual.")

    def analizar_factura(self, archivo):
        df = pd.read_excel(archivo)
        df.columns = df.columns.str.lower()
        celdas_vacias = []
        celdas_fecha_incorrecta = []
        tipo_factura = None
        subgrupo = None
        if df.map(lambda x: 'novedades' in str(x).lower()).any().any():
            tipo_factura = "Novedades"
        elif df.map(lambda x: 'auxilio' in str(x).lower()).any().any():
            tipo_factura = "facturas de nomina"
        else:
            if 'grupo' in df.columns and df.columns.get_loc("grupo") == 14:
                tipo_factura = "facturas de compra" if df.iloc[0, 14].lower() == "recibido" else "facturas de venta" if df.iloc[0, 14].lower() == "emitido" else None
                if tipo_factura == "facturas de compra":
                    subgrupo = random.choice(["venta de materiales", "venta de servicios"])

            for columna in df.columns:
                vacias = df[columna].isna()
                if vacias.any():
                    celdas_vacias.extend([(fila + 1, columna) for fila in df.index[vacias]])

        # Validación de formato en columnas de fecha
        if tipo_factura!="Novedades" and tipo_factura!="facturas de nomina":
            for columna in df.columns:
                if "fecha" in columna:
                    for index, valor in df[columna].dropna().items():  # Itera solo sobre valores no nulos
                        if not re.match(r"\b\d{2}-\d{2}-\d{4}\b", str(valor)):  # Regex para DD-MM-YYYY
                            celdas_fecha_incorrecta.append((index + 1, columna))

        # Mostrar advertencia si hay errores en las fechas
        if celdas_fecha_incorrecta:
            mensaje_error_fecha = "\n".join([f"Fila: {fila}, Columna: '{columna}'" for fila, columna in celdas_fecha_incorrecta])
            QMessageBox.warning(self, "Formato de Fecha Incorrecto", f"Las siguientes celdas tienen un formato de fecha incorrecto:\n{mensaje_error_fecha}")


        return tipo_factura, subgrupo, celdas_vacias

    def mostrar_celdas_vacias(self, celdas_vacias):
        mensaje = "\n".join([f"Fila: {fila+1}, Columna: '{columna}'" for fila, columna in celdas_vacias])
        ventana_vacias = QMessageBox(self)
        ventana_vacias.setWindowTitle("Celdas Vacías Encontradas")
        ventana_vacias.setText(mensaje)
        ventana_vacias.exec()

    
    
    def guardar_factura(self, archivo, tipo_factura, subgrupo):
        directorio_destino = os.path.join(os.getcwd(), tipo_factura)
        if subgrupo:
            directorio_destino = os.path.join(directorio_destino, subgrupo)
        os.makedirs(directorio_destino, exist_ok=True)
        
        nombre_archivo = os.path.basename(archivo)
        nuevo_path = os.path.join(directorio_destino, nombre_archivo)

        # Verificar si el archivo ya existe
        if os.path.exists(nuevo_path):
            # Mostrar mensaje y cerrar la ventana actual
            QMessageBox.information(self, "Archivo existente", f"El archivo '{nombre_archivo}' ya está ubicado en '{tipo_factura}/{subgrupo}'.")
            return "alreadyexist"  # Salir de la función sin mover el archivo

        # Si no existe, mover el archivo a la ruta de destino
        os.rename(archivo, nuevo_path)
        return "moved"

class VentanaAdministrarArchivos(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Administrar Archivos")
        self.setGeometry(200, 200, 400, 300)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        # Estilos CSS
        aplicar_estilos(self)
        # Poner el logo
        self.setWindowIcon(QIcon('Logo.jpg'))

        # Botones para las carpetas principales
        self.boton_facturas_compra = QPushButton("Facturas de Compra")
        self.boton_facturas_compra.clicked.connect(lambda: self.mostrar_archivos("facturas de compra"))
        layout.addWidget(self.boton_facturas_compra)

        self.boton_facturas_venta = QPushButton("Facturas de Venta")
        self.boton_facturas_venta.clicked.connect(lambda: self.mostrar_archivos("facturas de venta"))
        layout.addWidget(self.boton_facturas_venta)

        self.boton_facturas_nomina = QPushButton("Facturas de Nómina")
        self.boton_facturas_nomina.clicked.connect(lambda: self.mostrar_archivos("facturas de nomina"))
        layout.addWidget(self.boton_facturas_nomina)

        self.boton_sin_clasificar = QPushButton("Sin Clasificar")
        self.boton_sin_clasificar.clicked.connect(lambda: self.mostrar_archivos("requiere revisión"))
        layout.addWidget(self.boton_sin_clasificar)

        self.setLayout(layout)

    def mostrar_archivos(self, carpeta):
        # Ruta completa de la carpeta
        directorio = os.path.join(os.getcwd(), carpeta)

        # Verificar si la carpeta existe
        if not os.path.exists(directorio):
            QMessageBox.information(self, "Directorio no encontrado", f"No se encontró la carpeta '{carpeta}'.")
            return

        # Crear un diálogo para mostrar los archivos
        dialog_archivos = QDialog(self)
        dialog_archivos.setWindowTitle(f"Archivos en '{carpeta}'")
        layout_archivos = QVBoxLayout()

        # Listar archivos y subcarpetas en la carpeta actual
        archivos_excel = [f for f in os.listdir(directorio) if os.path.isfile(os.path.join(directorio, f)) and f.endswith((".xlsx", ".xls"))]
        subcarpetas = [d for d in os.listdir(directorio) if os.path.isdir(os.path.join(directorio, d))]

        # Agregar botones para los archivos si existen
        if archivos_excel:
            for archivo in archivos_excel:
                ruta_archivo = os.path.join(directorio, archivo)
                boton_archivo = QPushButton(archivo)
                boton_archivo.clicked.connect(lambda _, archivo=ruta_archivo: self.mostrar_opcionesdelarchivo(archivo))
                layout_archivos.addWidget(boton_archivo)
        elif subcarpetas:
            # Si no hay archivos, muestra las subcarpetas para navegar
            for subcarpeta in subcarpetas:
                boton_subcarpeta = QPushButton(f"Abrir '{subcarpeta}'")
                boton_subcarpeta.clicked.connect(lambda _, subcarpeta=subcarpeta: self.mostrar_archivos(os.path.join(carpeta, subcarpeta)))
                layout_archivos.addWidget(boton_subcarpeta)
        else:
            # Si no hay archivos ni subcarpetas
            layout_archivos.addWidget(QPushButton("No hay archivos ni subcarpetas"))

        dialog_archivos.setLayout(layout_archivos)
        dialog_archivos.exec()

    def mostrar_opcionesdelarchivo(self, ruta_archivo):
        print("aca q mrd paso")
        # Agregar opción para mover el archivo
        respuesta = QMessageBox.question(self, "Mover archivo", "¿Deseas mover este archivo?", 
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if respuesta == QMessageBox.StandardButton.Yes:
            self.mover_archivo(ruta_archivo)

    def mostrar_archivos_subcarpeta(self, subcarpeta):
        # Esta función permite mostrar los archivos dentro de una subcarpeta específica
        if not os.path.exists(subcarpeta):
            QMessageBox.information(self, "Directorio no encontrado", f"No se encontró la subcarpeta '{subcarpeta}'.")
            return

        dialog_archivos = QDialog(self)
        dialog_archivos.setWindowTitle(f"Archivos en '{subcarpeta}'")
        layout_archivos = QVBoxLayout()

        # Mostrar archivos dentro de la subcarpeta
        for archivo in os.listdir(subcarpeta):
            if archivo.endswith(".xlsx") or archivo.endswith(".xls"):
                boton_archivo = QPushButton(archivo)
                boton_archivo.clicked.connect(lambda _, archivo=archivo: self.mostrar_opcionesdelarchivo(os.path.join(subcarpeta, archivo)))
                layout_archivos.addWidget(boton_archivo)

        dialog_archivos.setLayout(layout_archivos)
        dialog_archivos.exec()

    def mover_archivo(self, archivo_origen):
        # Abrir un diálogo para seleccionar la carpeta de destino
        carpeta_destino = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta de destino")
        if carpeta_destino:
            try:
                # Obtener el nombre del archivo y su nueva ruta
                nombre_archivo = os.path.basename(archivo_origen)
                archivo_destino = os.path.join(carpeta_destino, nombre_archivo)
                
                # Mover el archivo
                shutil.move(archivo_origen, archivo_destino)
                QMessageBox.information(self, "Archivo movido", f"El archivo se ha movido a {archivo_destino}")
            except Exception as e:
                QMessageBox.critical(self, "Error al mover archivo", f"No se pudo mover el archivo: {e}")

class AdministrarFacturasSinPagar(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Administrar Facturas sin Pagar")
        self.setGeometry(200, 200, 600, 400)
        
        # Definir carpeta y archivo de facturas
        self.carpeta_facturas = "Facturas sin pagar"
        self.carpeta_facturas_pagadas = "Facturas_pagadas"
        self.archivo_excel = os.path.join(self.carpeta_facturas, "facturas_sin_pagar.xlsx")
        
        # Crear carpeta si no existe
        self.verificar_o_crear_carpeta(self.carpeta_facturas)
        self.verificar_o_crear_carpeta(self.carpeta_facturas_pagadas)

        # Inicializar interfaz
        self.init_ui()
    
    def verificar_o_crear_carpeta(self, carpeta):
        try:
            if not os.path.exists(carpeta):
                os.makedirs(carpeta)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al crear carpeta: {e}")

    def init_ui(self):
        self.layout = QVBoxLayout()
        # Estilos CSS
        aplicar_estilos(self)
        # Poner el logo
        self.setWindowIcon(QIcon('Logo.jpg'))
        
        # Botón para añadir factura
        self.boton_anadir_factura = QPushButton("Añadir Factura")
        self.boton_anadir_factura.clicked.connect(self.anadir_factura)
        self.layout.addWidget(self.boton_anadir_factura)
        
        # Botón para analizar facturas próximas a vencerse
        self.boton_analizar_facturas = QPushButton("Analizar Facturas sin Pagar")
        self.boton_analizar_facturas.clicked.connect(self.analizar_facturas)
        self.layout.addWidget(self.boton_analizar_facturas)
        
        # Contenedor para los botones de archivos de factura
        self.contenedor_botones = QVBoxLayout()
        self.layout.addLayout(self.contenedor_botones)

        # Mostrar lista de archivos en la carpeta de facturas
        self.actualizar_lista_archivos()

        self.setLayout(self.layout)

    def anadir_factura(self):
        archivo, _ = QFileDialog.getOpenFileName(self, "Seleccionar Factura", "", "Archivos de Excel (*.xlsx *.xls)")
        if archivo: 
            nombre_archivo = os.path.basename(archivo)
            nuevo_path = os.path.join(self.carpeta_facturas, nombre_archivo)
            # Solicitar fecha de vencimiento
            fecha_vencimiento, ok = QInputDialog.getText(self, "Fecha de Vencimiento", "Ingrese la fecha de vencimiento (YYYY-MM-DD):")
            if ok:
                try:
                    # Verificar que la fecha no esté vacía
                    if not fecha_vencimiento:
                        raise ValueError("La fecha no puede estar vacía.")
                    
                    # Intentar convertir la fecha
                    fecha_vencimiento_parsed = pd.to_datetime(fecha_vencimiento, errors='raise', format='%Y-%m-%d')
                    
                    # Verificar si la fecha es válida
                    if pd.isna(fecha_vencimiento_parsed):
                        raise ValueError("La fecha no es válida.")
                    
                    # Guardar la factura solo si la fecha es válida
                    if fecha_vencimiento:
                        self.guardar_factura(nombre_archivo, fecha_vencimiento_parsed.date())

                    #copiar el archivo en la ruta
                    shutil.copy(archivo, nuevo_path)
                    # Actualizar la lista de archivos después de guardar la factura
                    self.actualizar_lista_archivos()

                except ValueError as e:
                    # Mostrar un mensaje de error con la excepción que ocurrió
                    QMessageBox.warning(self, "Fecha Inválida", str(e))

    
    def guardar_factura(self, nombre_archivo, fecha_vencimiento):
        try:
            # Comprobar si el archivo de facturas existe y cargarlo
            if os.path.exists(self.archivo_excel):
                df = pd.read_excel(self.archivo_excel)
            else:
                df = pd.DataFrame(columns=["Archivo", "Fecha"])

            # Añadir la nueva factura
            nueva_factura = pd.DataFrame([[nombre_archivo, fecha_vencimiento]], columns=["Archivo", "Fecha"])
            df = pd.concat([df, nueva_factura], ignore_index=True)
            
            # Guardar los cambios en el archivo Excel
            df.to_excel(self.archivo_excel, index=False)
            
            # Mensaje de confirmación
            QMessageBox.information(self, "Éxito", "Factura añadida correctamente.")
        
        except Exception as e:
            # Mostrar mensaje de error si ocurre algún problema al guardar la factura
            QMessageBox.critical(self, "Error", f"No se pudo guardar la factura: {e}")

    def analizar_facturas(self):
        try:
            if os.path.exists(self.archivo_excel):
                df = pd.read_excel(self.archivo_excel)
                df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")

                facturas_sin_fecha = df[df["Fecha"].isna()]
                if not facturas_sin_fecha.empty:
                    mensaje = "\n".join([f"Archivo: {row['Archivo']}" for _, row in facturas_sin_fecha.iterrows()])
                    QMessageBox.warning(self, "Facturas sin Fecha", f"Estas facturas no tienen fecha:\n{mensaje}")
                else:
                    # Mostrar las 5 facturas más cercanas a vencerse
                    df = df.dropna(subset=["Fecha"])
                    df = df.sort_values(by="Fecha")
                    próximas_vencidas = df.head(5)
                    mensaje = "\n".join([f"Factura: {row['Archivo']}, Fecha: {row['Fecha'].date()}" for _, row in próximas_vencidas.iterrows()])
                    QMessageBox.information(self, "Próximas a Vencerse", mensaje)
            else:
                QMessageBox.warning(self, "Sin Facturas", "No hay facturas para analizar.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al analizar facturas: {e}")

    def actualizar_lista_archivos(self):
        # Limpiar el contenedor de botones de archivos
        for i in reversed(range(self.contenedor_botones.count())):
            widget = self.contenedor_botones.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)
        
        archivos = os.listdir(self.carpeta_facturas)
        if archivos:
            for archivo in archivos:
            # Verificar que el archivo no contenga "facturas_sin_pagar" en el nombre
                if archivo and "facturas_sin_pagar" not in archivo and "~$" not in archivo:
                    boton_archivo = QPushButton(archivo)
                    boton_archivo.clicked.connect(lambda _, archivo=archivo: self.confirmar_pago(archivo))
                    self.contenedor_botones.addWidget(boton_archivo)
        else:
            etiqueta = QLabel("No hay facturas sin pagar.")
            self.contenedor_botones.addWidget(etiqueta)
    
    def confirmar_pago(self, archivo):
        respuesta = QMessageBox.question(self, "Confirmación de Pago", f"¿Ya se realizó el pago de la factura {archivo}?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if respuesta == QMessageBox.StandardButton.Yes:
            # Mover el archivo a la carpeta de facturas pagadas
            self.mover_factura_pagada(archivo)
            
            # Intentar actualizar el archivo de Excel eliminando la factura pagada
            try:
                if os.path.exists(self.archivo_excel):
                    # Leer el archivo Excel
                    df = pd.read_excel(self.archivo_excel)
                    
                    # Filtrar el DataFrame para eliminar la fila con el archivo pagado
                    df = df[df['Archivo'] != archivo]
                    
                    # Guardar el DataFrame actualizado
                    df.to_excel(self.archivo_excel, index=False)
                    
                    QMessageBox.information(self, "Factura Actualizada", f"La factura {archivo} ha sido eliminada del registro de facturas sin pagar.")
                else:
                    QMessageBox.warning(self, "Archivo No Encontrado", "No se encontró el archivo de facturas sin pagar.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al actualizar el archivo Excel: {e}")
            
            # Actualizar la lista de archivos después de confirmar el pago
            self.actualizar_lista_archivos()

    def mover_factura_pagada(self, archivo):
        try:
            origen = os.path.join(self.carpeta_facturas, archivo)
            destino = os.path.join(self.carpeta_facturas_pagadas, archivo)
            
            # Intentar mover el archivo
            shutil.move(origen, destino)
            
            # Actualizar la lista de archivos después de mover
            self.actualizar_lista_archivos()
            
            # Mensaje de confirmación
            QMessageBox.information(self, "Factura Movida", f"La factura {archivo} ha sido movida a la carpeta de facturas pagadas.")
        
        except Exception as e:
            # Mostrar mensaje de error si ocurre un problema al mover el archivo
            QMessageBox.critical(self, "Error", f"No se pudo mover la factura {archivo}: {e}")

class VentanaInformePersonalizado(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Informe Personalizado")
        self.setGeometry(100, 100, 800, 600)

        # Estilos CSS
        aplicar_estilos(self)
        # Poner el logo
        self.setWindowIcon(QIcon('Logo.jpg'))
        # Layout
        layout = QVBoxLayout()
        self.label = QLabel("Seleccione un archivo Excel para generar el informe.")
        layout.addWidget(self.label)

        self.boton_seleccionar_archivo = QPushButton("Seleccionar Archivo")
        self.boton_seleccionar_archivo.clicked.connect(self.analizar_archivo)
        layout.addWidget(self.boton_seleccionar_archivo)

        self.setLayout(layout)

    def analizar_archivo(self):
        archivo, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo de Ventas", "", "Archivos de Excel (*.xlsx *.xls)")
        if archivo:
            df = pd.read_excel(archivo)
            df.columns = df.columns.str.lower()  # Convertir columnas a minúsculas

            # Verificar que existan las columnas requeridas
            if 'fecha recepción' in df.columns and 'total' in df.columns:
                # Convertir la columna de fechas y agrupar por mes
                df['fecha recepción'] = pd.to_datetime(df['fecha recepción'], errors='coerce')
                df.dropna(subset=['fecha recepción'], inplace=True)
                df['mes'] = df['fecha recepción'].dt.to_period("M")
                ventas_mensuales = df.groupby('mes')['total'].sum()

                # Generar la gráfica de barras
                plt.figure(figsize=(10, 5))
                ventas_mensuales.plot(kind='bar', color='skyblue')
                plt.title("Ventas Mensuales")
                plt.xlabel("Mes")
                plt.ylabel("Total de Ventas")
                plt.xticks(rotation=45)
                plt.tight_layout()
                plt.show()
            else:
                QMessageBox.warning(self, "Error", "El archivo seleccionado no contiene las columnas requeridas ('fecha recepción' y 'total').")


class Aplicacion(QApplication):
    def __init__(self, sys_argv):
        super().__init__(sys_argv)
        self.validador = Validador('api.xlsx')
        self.ventana_inicio = VentanaInicioSesion(self.validador)
        self.ventana_inicio.show()

if __name__ == '__main__':
    app = Aplicacion(sys.argv)
    sys.exit(app.exec())
