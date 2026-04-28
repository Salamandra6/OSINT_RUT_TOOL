import os
import re
import sys
import threading
from pathlib import Path
from datetime import datetime

import customtkinter as ctk
import pandas as pd
import pdfplumber
import pytesseract

from pdf2image import convert_from_path
from tkinter import filedialog, messagebox
from tkinter import ttk


# ============================================================
# CONFIGURACION PORTABLE
# ============================================================

def get_base_dir():
    """Devuelve la carpeta base del programa."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


BASE_DIR = get_base_dir()
RESULTADOS_DIR = BASE_DIR / "resultados"
TESSERACT_DIR = BASE_DIR / "tesseract"
POPPLER_DIR = BASE_DIR / "poppler"

RESULTADOS_DIR.mkdir(exist_ok=True)


def configurar_tesseract():
    """
    Configura Tesseract de forma portable.

    Estructura esperada:
    programa/
    ├── OSINT_RUT_GUI.exe
    ├── tesseract/tesseract.exe
    └── poppler/Library/bin/
    """
    tesseract_exe = TESSERACT_DIR / "tesseract.exe"
    if tesseract_exe.exists():
        pytesseract.pytesseract.tesseract_cmd = str(tesseract_exe)


def get_poppler_path():
    """Devuelve la ruta portable de Poppler si existe."""
    posibles_rutas = [
        POPPLER_DIR / "Library" / "bin",
        POPPLER_DIR / "bin",
        POPPLER_DIR,
    ]
    for ruta in posibles_rutas:
        if ruta.exists():
            return str(ruta)
    return None


configurar_tesseract()


# ============================================================
# LIMPIEZA Y VALIDACION DE RUT
# ============================================================

PALABRAS_IRRELEVANTES = {
    "rut", "run", "cedula", "cédula", "identidad", "nacional",
    "republica", "república", "chile", "firma", "documento",
    "fecha", "nacimiento", "domicilio", "direccion", "dirección",
    "comuna", "region", "región", "sexo", "estado", "civil",
    "telefono", "teléfono", "correo", "email", "sr", "sra",
    "señor", "señora", "nombre", "nombres", "apellido", "apellidos",
    "profesion", "profesión", "actividad", "vigencia", "serie",
    "emision", "emisión", "vencimiento", "numero", "número",
}


def limpiar_texto(texto):
    """Normaliza espacios y caracteres innecesarios."""
    if not texto:
        return ""
    texto = texto.replace("\n", " ").replace("\t", " ")
    texto = re.sub(r"\s+", " ", texto)
    return texto.strip()


def normalizar_rut(rut):
    """Convierte distintos formatos a formato estandar sin puntos."""
    rut = rut.upper().replace(".", "").replace(" ", "")

    if "-" in rut:
        cuerpo, dv = rut.split("-", 1)
        cuerpo = re.sub(r"\D", "", cuerpo)
        dv = re.sub(r"[^0-9K]", "", dv)
        if cuerpo and dv:
            return f"{cuerpo}-{dv}"

    solo = re.sub(r"[^0-9K]", "", rut)
    if len(solo) >= 9:
        cuerpo = solo[:-1]
        dv = solo[-1]
        if cuerpo.isdigit() and (dv.isdigit() or dv == "K"):
            return f"{cuerpo}-{dv}"

    return rut


def calcular_dv(cuerpo):
    """Calcula el digito verificador de un RUT chileno."""
    suma = 0
    multiplicador = 2

    for numero in reversed(cuerpo):
        suma += int(numero) * multiplicador
        multiplicador += 1
        if multiplicador > 7:
            multiplicador = 2

    dv = 11 - (suma % 11)
    if dv == 11:
        return "0"
    if dv == 10:
        return "K"
    return str(dv)


def validar_rut(rut):
    """Valida un RUT con digito verificador."""
    rut = normalizar_rut(rut)
    if "-" not in rut:
        return False

    cuerpo, dv = rut.split("-")
    if not cuerpo.isdigit() or not 7 <= len(cuerpo) <= 8:
        return False

    return calcular_dv(cuerpo) == dv.upper()


def formatear_rut(rut):
    """Entrega el RUT en formato con puntos y guion."""
    rut = normalizar_rut(rut)
    if "-" not in rut:
        return rut

    cuerpo, dv = rut.split("-")
    return f"{int(cuerpo):,}".replace(",", ".") + f"-{dv}"


# ============================================================
# EXTRACCION DE TEXTO DESDE PDF
# ============================================================

def extraer_texto_pdf_digital(pdf_path):
    """Intenta extraer texto desde un PDF digital usando pdfplumber."""
    texto_total = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, pagina in enumerate(pdf.pages, start=1):
                texto = pagina.extract_text() or ""
                texto_total += f"\n--- Pagina {i} ---\n{texto}"
    except Exception as e:
        raise RuntimeError(f"Error leyendo PDF digital: {e}")

    return limpiar_texto(texto_total)


def extraer_texto_pdf_ocr(pdf_path, log_callback=None):
    """Aplica OCR a un PDF escaneado usando pdf2image y pytesseract."""
    texto_total = ""
    poppler_path = get_poppler_path()

    try:
        if log_callback:
            log_callback("Convirtiendo PDF escaneado a imagenes...")

        paginas = convert_from_path(pdf_path, dpi=300, poppler_path=poppler_path)

        for i, imagen in enumerate(paginas, start=1):
            if log_callback:
                log_callback(f"Aplicando OCR en pagina {i}...")
            texto = pytesseract.image_to_string(imagen, lang="spa+eng")
            texto_total += f"\n--- Pagina {i} ---\n{texto}"

    except Exception as e:
        raise RuntimeError(f"Error aplicando OCR: {e}")

    return limpiar_texto(texto_total)


def obtener_texto_pdf(pdf_path, log_callback=None):
    """Primero intenta lectura digital; si no hay texto suficiente, aplica OCR."""
    if log_callback:
        log_callback("Intentando lectura digital del PDF...")

    texto = extraer_texto_pdf_digital(pdf_path)

    if len(texto.strip()) >= 50:
        if log_callback:
            log_callback("PDF digital detectado. Texto extraido correctamente.")
        return texto, "Digital"

    if log_callback:
        log_callback("No se detecto texto suficiente. Se aplicara OCR automaticamente.")

    texto_ocr = extraer_texto_pdf_ocr(pdf_path, log_callback=log_callback)
    return texto_ocr, "OCR"


# ============================================================
# DETECCION DE RUT Y NOMBRES ASOCIADOS
# ============================================================

def detectar_ruts(texto):
    """Detecta RUTs en distintos formatos."""
    patrones = [
        r"\b\d{1,2}\.\d{3}\.\d{3}-[\dkK]\b",
        r"\b\d{7,8}-[\dkK]\b",
        r"\b\d{8,9}\b",
    ]

    encontrados = []
    for patron in patrones:
        for match in re.finditer(patron, texto):
            encontrados.append({
                "rut_original": match.group(),
                "inicio": match.start(),
                "fin": match.end(),
            })

    unicos = []
    vistos = set()
    for item in encontrados:
        clave = (item["rut_original"], item["inicio"], item["fin"])
        if clave not in vistos:
            vistos.add(clave)
            unicos.append(item)

    return unicos


def limpiar_nombre(nombre):
    """Limpia nombres candidatos eliminando ruido contextual."""
    if not nombre:
        return ""

    nombre = nombre.replace(":", " ").replace("-", " ")
    nombre = re.sub(r"[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ\s]", " ", nombre)
    nombre = re.sub(r"\s+", " ", nombre).strip()

    palabras_limpias = []
    for palabra in nombre.split():
        if palabra.lower() in PALABRAS_IRRELEVANTES:
            continue
        if len(palabra) <= 1:
            continue
        palabras_limpias.append(palabra)

    partes = palabras_limpias[-6:]
    return " ".join(partes).strip()


def buscar_nombre_cercano(texto, inicio_rut, fin_rut):
    """Busca un posible nombre cerca del RUT."""
    ventana_antes = texto[max(0, inicio_rut - 180):inicio_rut]
    ventana_despues = texto[fin_rut:min(len(texto), fin_rut + 180)]
    contexto = ventana_antes + " " + ventana_despues

    candidatos = []
    patrones_nombre = [
        r"(?:nombre(?:s)?|apellido(?:s)?|titular|persona|don|doña)\s*[:\-]?\s*([A-ZÁÉÍÓÚÜÑ][A-Za-zÁÉÍÓÚÜÑáéíóúüñ\s]{5,80})",
        r"([A-ZÁÉÍÓÚÜÑ]{2,}(?:\s+[A-ZÁÉÍÓÚÜÑ]{2,}){1,5})",
    ]

    for patron in patrones_nombre:
        for match in re.findall(patron, contexto, flags=re.IGNORECASE):
            nombre = limpiar_nombre(match)
            if nombre and len(nombre.split()) >= 2:
                candidatos.append(nombre)

    if candidatos:
        candidatos = sorted(candidatos, key=lambda x: len(x.split()), reverse=True)
        return candidatos[0]

    posible = limpiar_nombre(ventana_antes)
    partes = posible.split()
    if len(partes) >= 2:
        return " ".join(partes[-4:])

    return "No identificado"


def procesar_texto(texto, nombre_archivo):
    """Detecta, normaliza, valida RUTs y busca nombres cercanos."""
    resultados = []
    vistos = set()

    for item in detectar_ruts(texto):
        rut_original = item["rut_original"]
        rut_normalizado = normalizar_rut(rut_original)
        tiene_dv = "-" in rut_normalizado

        if tiene_dv:
            rut_valido = validar_rut(rut_normalizado)
            estado = "Valido" if rut_valido else "Invalido"
            rut_mostrar = formatear_rut(rut_normalizado)
        else:
            estado = "Sin DV / No validable"
            rut_mostrar = rut_original

        clave = (rut_mostrar, item["inicio"])
        if clave in vistos:
            continue
        vistos.add(clave)

        resultados.append({
            "Archivo": nombre_archivo,
            "RUT detectado": rut_original,
            "RUT normalizado": rut_mostrar,
            "Estado validacion": estado,
            "Nombre asociado": buscar_nombre_cercano(texto, item["inicio"], item["fin"]),
        })

    return resultados


# ============================================================
# EXPORTACION A EXCEL
# ============================================================

def exportar_excel(resultados):
    """Exporta resultados a archivo Excel en carpeta resultados."""
    if not resultados:
        raise ValueError("No existen resultados para exportar.")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    salida = RESULTADOS_DIR / f"resultados_osint_rut_{timestamp}.xlsx"
    pd.DataFrame(resultados).to_excel(salida, index=False)
    return salida


# ============================================================
# INTERFAZ GRAFICA
# ============================================================

class OSINTRutApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Herramienta OSINT RUT Chile - Laboratorio Ciberinteligencia")
        self.geometry("1150x720")
        self.minsize(1000, 650)

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.pdf_path = None
        self.resultados = []
        self.ultimo_excel = None

        self.crear_interfaz()

    def crear_interfaz(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        self.grid_rowconfigure(4, weight=1)

        titulo = ctk.CTkLabel(
            self,
            text="Herramienta OSINT para deteccion de RUT en PDF",
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        titulo.grid(row=0, column=0, padx=20, pady=(15, 5), sticky="w")

        subtitulo = ctk.CTkLabel(
            self,
            text="Procesamiento local: PDF digital + OCR automatico + validacion DV + exportacion Excel",
            font=ctk.CTkFont(size=14),
        )
        subtitulo.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="w")

        frame_botones = ctk.CTkFrame(self)
        frame_botones.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        frame_botones.grid_columnconfigure(3, weight=1)

        self.btn_cargar = ctk.CTkButton(frame_botones, text="Cargar PDF", command=self.cargar_pdf, width=160)
        self.btn_cargar.grid(row=0, column=0, padx=10, pady=10)

        self.btn_procesar = ctk.CTkButton(frame_botones, text="Iniciar / Procesar", command=self.iniciar_procesamiento, width=160)
        self.btn_procesar.grid(row=0, column=1, padx=10, pady=10)

        self.btn_guardar = ctk.CTkButton(frame_botones, text="Guardar resultados", command=self.guardar_resultados_manual, width=160)
        self.btn_guardar.grid(row=0, column=2, padx=10, pady=10)

        self.lbl_archivo = ctk.CTkLabel(frame_botones, text="Ningun PDF cargado", anchor="w")
        self.lbl_archivo.grid(row=0, column=3, padx=10, pady=10, sticky="ew")

        frame_tabla = ctk.CTkFrame(self)
        frame_tabla.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")
        frame_tabla.grid_rowconfigure(0, weight=1)
        frame_tabla.grid_columnconfigure(0, weight=1)

        columnas = ("Archivo", "RUT detectado", "RUT normalizado", "Estado validacion", "Nombre asociado")
        self.tabla = ttk.Treeview(frame_tabla, columns=columnas, show="headings")

        for col in columnas:
            self.tabla.heading(col, text=col)
            self.tabla.column(col, width=210, anchor="w")

        self.tabla.grid(row=0, column=0, sticky="nsew")

        scrollbar_y = ttk.Scrollbar(frame_tabla, orient="vertical", command=self.tabla.yview)
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        self.tabla.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(frame_tabla, orient="horizontal", command=self.tabla.xview)
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        self.tabla.configure(xscrollcommand=scrollbar_x.set)

        frame_logs = ctk.CTkFrame(self)
        frame_logs.grid(row=4, column=0, padx=20, pady=(10, 20), sticky="nsew")
        frame_logs.grid_rowconfigure(1, weight=1)
        frame_logs.grid_columnconfigure(0, weight=1)

        lbl_logs = ctk.CTkLabel(frame_logs, text="Logs de proceso", font=ctk.CTkFont(size=15, weight="bold"))
        lbl_logs.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="w")

        self.txt_logs = ctk.CTkTextbox(frame_logs, height=150)
        self.txt_logs.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        self.log("Aplicacion iniciada correctamente.")
        self.log(f"Carpeta base: {BASE_DIR}")
        self.log(f"Carpeta resultados: {RESULTADOS_DIR}")

    def log(self, mensaje):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_logs.insert("end", f"[{timestamp}] {mensaje}\n")
        self.txt_logs.see("end")
        self.update_idletasks()

    def cargar_pdf(self):
        archivo = filedialog.askopenfilename(title="Seleccionar PDF", filetypes=[("Archivos PDF", "*.pdf")])
        if archivo:
            self.pdf_path = Path(archivo)
            self.lbl_archivo.configure(text=str(self.pdf_path.name))
            self.log(f"PDF cargado: {self.pdf_path}")

    def limpiar_tabla(self):
        for item in self.tabla.get_children():
            self.tabla.delete(item)

    def cargar_resultados_en_tabla(self):
        self.limpiar_tabla()
        for fila in self.resultados:
            self.tabla.insert("", "end", values=(
                fila.get("Archivo", ""),
                fila.get("RUT detectado", ""),
                fila.get("RUT normalizado", ""),
                fila.get("Estado validacion", ""),
                fila.get("Nombre asociado", ""),
            ))

    def iniciar_procesamiento(self):
        if not self.pdf_path:
            messagebox.showwarning("Advertencia", "Primero debes cargar un PDF.")
            return
        threading.Thread(target=self.procesar_pdf, daemon=True).start()

    def procesar_pdf(self):
        try:
            self.btn_procesar.configure(state="disabled")
            self.resultados = []
            self.ultimo_excel = None
            self.limpiar_tabla()

            self.log("Iniciando procesamiento...")
            self.log(f"Archivo: {self.pdf_path.name}")

            texto, metodo = obtener_texto_pdf(str(self.pdf_path), log_callback=self.log)

            if not texto.strip():
                self.log("No se pudo extraer texto del PDF.")
                messagebox.showerror("Error", "No se pudo extraer texto del PDF.")
                return

            self.log(f"Metodo utilizado: {metodo}")
            self.log("Buscando RUTs en el contenido...")

            self.resultados = procesar_texto(texto, self.pdf_path.name)
            self.cargar_resultados_en_tabla()

            total = len(self.resultados)
            self.log(f"Total de posibles RUT encontrados: {total}")

            if total > 0:
                self.ultimo_excel = exportar_excel(self.resultados)
                self.log(f"Resultados exportados automaticamente: {self.ultimo_excel}")
                messagebox.showinfo("Proceso terminado", f"Se encontraron {total} registros.\n\nExcel generado:\n{self.ultimo_excel}")
            else:
                self.log("No se encontraron RUTs en el documento.")
                messagebox.showinfo("Proceso terminado", "No se encontraron RUTs en el documento.")

        except Exception as e:
            self.log(f"Error: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self.btn_procesar.configure(state="normal")

    def guardar_resultados_manual(self):
        try:
            if not self.resultados:
                messagebox.showwarning("Advertencia", "No hay resultados para guardar.")
                return

            salida = exportar_excel(self.resultados)
            self.ultimo_excel = salida
            self.log(f"Resultados guardados manualmente: {salida}")
            messagebox.showinfo("Guardado correcto", f"Archivo generado:\n{salida}")

        except Exception as e:
            self.log(f"Error al guardar resultados: {e}")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    app = OSINTRutApp()
    app.mainloop()
