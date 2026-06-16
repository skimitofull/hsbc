import streamlit as st
import pandas as pd
import numpy as np
from fpdf import FPDF
from datetime import datetime
import io

# =========================================================
# CONFIGURACIÓN GENERAL - FORMATO PRUEBA.pdf
# =========================================================

PAGE_W_PT = 612.0
PAGE_H_PT = 792.0

FONT_NAME = "Helvetica"
FONT_SIZE = 7.0

# Coordenadas tomadas del PDF PRUEBA.pdf
X_DIA_PT = 43.2
X_CONCEPTO_PT = 61.2
X_SERIAL_PT = 300.4

X_RETIRO_RIGHT_PT = 401.4
X_DEPOSITO_RIGHT_PT = 472.0
X_SALDO_RIGHT_PT = 564.6

W_DIA_PT = 18
W_CONCEPTO_PT = X_SERIAL_PT - X_CONCEPTO_PT - 4
W_SERIAL_PT = 38
W_RETIRO_PT = 75
W_DEPOSITO_PT = 75
W_SALDO_PT = 85

X_RETIRO_PT = X_RETIRO_RIGHT_PT - W_RETIRO_PT
X_DEPOSITO_PT = X_DEPOSITO_RIGHT_PT - W_DEPOSITO_PT
X_SALDO_PT = X_SALDO_RIGHT_PT - W_SALDO_PT

# Verticales
Y_START_FIRST_PT = 506.8
Y_START_NORMAL_PT = 50.0
Y_END_PT = 730.0

LINE_H_PT = 18.72        # separación entre movimientos
SERIAL_LINE_H_PT = 9.36  # segunda línea del serial
CELL_H_PT = 8.5

# =========================================================
# FUNCIONES AUXILIARES
# =========================================================

def clean_cell(val):
    if val is None:
        return ""
    if isinstance(val, float) and np.isnan(val):
        return ""
    txt = str(val).strip()
    if txt.lower() in ["nan", "none", "null"]:
        return ""
    txt = txt.replace("\r", " ").replace("\n", " ")
    while "  " in txt:
        txt = txt.replace("  ", " ")
    return txt


def clean_day(val):
    txt = clean_cell(val)
    if txt == "":
        return ""
    try:
        if isinstance(val, (pd.Timestamp, datetime)):
            return f"{val.day:02d}"
        f = float(txt)
        return f"{int(f):02d}"
    except Exception:
        return txt[:2].zfill(2) if txt[:2].isdigit() else txt


def clean_serial(val):
    txt = clean_cell(val)
    if txt == "":
        return ""
    try:
        f = float(txt)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return txt


def money_cell(val):
    if val is None:
        return ""
    if isinstance(val, float) and np.isnan(val):
        return ""
    if isinstance(val, str) and val.strip().lower() in ["", "nan", "none", "null"]:
        return ""
    try:
        if isinstance(val, str):
            val = val.replace("$", "").replace(",", "").replace(" ", "").strip()
        num = float(val)
        if np.isnan(num) or abs(num) < 0.000001:
            return ""
        return f"$ {num:,.2f}"
    except Exception:
        return clean_cell(val)


def read_excel_file(uploaded_file):
    data = uploaded_file.read()

    # 1) Intenta como xlsx/xlsm aunque la extensión sea .xls
    try:
        return pd.read_excel(io.BytesIO(data), engine="openpyxl", header=None)
    except Exception as err_openpyxl:
        # 2) Intenta xls real
        try:
            return pd.read_excel(io.BytesIO(data), engine="xlrd", header=None)
        except Exception as err_xlrd:
            raise ValueError(
                "No fue posible leer el Excel. "
                f"Error openpyxl: {err_openpyxl}. "
                f"Error xlrd: {err_xlrd}."
            )


def parse_excel(df_raw):
    if df_raw.shape[1] < 6:
        raise ValueError(
            "El Excel debe tener 6 columnas: Fecha, Concepto, SERIAL, Retiro, Deposito y Saldo en PDF."
        )

    df = df_raw.iloc[:, :6].copy()
    df.columns = ["FECHA", "CONCEPTO", "SERIAL", "RETIRO", "DEPOSITO", "SALDO"]

    # Quita fila de encabezados si existe
    primera = " ".join([clean_cell(x).upper() for x in df.iloc[0].tolist()]) if len(df) else ""
    if "FECHA" in primera and "CONCEPTO" in primera:
        df = df.iloc[1:].copy()

    df = df.replace([np.nan, "nan", "NaN", "None", "NULL", "null"], "")

    movimientos = []
    for _, row in df.iterrows():
        fecha = clean_cell(row["FECHA"])
        concepto = clean_cell(row["CONCEPTO"])
        serial = clean_serial(row["SERIAL"])
        retiro = row["RETIRO"]
        deposito = row["DEPOSITO"]
        saldo = row["SALDO"]

        # Si es la segunda línea del serial, se pega al movimiento anterior.
        if fecha == "" and concepto == "" and serial != "":
            if movimientos:
                movimientos[-1]["SERIAL_2"] = serial
            continue

        # Ignora filas totalmente vacías
        if fecha == "" and concepto == "" and serial == "" and clean_cell(retiro) == "" and clean_cell(deposito) == "" and clean_cell(saldo) == "":
            continue

        movimientos.append({
            "DIA": row["FECHA"],
            "CONCEPTO": concepto,
            "SERIAL_1": serial,
            "SERIAL_2": "",
            "RETIRO": retiro,
            "DEPOSITO": deposito,
            "SALDO": saldo,
        })

    if not movimientos:
        raise ValueError("No se encontraron movimientos útiles en el Excel.")

    return pd.DataFrame(movimientos)


def get_pdf_bytes(pdf):
    out = pdf.output(dest="S")
    if isinstance(out, str):
        return out.encode("latin1")
    return bytes(out)

# =========================================================
# CLASE PDF
# =========================================================

class EstadoCuentaPDF(FPDF):
    def __init__(self):
        super().__init__(unit="pt", format=(PAGE_W_PT, PAGE_H_PT))
        self.set_auto_page_break(False)
        self.current_y = Y_START_FIRST_PT

    def header(self):
        self.set_font(FONT_NAME, "", FONT_SIZE)
        self.current_y = Y_START_FIRST_PT if self.page_no() == 1 else Y_START_NORMAL_PT

    def footer(self):
        pass

    def check_page_break(self):
        if self.current_y + LINE_H_PT > Y_END_PT:
            self.add_page()
            self.current_y = Y_START_NORMAL_PT

    def add_movement_row(self, dia, concepto, serial1, serial2, retiro, deposito, saldo):
        self.check_page_break()
        y = self.current_y

        self.set_font(FONT_NAME, "", FONT_SIZE)
        self.set_text_color(0, 0, 0)

        dia_str = clean_day(dia)
        concepto_str = clean_cell(concepto)
        serial1_str = clean_serial(serial1)
        serial2_str = clean_serial(serial2)
        retiro_str = money_cell(retiro)
        deposito_str = money_cell(deposito)
        saldo_str = money_cell(saldo)

        # DÍA
        self.set_xy(X_DIA_PT, y)
        self.cell(W_DIA_PT, CELL_H_PT, dia_str, border=0, align="L")

        # CONCEPTO
        self.set_xy(X_CONCEPTO_PT, y)
        self.cell(W_CONCEPTO_PT, CELL_H_PT, concepto_str, border=0, align="L")

        # SERIAL en dos líneas dentro del mismo movimiento
        self.set_xy(X_SERIAL_PT, y)
        self.cell(W_SERIAL_PT, CELL_H_PT, serial1_str, border=0, align="L")

        if serial2_str:
            self.set_xy(X_SERIAL_PT + 13.0, y + SERIAL_LINE_H_PT)
            self.cell(W_SERIAL_PT, CELL_H_PT, serial2_str, border=0, align="L")

        # RETIRO
        self.set_xy(X_RETIRO_PT, y)
        self.cell(W_RETIRO_PT, CELL_H_PT, retiro_str, border=0, align="R")

        # DEPÓSITO
        self.set_xy(X_DEPOSITO_PT, y)
        self.cell(W_DEPOSITO_PT, CELL_H_PT, deposito_str, border=0, align="R")

        # SALDO
        self.set_xy(X_SALDO_PT, y)
        self.cell(W_SALDO_PT, CELL_H_PT, saldo_str, border=0, align="R")

        self.current_y += LINE_H_PT

# =========================================================
# STREAMLIT
# =========================================================

st.set_page_config(
    page_title="Generador PDF Estado de Cuenta",
    layout="wide",
    page_icon="📄"
)

st.title("📄 Generador de Estado de Cuenta")

st.markdown("""
Carga un Excel con 6 columnas en este orden:

**Fecha | Concepto | SERIAL | Retiro | Deposito | Saldo en PDF**

La segunda línea del **SERIAL** puede venir en la fila siguiente con Fecha y Concepto vacíos; el sistema la acomodará dentro del mismo movimiento.
""")

excel_file = st.file_uploader(
    "Sube tu archivo Excel",
    type=["xlsx", "xlsm", "xls"]
)

if excel_file:
    try:
        df_raw = read_excel_file(excel_file)
        df = parse_excel(df_raw)

        st.success(f"Archivo cargado correctamente: {len(df)} movimientos útiles.")
        st.dataframe(df.head(30), use_container_width=True)

        if st.button("Generar PDF", type="primary", use_container_width=True):
            pdf = EstadoCuentaPDF()
            pdf.add_page()

            for _, row in df.iterrows():
                pdf.add_movement_row(
                    row["DIA"],
                    row["CONCEPTO"],
                    row["SERIAL_1"],
                    row["SERIAL_2"],
                    row["RETIRO"],
                    row["DEPOSITO"],
                    row["SALDO"],
                )

            pdf_bytes = get_pdf_bytes(pdf)

            st.success("PDF generado correctamente.")
            st.download_button(
                label="📥 Descargar PDF",
                data=pdf_bytes,
                file_name=f"Estado_Cuenta_{datetime.now():%Y%m%d_%H%M%S}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"Error al leer el Excel: {e}")
else:
    st.info("Sube un Excel para comenzar.")
