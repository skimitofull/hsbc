import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import numpy as np

# =========================================================
# GENERADOR PDF ESTADO DE CUENTA - FORMATO PRUEBA.pdf
# =========================================================
# Estructura esperada del Excel:
# DIA | CONCEPTO | REFERENCIA | CLAVE 1 | CLAVE 2 | CARGO | ABONO | SALDO
#
# - DIA: día del movimiento, ejemplo: 01
# - CONCEPTO: texto principal, ejemplo: PRESTAMO
# - REFERENCIA: número junto al concepto, ejemplo: 0010426
# - CLAVE 1: número superior de referencia, ejemplo: 08045221
# - CLAVE 2: número inferior de referencia, ejemplo: 80225
# - CARGO: salida / retiro
# - ABONO: entrada / depósito
# - SALDO: saldo final del movimiento
# =========================================================

MM_TO_PT = 2.83465

# Carta: 8.5 x 11 pulgadas, igual al PDF de prueba: 612 x 792 pt
PAGE_W_PT = 612.00
PAGE_H_PT = 792.00

# Fuente detectada/aproximada al PDF de prueba.
# El PDF original usa una fuente interna F6 de 8 pt; Helvetica 8 es la equivalencia más estable en FPDF.
FONT_NAME = "Helvetica"
FONT_SIZE = 8

# =========================================================
# POSICIONES HORIZONTALES EN PUNTOS
# Calibradas contra PRUEBA.pdf
# =========================================================

X_DIA_PT = 43.20
X_CONCEPTO_PT = 61.20
X_REFERENCIA_PT = 159.00

X_CLAVE_RIGHT_PT = 336.10
X_CARGO_RIGHT_PT = 402.60
X_ABONO_RIGHT_PT = 473.20
X_SALDO_RIGHT_PT = 565.80

W_DIA_PT = 15.00
W_CONCEPTO_PT = 95.00
W_REFERENCIA_PT = 58.00
W_CLAVE_PT = 45.00
W_MONTO_PT = 70.00

X_CLAVE_PT = X_CLAVE_RIGHT_PT - W_CLAVE_PT
X_CARGO_PT = X_CARGO_RIGHT_PT - W_MONTO_PT
X_ABONO_PT = X_ABONO_RIGHT_PT - W_MONTO_PT
X_SALDO_PT = X_SALDO_RIGHT_PT - W_MONTO_PT

# =========================================================
# POSICIONES VERTICALES EN PUNTOS
# Calibradas contra PRUEBA.pdf
# =========================================================

# Primera página inicia casi al final, como el documento de prueba.
Y_START_FIRST_PT = 506.80

# Segunda página y posteriores.
Y_START_NORMAL_PT = 132.40

# Último renglón útil observado en el PDF de prueba.
Y_END_PT = 721.00

# Cada movimiento ocupa dos líneas visuales.
ROW_H_PT = 18.72
SECOND_LINE_OFFSET_PT = 9.35
CELL_H_PT = 8.20

# En PRUEBA.pdf entran 12 movimientos en primera página.
MAX_ROWS_FIRST_PAGE = 12

# =========================================================
# FUNCIONES AUXILIARES
# =========================================================

def clean_cell(val):
    if val is None:
        return ""

    if isinstance(val, float) and np.isnan(val):
        return ""

    txt = str(val).strip()

    if txt.lower() in ["nan", "none", "null", "nat"]:
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
        return f"{int(float(txt)):02d}"
    except Exception:
        txt = txt.replace(".0", "")
        return txt.zfill(2)


def clean_ref(val):
    txt = clean_cell(val)

    if txt == "":
        return ""

    # Conserva ceros a la izquierda si vienen como texto.
    # Si Excel lo mandó como número, elimina .0.
    try:
        if txt.endswith(".0"):
            return str(int(float(txt)))
    except Exception:
        pass

    return txt


def money_cell(val):
    if val is None:
        return ""

    if isinstance(val, float) and np.isnan(val):
        return ""

    txt = str(val).strip()

    if txt.lower() in ["nan", "none", "null", "nat", ""]:
        return ""

    try:
        num = (
            txt.replace("$", "")
               .replace(",", "")
               .replace(" ", "")
               .strip()
        )

        fval = float(num)

        if np.isnan(fval) or abs(fval) < 0.005:
            return ""

        if fval < 0:
            return f"$ ({abs(fval):,.2f})"

        return f"$ {fval:,.2f}"

    except Exception:
        return clean_cell(val)


def read_excel_file(uploaded_file):
    filename = uploaded_file.name.lower()

    if filename.endswith(".xls"):
        return pd.read_excel(uploaded_file, engine="xlrd", header=None, dtype=object)

    if filename.endswith(".xlsx") or filename.endswith(".xlsm"):
        return pd.read_excel(uploaded_file, engine="openpyxl", header=None, dtype=object)

    raise ValueError("Formato de archivo no compatible. Usa .xls, .xlsx o .xlsm.")


def normalize_header(txt):
    txt = clean_cell(txt).upper()
    repl = {
        "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U", "Ü": "U", "Ñ": "N"
    }
    for a, b in repl.items():
        txt = txt.replace(a, b)
    txt = txt.replace("_", " ").replace("-", " ")
    while "  " in txt:
        txt = txt.replace("  ", " ")
    return txt.strip()


def find_header_row(df):
    required = ["DIA", "CONCEPTO", "SALDO"]

    max_rows = min(len(df), 20)

    for idx in range(max_rows):
        row = [normalize_header(x) for x in df.iloc[idx].tolist()]
        hits = sum(1 for req in required if req in row)
        if hits >= 2:
            return idx

    return None


def parse_excel(df_raw):
    if df_raw.shape[1] < 8:
        raise ValueError(
            "El Excel debe tener al menos 8 columnas: DIA, CONCEPTO, REFERENCIA, CLAVE 1, CLAVE 2, CARGO, ABONO y SALDO."
        )

    header_row = find_header_row(df_raw)

    if header_row is not None:
        headers = [normalize_header(x) for x in df_raw.iloc[header_row].tolist()]
        data = df_raw.iloc[header_row + 1:].copy()
        data.columns = headers

        aliases = {
            "DIA": ["DIA", "FECHA DIA"],
            "CONCEPTO": ["CONCEPTO", "DESCRIPCION", "MOVIMIENTO", "DETALLE"],
            "REFERENCIA": ["REFERENCIA", "REF", "NUMERO", "NO", "DOCUMENTO"],
            "CLAVE 1": ["CLAVE 1", "CLAVE1", "AUTORIZACION 1", "REFERENCIA 1", "LINEA 1"],
            "CLAVE 2": ["CLAVE 2", "CLAVE2", "AUTORIZACION 2", "REFERENCIA 2", "LINEA 2"],
            "CARGO": ["CARGO", "RETIRO", "RETIROS", "EGRESO", "EGRESOS", "DEBE"],
            "ABONO": ["ABONO", "DEPOSITO", "DEPOSITOS", "INGRESO", "INGRESOS", "HABER"],
            "SALDO": ["SALDO", "SALDO FINAL", "BALANCE"]
        }

        final = pd.DataFrame()

        for final_col, possible_names in aliases.items():
            found = None
            for name in possible_names:
                if name in data.columns:
                    found = name
                    break
            final[final_col] = data[found] if found is not None else ""

    else:
        # Si no trae encabezados, toma las primeras 8 columnas en el orden requerido.
        final = df_raw.iloc[:, :8].copy()
        final.columns = [
            "DIA", "CONCEPTO", "REFERENCIA", "CLAVE 1", "CLAVE 2", "CARGO", "ABONO", "SALDO"
        ]

    final = final.fillna("")

    for col in final.columns:
        final[col] = final[col].map(clean_cell)

    final = final[
        ~(
            (final["DIA"] == "") &
            (final["CONCEPTO"] == "") &
            (final["REFERENCIA"] == "") &
            (final["CLAVE 1"] == "") &
            (final["CLAVE 2"] == "") &
            (final["CARGO"] == "") &
            (final["ABONO"] == "") &
            (final["SALDO"] == "")
        )
    ].copy()

    return final.reset_index(drop=True)


def get_pdf_bytes(pdf):
    pdf_output = pdf.output(dest="S")

    if isinstance(pdf_output, str):
        return pdf_output.encode("latin1")

    return bytes(pdf_output)

# =========================================================
# CLASE PDF
# =========================================================

class EstadoCuentaPDF(FPDF):

    def __init__(self):
        super().__init__(unit="pt", format=(PAGE_W_PT, PAGE_H_PT))
        self.set_auto_page_break(False)
        self.alias_nb_pages()
        self.current_y = Y_START_FIRST_PT
        self.rows_on_current_page = 0

    def header(self):
        self.set_font(FONT_NAME, "", FONT_SIZE)
        self.set_text_color(0, 0, 0)

        if self.page_no() == 1:
            self.current_y = Y_START_FIRST_PT
        else:
            self.current_y = Y_START_NORMAL_PT

        self.rows_on_current_page = 0

    def footer(self):
        pass

    def check_page_break(self):
        if self.page_no() == 1 and self.rows_on_current_page >= MAX_ROWS_FIRST_PAGE:
            self.add_page()
            self.current_y = Y_START_NORMAL_PT
            self.rows_on_current_page = 0
            return

        if self.current_y + ROW_H_PT > Y_END_PT:
            self.add_page()
            self.current_y = Y_START_NORMAL_PT
            self.rows_on_current_page = 0

    def text_cell(self, x, y, w, txt, align="L"):
        self.set_xy(x, y)
        self.cell(w, CELL_H_PT, txt, border=0, ln=0, align=align)

    def add_movement_row(self, dia, concepto, referencia, clave1, clave2, cargo, abono, saldo):
        self.check_page_break()

        y1 = self.current_y
        y2 = self.current_y + SECOND_LINE_OFFSET_PT

        self.set_font(FONT_NAME, "", FONT_SIZE)
        self.set_text_color(0, 0, 0)

        dia_str = clean_day(dia)
        concepto_str = clean_cell(concepto)
        referencia_str = clean_ref(referencia)
        clave1_str = clean_ref(clave1)
        clave2_str = clean_ref(clave2)
        cargo_str = money_cell(cargo)
        abono_str = money_cell(abono)
        saldo_str = money_cell(saldo)

        # Línea superior del movimiento
        self.text_cell(X_DIA_PT, y1, W_DIA_PT, dia_str, "L")
        self.text_cell(X_CONCEPTO_PT, y1, W_CONCEPTO_PT, concepto_str, "L")
        self.text_cell(X_REFERENCIA_PT, y1, W_REFERENCIA_PT, referencia_str, "L")
        self.text_cell(X_CLAVE_PT, y1, W_CLAVE_PT, clave1_str, "R")
        self.text_cell(X_CARGO_PT, y1, W_MONTO_PT, cargo_str, "R")
        self.text_cell(X_ABONO_PT, y1, W_MONTO_PT, abono_str, "R")
        self.text_cell(X_SALDO_PT, y1, W_MONTO_PT, saldo_str, "R")

        # Línea inferior del movimiento: segunda clave/referencia debajo de la primera.
        if clave2_str:
            self.text_cell(X_CLAVE_PT, y2, W_CLAVE_PT, clave2_str, "R")

        self.current_y += ROW_H_PT
        self.rows_on_current_page += 1

# =========================================================
# STREAMLIT
# =========================================================

st.set_page_config(
    page_title="Generador PDF Estado de Cuenta - Formato PRUEBA",
    layout="wide",
    page_icon="📄"
)

st.title("📄 Generador PDF Estado de Cuenta - Formato PRUEBA")

st.markdown("""
Carga un Excel con esta estructura:

**DIA | CONCEPTO | REFERENCIA | CLAVE 1 | CLAVE 2 | CARGO | ABONO | SALDO**

El PDF se genera con posiciones, fuente e interlineado calibrados contra el documento **PRUEBA.pdf**.
""")

excel_file = st.file_uploader(
    "Sube tu archivo Excel",
    type=["xlsx", "xls", "xlsm"]
)

if excel_file:
    try:
        df_raw = read_excel_file(excel_file)
        df = parse_excel(df_raw)

        st.success(f"Archivo cargado correctamente: {len(df)} filas útiles.")
        st.dataframe(df.head(30), use_container_width=True)

        if st.button("Generar PDF", type="primary", use_container_width=True):
            try:
                pdf = EstadoCuentaPDF()
                pdf.add_page()

                for _, row in df.iterrows():
                    pdf.add_movement_row(
                        row["DIA"],
                        row["CONCEPTO"],
                        row["REFERENCIA"],
                        row["CLAVE 1"],
                        row["CLAVE 2"],
                        row["CARGO"],
                        row["ABONO"],
                        row["SALDO"]
                    )

                pdf_bytes = get_pdf_bytes(pdf)

                st.success("PDF generado correctamente.")

                st.download_button(
                    label="📥 Descargar PDF",
                    data=pdf_bytes,
                    file_name=f"Estado_Cuenta_PRUEBA_{datetime.now():%Y%m%d_%H%M%S}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Error al generar el PDF: {e}")

    except Exception as e:
        st.error(f"Error al leer el Excel: {e}")

else:
    st.info("Sube un Excel para comenzar.")
