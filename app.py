import streamlit as st
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import numpy as np

# =========================================================
# GENERADOR PDF ESTADO DE CUENTA - FORMATO PRUEBA.pdf
# =========================================================
# Estructura esperada del Excel:
# FECHA | CONCEPTO | SERIAL | RETIRO | DEPOSITO | SALDO EN PDF
#
# También acepta encabezados equivalentes:
# Fecha: DIA, FECHA, FECHA OPERACION
# Concepto: CONCEPTO, DESCRIPCION, MOVIMIENTO
# Serial: SERIAL, REFERENCIA, REF
# Retiro: RETIRO, RETIROS, CARGO, CARGOS
# Deposito: DEPOSITO, DEPOSITOS, ABONO, ABONOS
# Saldo en PDF: SALDO EN PDF, SALDO, SALDO FINAL
# =========================================================

# Carta: 8.5 x 11 pulgadas = 612 x 792 pt
PAGE_W_PT = 612.00
PAGE_H_PT = 792.00

# Fuente aproximada al PDF de prueba
FONT_NAME = "Helvetica"
FONT_SIZE = 8

# =========================================================
# POSICIONES HORIZONTALES EN PUNTOS
# =========================================================
X_DIA_PT = 43.20
X_CONCEPTO_PT = 61.20
X_SERIAL_PT = 159.00

X_RETIRO_RIGHT_PT = 402.60
X_DEPOSITO_RIGHT_PT = 473.20
X_SALDO_RIGHT_PT = 565.80

W_DIA_PT = 15.00
W_CONCEPTO_PT = 170.00
W_SERIAL_PT = 75.00
W_MONTO_PT = 70.00

X_RETIRO_PT = X_RETIRO_RIGHT_PT - W_MONTO_PT
X_DEPOSITO_PT = X_DEPOSITO_RIGHT_PT - W_MONTO_PT
X_SALDO_PT = X_SALDO_RIGHT_PT - W_MONTO_PT

# =========================================================
# POSICIONES VERTICALES EN PUNTOS
# =========================================================
Y_START_FIRST_PT = 506.80
Y_START_NORMAL_PT = 132.40
Y_END_PT = 721.00

ROW_H_PT = 18.72
CELL_H_PT = 8.20
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


def normalize_header(txt):
    txt = clean_cell(txt).upper()
    repl = {
        "Á": "A", "É": "E", "Í": "I", "Ó": "O", "Ú": "U", "Ü": "U", "Ñ": "N"
    }
    for a, b in repl.items():
        txt = txt.replace(a, b)
    txt = txt.replace("_", " ").replace("-", " ").replace(".", "")
    while "  " in txt:
        txt = txt.replace("  ", " ")
    return txt.strip()


def clean_day_or_date(val):
    if val is None:
        return ""

    if isinstance(val, (pd.Timestamp, datetime)):
        return f"{val.day:02d}"

    txt = clean_cell(val)
    if txt == "":
        return ""

    # Si viene como fecha de Excel serial
    try:
        f = float(txt)
        if f > 1000:
            dt = pd.to_datetime(f, unit="D", origin="1899-12-30", errors="coerce")
            if pd.notna(dt):
                return f"{dt.day:02d}"
        return f"{int(f):02d}"
    except Exception:
        pass

    # Si viene como texto fecha: dd/mm/aaaa, aaaa-mm-dd, etc.
    dt = pd.to_datetime(txt, dayfirst=True, errors="coerce")
    if pd.notna(dt):
        return f"{dt.day:02d}"

    # Último recurso: toma los primeros 2 dígitos
    only_digits = "".join(ch for ch in txt if ch.isdigit())
    if len(only_digits) >= 2:
        return only_digits[:2]

    return txt.zfill(2)


def clean_serial(val):
    txt = clean_cell(val)
    if txt == "":
        return ""
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

    if not (filename.endswith(".xlsx") or filename.endswith(".xlsm") or filename.endswith(".xls")):
        raise ValueError("Formato no compatible. Usa .xlsx, .xlsm o .xls.")

    try:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, engine="openpyxl", header=None, dtype=object)
    except Exception as err_openpyxl:
        try:
            uploaded_file.seek(0)
            return pd.read_excel(uploaded_file, engine="xlrd", header=None, dtype=object)
        except Exception as err_xlrd:
            raise ValueError(
                "No fue posible leer el Excel. Puede estar dañado, protegido o no ser un Excel válido.\n\n"
                f"Archivo: {filename}\n\n"
                f"Error OpenPyXL: {err_openpyxl}\n\n"
                f"Error XLRD: {err_xlrd}"
            )


def find_header_row(df):
    required_sets = [
        ["FECHA", "CONCEPTO", "SERIAL", "RETIRO", "DEPOSITO", "SALDO EN PDF"],
        ["FECHA", "CONCEPTO", "SERIAL", "RETIRO", "DEPOSITO", "SALDO"],
        ["DIA", "CONCEPTO", "REFERENCIA", "CARGO", "ABONO", "SALDO"],
    ]

    max_rows = min(len(df), 25)
    for idx in range(max_rows):
        row = [normalize_header(x) for x in df.iloc[idx].tolist()]
        row_set = set(row)
        for required in required_sets:
            hits = sum(1 for req in required if req in row_set)
            if hits >= 4:
                return idx
    return None


def first_existing_column(data, names):
    for name in names:
        if name in data.columns:
            return data[name]
    return ""


def parse_excel(df_raw):
    if df_raw.shape[1] < 6:
        raise ValueError(
            "El Excel debe tener al menos 6 columnas: FECHA, CONCEPTO, SERIAL, RETIRO, DEPOSITO y SALDO EN PDF."
        )

    header_row = find_header_row(df_raw)

    if header_row is not None:
        headers = [normalize_header(x) for x in df_raw.iloc[header_row].tolist()]
        data = df_raw.iloc[header_row + 1:].copy()
        data.columns = headers

        final = pd.DataFrame()
        final["FECHA"] = first_existing_column(data, ["FECHA", "DIA", "FECHA OPERACION", "FECHA DE OPERACION"])
        final["CONCEPTO"] = first_existing_column(data, ["CONCEPTO", "DESCRIPCION", "MOVIMIENTO", "DETALLE"])
        final["SERIAL"] = first_existing_column(data, ["SERIAL", "REFERENCIA", "REF", "NUMERO", "NO", "DOCUMENTO"])
        final["RETIRO"] = first_existing_column(data, ["RETIRO", "RETIROS", "CARGO", "CARGOS", "EGRESO", "EGRESOS", "DEBE"])
        final["DEPOSITO"] = first_existing_column(data, ["DEPOSITO", "DEPOSITOS", "ABONO", "ABONOS", "INGRESO", "INGRESOS", "HABER"])
        final["SALDO EN PDF"] = first_existing_column(data, ["SALDO EN PDF", "SALDO", "SALDO FINAL", "BALANCE"])

    else:
        final = df_raw.iloc[:, :6].copy()
        final.columns = ["FECHA", "CONCEPTO", "SERIAL", "RETIRO", "DEPOSITO", "SALDO EN PDF"]

    final = final.fillna("")

    for col in final.columns:
        final[col] = final[col].map(clean_cell)

    # Elimina encabezado duplicado si fue tomado como primera fila sin detectar encabezados
    final = final[~(
        final["FECHA"].map(normalize_header).eq("FECHA") &
        final["CONCEPTO"].map(normalize_header).eq("CONCEPTO")
    )]

    final = final[~(
        (final["FECHA"] == "") &
        (final["CONCEPTO"] == "") &
        (final["SERIAL"] == "") &
        (final["RETIRO"] == "") &
        (final["DEPOSITO"] == "") &
        (final["SALDO EN PDF"] == "")
    )].copy()

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

    def add_movement_row(self, fecha, concepto, serial, retiro, deposito, saldo):
        self.check_page_break()

        y = self.current_y
        self.set_font(FONT_NAME, "", FONT_SIZE)
        self.set_text_color(0, 0, 0)

        dia_str = clean_day_or_date(fecha)
        concepto_str = clean_cell(concepto)
        serial_str = clean_serial(serial)
        retiro_str = money_cell(retiro)
        deposito_str = money_cell(deposito)
        saldo_str = money_cell(saldo)

        self.text_cell(X_DIA_PT, y, W_DIA_PT, dia_str, "L")
        self.text_cell(X_CONCEPTO_PT, y, W_CONCEPTO_PT, concepto_str, "L")
        self.text_cell(X_SERIAL_PT, y, W_SERIAL_PT, serial_str, "L")
        self.text_cell(X_RETIRO_PT, y, W_MONTO_PT, retiro_str, "R")
        self.text_cell(X_DEPOSITO_PT, y, W_MONTO_PT, deposito_str, "R")
        self.text_cell(X_SALDO_PT, y, W_MONTO_PT, saldo_str, "R")

        self.current_y += ROW_H_PT
        self.rows_on_current_page += 1

# =========================================================
# STREAMLIT
# =========================================================

st.set_page_config(
    page_title="Generador PDF Estado de Cuenta",
    layout="wide",
    page_icon="📄"
)

st.title("📄 Generador PDF Estado de Cuenta")

st.markdown("""
Carga un Excel con esta estructura de **6 columnas**:

**Fecha | Concepto | SERIAL | Retiro | Deposito | Saldo en PDF**

Acepta archivos **.xlsx, .xlsm y .xls**.
""")

excel_file = st.file_uploader(
    "Sube tu archivo Excel",
    type=["xlsx", "xlsm", "xls"]
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
                        row["FECHA"],
                        row["CONCEPTO"],
                        row["SERIAL"],
                        row["RETIRO"],
                        row["DEPOSITO"],
                        row["SALDO EN PDF"]
                    )

                pdf_bytes = get_pdf_bytes(pdf)

                st.success("PDF generado correctamente.")

                st.download_button(
                    label="📥 Descargar PDF",
                    data=pdf_bytes,
                    file_name=f"Estado_Cuenta_{datetime.now():%Y%m%d_%H%M%S}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Error al generar el PDF: {e}")

    except Exception as e:
        st.error(f"Error al leer el Excel: {e}")

else:
    st.info("Sube un Excel para comenzar.")
