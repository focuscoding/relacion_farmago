import streamlit as st
import xmlrpc.client
import pandas as pd
import io
from datetime import date, timedelta
import unicodedata
import urllib.parse

# -----------------------------
# CONFIGURACIÓN GENERAL
# -----------------------------
st.set_page_config(page_title="Reporte Facturas", layout="wide")
st.title("📊Facturación Farmago")

# -----------------------------
# CLASE CONEXIÓN ODOO
# -----------------------------
class OdooClient:
    def __init__(self, url, db, username, password):
        self.url = url
        self.db = db
        self.username = username
        self.password = password

        self.common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common')
        self.uid = self.common.authenticate(db, username, password, {})

        if not self.uid:
            raise Exception("Error de autenticación en Odoo")

        self.models = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object')

    def search_read(self, model, domain, fields):
        return self.models.execute_kw(
            self.db,
            self.uid,
            self.password,
            model,
            'search_read',
            [domain],
            {'fields': fields}
        )

# -----------------------------
# FUNCIONES AUXILIARES
# -----------------------------
def procesar_facturas(data):
    if not data:
        return pd.DataFrame()

    df = pd.DataFrame(data)
    df["Moneda"] = df["currency_id"].apply(lambda x: x[1] if x else "")

    def obtener_impuesto(row):
        if row["Moneda"] == "Dolares":
            return row.get("amount_tax_usd", 0) or 0
        else:
            return row.get("amount_tax_bs", 0) or 0

    df["Impuesto"] = df.apply(obtener_impuesto, axis=1)
    df["Total Gravado"] = df["Impuesto"] / 0.16
    df["Exento"] = df["iva_exempt"].fillna(0)
    df["Total"] = df["Exento"] + df["Total Gravado"] + (df["Impuesto"] * 0.25)

    mask_rnc = df["name"].str.contains("RNCVTA", case=False, na=False)
    df.loc[mask_rnc, "Exento"] = -df.loc[mask_rnc, "Exento"]
    df.loc[mask_rnc, "Total Gravado"] = -df.loc[mask_rnc, "Total Gravado"]
    df.loc[mask_rnc, "Impuesto"] = -df.loc[mask_rnc, "Impuesto"]
    df.loc[mask_rnc, "Total"] = -df.loc[mask_rnc, "amount_total"]

    df_final = pd.DataFrame({
        "Empresa": "BLV",
        "Número": df["name"],
        "Fecha": df["invoice_date"],
        "Nro. Factura": df["invoice_number_next"],
        "Cliente": df["partner_id"].apply(lambda x: x[1] if x else ""),
        "Exento": df["Exento"],
        "Total Gravado": df["Total Gravado"],
        "Impuesto": df["Impuesto"],
        "Total": df["Total"],
        "Moneda": df["Moneda"]
    })

    return df_final

def calcular_resumen(df):

    

    resumen = (
        df.groupby(["Empresa", "Moneda"])["Total"]
        .sum()
        .round(2)
        .reset_index()
    )

    resumen_dict = {}

    for _, row in resumen.iterrows():
        empresa = row["Empresa"]
        moneda = row["Moneda"]
        total = row["Total"]

        resumen_dict[(empresa, moneda)] = total

    return resumen_dict

def generar_excel_formateado(df):
    resumen = calcular_resumen(df)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Reporte")
        workbook  = writer.book
        worksheet = writer.sheets["Reporte"]

        header_format = workbook.add_format({
            "bold": True,
            "border": 0,
            "align": "center",
            "valign": "vcenter",
        })

        dollar_format = workbook.add_format({"num_format": "$#,##0.00"})
        bs_format = workbook.add_format({"num_format": '"Bs." #,##0.00'})

        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        for row_num, moneda in enumerate(df["Moneda"], start=1):
            fmt = dollar_format if str(moneda).lower() == "dolares" else bs_format
            for col in range(5, 9):
                val = df.iloc[row_num-1, col]
                try:
                    val_num = float(val) if pd.notna(val) else 0
                except:
                    val_num = 0
                worksheet.write_number(row_num, col, val_num, fmt)
            worksheet.write(row_num, 9, df.iloc[row_num-1, 9])

        for i, col in enumerate(df.columns):

            # convertir todo a string evitando NaN
            column_data = df[col].astype(str).fillna("")

            # calcular longitud máxima entre header y valores
            max_len = max(
                column_data.map(len).max(),
                len(str(col))
            )

            # pequeño padding visual
            adjusted_width = max_len + 3

            worksheet.set_column(i, i, adjusted_width)

        # -----------------------------
        # RESUMEN EN EXCEL
        # -----------------------------

        bold_format = workbook.add_format({"bold": True})

        worksheet.write("L1", "BLV", bold_format)
        worksheet.write("L2", "Bolívares")
        worksheet.write("L3", "Dolares")

        worksheet.write("L4", "CRLV", bold_format)
        worksheet.write("L5", "Dolares")

        worksheet.write("M1", "Monto", bold_format)
        worksheet.write_number("M2", resumen.get(("BLV", "Bolívares"), 0), bs_format)
        worksheet.write_number("M3", resumen.get(("BLV", "Dolares"), 0), dollar_format)
        worksheet.write_number("M5", resumen.get(("CRLV", "Dolares"), 0), dollar_format)    

        worksheet.set_column("L:L", 12)
        worksheet.set_column("M:M", 18)
    return output.getvalue()

def limpiar_nombre(nombre):
    nombre = unicodedata.normalize('NFKD', nombre).encode('ASCII', 'ignore').decode('ASCII')
    nombre = nombre.replace(" ", "_")
    return nombre

def formato_moneda(valor, simbolo=""):
    if valor is None:
        valor = 0

    texto = f"{valor:,.2f}"

    # invertir separadores
    texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")

    return f"{simbolo} {texto}"

def construir_resumen_correo(resumen):

    blv_bs = formato_moneda(resumen.get(("BLV", "Bolívares"), 0), "Bs.")
    blv_usd = formato_moneda(resumen.get(("BLV", "Dolares"), 0), "$")
    crlv_usd = formato_moneda(resumen.get(("CRLV", "Dolares"), 0), "$")

    texto = f"""
Espero estén bien.

Comparto relación de la semana pasada.

BLV 
Bolívares: {blv_bs}
Dólares: {blv_usd}

CRLV
Dólares: {crlv_usd}

Saludos,

"""

    return texto

# -----------------------------
# INTERFAZ
# -----------------------------

hoy = date.today()

# lunes de esta semana
lunes_semana_actual = hoy - timedelta(days=hoy.weekday())

# lunes y viernes de la semana pasada
lunes_anterior = lunes_semana_actual - timedelta(days=7)
viernes_anterior = lunes_semana_actual - timedelta(days=3)

col1, col2 = st.columns(2)
with col1:
    fecha_inicio = st.date_input(
    "Fecha inicio",
    value=lunes_anterior,
    format="DD/MM/YYYY"
)
with col2:
    fecha_fin = st.date_input(
    "Fecha fin",
    value=viernes_anterior,
    format="DD/MM/YYYY"
)
# -----------------------------
# Inicializar session_state
# -----------------------------
if "df_final" not in st.session_state:
    st.session_state.df_final = None
if "excel_file" not in st.session_state:
    st.session_state.excel_file = None
if "nombre_archivo" not in st.session_state:
    st.session_state.nombre_archivo = None

# -----------------------------
# Botón Consultar Facturas
# -----------------------------
if st.button("🔍 Consultar Facturas"):
    try:
        config = st.secrets["odoo_bd1"]
        client = OdooClient(config["url"], config["db"], config["username"], config["password"])

        domain = [
            ('move_type', 'in', ['out_invoice', 'out_refund']),
            ('invoice_partner_display_name', '=', 'FARMACIA FARMAGO, C.A.'),
            ('invoice_date', '>=', str(fecha_inicio)),
            ('invoice_date', '<=', str(fecha_fin)),
            ('state', '=', 'posted')
        ]

        fields = [
            'name','invoice_date','invoice_number_next','partner_id',
            'iva_exempt','amount_tax_usd','amount_tax_bs','currency_id','amount_total'
        ]

        with st.spinner("Consultando Odoo..."):
            data = client.search_read('account.move', domain, fields)

        # -----------------------------
        # CONEXIÓN BD2
        # -----------------------------
        config2 = st.secrets["odoo_bd2"]
        client2 = OdooClient(config2["url"], config2["db"], config2["username"], config2["password"])

        domain_bd2 = [
            ('move_type', 'in', ['out_invoice', 'out_refund']),
            ('invoice_partner_display_name', '=', 'FARMACIA FARMAGO, C.A.'),
            ('invoice_date', '>=', str(fecha_inicio)),
            ('invoice_date', '<=', str(fecha_fin)),
            ('state', '=', 'posted')
        ]

        fields_bd2 = [
            'name',
            'invoice_date',
            'invoice_number_next',
            'partner_id',
            'amount_exento',
            'amount_untaxed_signed',
            'amount_tax_signed',
            'amount_total_signed',
            'currency_id',
            'tasa'
        ]

        data_bd2 = client2.search_read('account.move', domain_bd2, fields_bd2)



        st.session_state.df_final = procesar_facturas(data)

        # -----------------------------
        # PROCESAR BD2
        # -----------------------------
        if data_bd2:
            df_bd2 = pd.DataFrame(data_bd2)

            df_bd2["Moneda"] = df_bd2["currency_id"].apply(lambda x: x[1] if x else "")

            # Usar el campo 'tasa' si existe, si no asumimos 1
            if "tasa" in df_bd2.columns:
                tasa = df_bd2["tasa"].replace(0, 1)  # evitar dividir entre 0
            else:
                tasa = pd.Series([1]*len(df_bd2))

            exento = df_bd2["amount_exento"] / tasa
            total_gravado = df_bd2["amount_untaxed_signed"] / tasa
            impuesto = df_bd2["amount_tax_signed"] / tasa
            total_calculado = exento + total_gravado + (impuesto * 0.25)

            df_bd2_final = pd.DataFrame({
                "Empresa": "CRLV",
                "Número": df_bd2["name"],
                "Fecha": df_bd2["invoice_date"],
                "Nro. Factura": df_bd2["invoice_number_next"],
                "Cliente": df_bd2["partner_id"].apply(lambda x: x[1] if x else ""),
                "Exento": exento,
                "Total Gravado": total_gravado,
                "Impuesto": impuesto,
                "Total": total_calculado,
                "Moneda": "Dolares"
            })

            # -----------------------------
            # Ajuste especial para NC (usar amount_total_signed)
            # -----------------------------
            mask_nc = df_bd2["name"].str.contains("NC", case=False, na=False)

            df_bd2_final.loc[mask_nc, "Total"] = df_bd2.loc[mask_nc, "amount_total_signed"] / tasa[mask_nc]

            # -----------------------------
            # Redondear valores a 2 decimales
            # -----------------------------
            df_bd2_final[["Exento", "Total Gravado", "Impuesto", "Total"]] = \
            df_bd2_final[["Exento", "Total Gravado", "Impuesto", "Total"]].round(2)

            # Concatenar con BD1
            st.session_state.df_final = pd.concat(
                [st.session_state.df_final, df_bd2_final],
                ignore_index=True
            )




        if not st.session_state.df_final.empty:
            excel_bytesio = io.BytesIO()
            excel_bytesio.write(generar_excel_formateado(st.session_state.df_final))
            excel_bytesio.seek(0)

            st.session_state.excel_file = excel_bytesio

            nombre_archivo = f"Relación Farmago del {fecha_inicio.strftime('%d-%m-%Y')} al {fecha_fin.strftime('%d-%m-%Y')}.xlsx"
            st.session_state.nombre_archivo = limpiar_nombre(nombre_archivo)
    except Exception as e:
        st.error(f"Ocurrió un error: {str(e)}")

# -----------------------------
# Mostrar resultados y descarga
# -----------------------------
if st.session_state.df_final is not None and not st.session_state.df_final.empty:
    # Aseguramos df_final
    df_final = st.session_state.df_final

    st.success(f"Se encontraron {len(df_final)} registros.")
    
    # st.subheader("📄 Previsualización")
    # st.dataframe(df_final, width="stretch")


    # -----------------------------
    # FILTROS EXCLUSIONES ND
    # -----------------------------
    colA, colB, colC = st.columns(3)

    with colA:
        excluir_nd = st.checkbox("Excluir todas las ND?", value=False)

    with colB:
        exclusiones_input = st.text_input(
            "Excluir ND específicas (Nro. Nota, separadas por coma)",
            value="",
            disabled=excluir_nd
        )

    # with colC:
    #     frases_nd_input = st.text_input(
    #         "Excluir ND por frases (separadas por coma)",
    #         value="",
    #         disabled=excluir_nd or bool(exclusiones_input)
    #     )

    # -----------------------------
    # Convertir inputs a listas
    # -----------------------------
    exclusiones_list = [x.strip() for x in exclusiones_input.split(",") if x.strip()]
    # frases_nd_list = [x.strip() for x in frases_nd_input.split(",") if x.strip()]

    # -----------------------------
    # Aplicar filtros
    # -----------------------------
    df_filtrado = df_final.copy()

    if excluir_nd:
        # Excluir todas las filas cuyo 'name' contiene 'ND'
        df_filtrado = df_filtrado[~df_filtrado["Número"].str.contains("ND", case=False, na=False)]
    elif exclusiones_list:
        # Excluir filas donde 'name' contiene 'ND' y 'Nro. Factura' contiene alguno de los valores
        mask_excluir = df_filtrado["Número"].str.contains("ND", case=False, na=False) & \
                    df_filtrado["Nro. Factura"].astype(str).apply(lambda x: any(e in x for e in exclusiones_list))
        df_filtrado = df_filtrado[~mask_excluir]
    # elif frases_nd_list:
    #     # Excluir filas donde 'name' contiene 'ND' y además contiene alguna de las frases
    #     mask_frases = df_filtrado["Número"].str.contains("ND", case=False, na=False) & \
    #                 df_filtrado["Número"].astype(str).apply(lambda x: any(frase.lower() in x.lower() for frase in frases_nd_list))
    #     df_filtrado = df_filtrado[~mask_frases]

    st.subheader("📈 Resumen por Empresa y Moneda")

    resumen = calcular_resumen(df_filtrado)

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "BLV - Bolívares",
        formato_moneda(resumen.get(("BLV", "Bolívares"), 0), "Bs.")
    )

    col2.metric(
        "BLV - Dólares",
        formato_moneda(resumen.get(("BLV", "Dolares"), 0), "$")
    )

    col3.metric(
        "CRLV - Dólares",
        formato_moneda(resumen.get(("CRLV", "Dolares"), 0), "$")
    )

    resumen_correo = construir_resumen_correo(resumen)

    to = "mramos.farmago@gmail.com;staddeo@drogueriablv.com"
    cc = "vromero@drogueriablv.com"

    asunto = st.session_state.nombre_archivo.replace("_"," ")
    asunto = urllib.parse.quote(asunto)
    mensaje = urllib.parse.quote(resumen_correo)

    mailto_link = f"mailto:{to}?cc={cc}&subject={asunto}&body={mensaje}"



    # -----------------------------
    # Generar Excel
    # -----------------------------
    excel_bytesio = io.BytesIO()
    excel_bytesio.write(generar_excel_formateado(df_filtrado))
    excel_bytesio.seek(0)
    st.session_state.excel_file = excel_bytesio

    st.download_button(
        label="⬇️ Descargar Excel",
        data=st.session_state.excel_file,
        file_name=st.session_state.nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_excel"
    )
    
    st.link_button(
    "📧 Crear correo con resumen",
    mailto_link
    )