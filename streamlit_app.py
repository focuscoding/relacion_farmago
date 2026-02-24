import streamlit as st
import xmlrpc.client
import pandas as pd
import io
from datetime import date
import unicodedata

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

def generar_excel_formateado(df):
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
            max_len = max(df[col].astype(str).map(len).max(), len(col))
            worksheet.set_column(i, i, max_len + 2)

    return output.getvalue()

def limpiar_nombre(nombre):
    nombre = unicodedata.normalize('NFKD', nombre).encode('ASCII', 'ignore').decode('ASCII')
    nombre = nombre.replace(" ", "_")
    return nombre

# -----------------------------
# INTERFAZ
# -----------------------------
col1, col2 = st.columns(2)
with col1:
    fecha_inicio = st.date_input("Fecha inicio", date.today(), format="DD/MM/YYYY")
with col2:
    fecha_fin = st.date_input("Fecha fin", date.today(), format="DD/MM/YYYY")

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

            df_bd2_final = pd.DataFrame({
                "Empresa": "CRLV",
                "Número": df_bd2["name"],
                "Fecha": df_bd2["invoice_date"],
                "Nro. Factura": df_bd2["invoice_number_next"],
                "Cliente": df_bd2["partner_id"].apply(lambda x: x[1] if x else ""),
                "Exento": exento,
                "Total Gravado": total_gravado,
                "Impuesto": impuesto,
                "Total": exento + total_gravado + (impuesto * 0.25),
                "Moneda": "Dolares"
            })

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
    
    st.subheader("📄 Previsualización")
    st.dataframe(df_final, width="stretch")

    st.subheader("📈 Resumen")
    colA, colB, colC = st.columns(3)
    colA.metric("Total Facturas", len(df_final))
    colB.metric("Total Impuesto", round(df_final["Impuesto"].sum(), 2))
    colC.metric("Total General", round(df_final["Total"].sum(), 2))

    # -----------------------------
    # FILTROS EXCLUSIONES
    # -----------------------------
    colA, colB = st.columns(2)

    with colA:
        excluir_nd = st.checkbox("Excluir todas las ND?", value=False)

    with colB:
        exclusiones_input = st.text_input(
            "Excluir ND específicas (separadas por coma)",
            value="",
            disabled=excluir_nd
        )

    # Preprocesar exclusiones específicas
    exclusiones_list = [x.strip() for x in exclusiones_input.split(",") if x.strip()]

    # -----------------------------
    # Aplicar filtros según checkbox/input
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