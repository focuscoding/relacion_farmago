import streamlit as st
import xmlrpc.client
import pandas as pd
import io
from datetime import date

# -----------------------------
# CONFIGURACIÓN GENERAL
# -----------------------------

st.set_page_config(page_title="Reporte Facturas BD1", layout="wide")
st.title("📊 Reporte de Facturas")

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
# FUNCIONES DE PROCESAMIENTO
# -----------------------------

def procesar_facturas(data):

    if not data:
        return pd.DataFrame()

    df = pd.DataFrame(data)

    # Extraer nombre moneda
    df["Moneda"] = df["currency_id"].apply(lambda x: x[1] if x else "")

    # Obtener impuesto según moneda
    def obtener_impuesto(row):
        if row["Moneda"] == "Dolares":
            return row.get("amount_tax_usd", 0) or 0
        else:
            return row.get("amount_tax_bs", 0) or 0

    df["Impuesto"] = df.apply(obtener_impuesto, axis=1)

    # Total Gravado = Impuesto / 16%
    df["Total Gravado"] = df["Impuesto"] / 0.16

    # Total = Exento + Total Gravado + 25% del impuesto
    df["Exento"] = df["iva_exempt"].fillna(0)

    df["Total"] = (
        df["Exento"]
        + df["Total Gravado"]
        + (df["Impuesto"] * 0.25)
    )

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

def generar_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output) as writer:
        df.to_excel(writer, index=False, sheet_name='Reporte')
    return output.getvalue()

# -----------------------------
# INTERFAZ
# -----------------------------

col1, col2 = st.columns(2)

with col1:
    fecha_inicio = st.date_input("Fecha inicio", date.today(), format="DD/MM/YYYY")

with col2:
    fecha_fin = st.date_input("Fecha fin", date.today(), format="DD/MM/YYYY")

if st.button("🔍 Consultar Facturas"):

    try:
        config = st.secrets["odoo_bd1"]

        client = OdooClient(
            config["url"],
            config["db"],
            config["username"],
            config["password"]
        )

        domain = [
            ('move_type', 'in', ['out_invoice', 'out_refund']),
            ('invoice_partner_display_name', '=', 'FARMACIA FARMAGO, C.A.'),
            ('invoice_date', '>=', str(fecha_inicio)),
            ('invoice_date', '<=', str(fecha_fin)),
            ('state', '=', 'posted')
        ]

        fields = [
            'name',
            'invoice_date',
            'invoice_number_next',
            'partner_id',
            'iva_exempt',
            'amount_tax_usd',
            'amount_tax_bs',
            'currency_id'
        ]

        with st.spinner("Consultando Odoo..."):
            data = client.search_read('account.move', domain, fields)

        df_final = procesar_facturas(data)

        if df_final.empty:
            st.warning("No se encontraron registros.")
        else:

            st.success(f"Se encontraron {len(df_final)} registros.")

            st.subheader("📄 Previsualización")
            st.dataframe(df_final, width="stretch")

            # Métricas resumen
            st.subheader("📈 Resumen")
            colA, colB, colC = st.columns(3)

            colA.metric("Total Facturas", len(df_final))
            colB.metric("Total Impuesto", round(df_final["Impuesto"].sum(), 2))
            colC.metric("Total General", round(df_final["Total"].sum(), 2))

            excel_file = generar_excel(df_final)

            st.download_button(
                label="⬇️ Descargar Excel",
                data=excel_file,
                file_name="reporte_facturas_bd1.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ocurrió un error: {str(e)}")

