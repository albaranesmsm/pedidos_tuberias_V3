import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from openpyxl import Workbook
from google.oauth2.service_account import Credentials
import gspread

# --- CONSTANTES FIJAS ---
COMPRADOR = "612539"
ADMIN_SHEET = "TUBQUIM_ADMIN"  # Nombre de tu Google Sheet con pestañas admin_user, disclaimer, materiales_bloqueados

# --- AUTENTICACIÓN GOOGLE SHEETS ---
SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=SCOPE
)
client = gspread.authorize(creds)

# --- CARGAR DATOS ADMIN ---
def cargar_admin_data():
    # Usuarios admin (username, password)
    try:
        ws_admin = client.open(ADMIN_SHEET).worksheet("admin_user")
        data = ws_admin.get_all_values()
        headers = data[0]
        rows = data[1:]
        usuarios = {row[0]: row[1] for row in rows if len(row) >= 2}
    except Exception as e:
        usuarios = {}
    # Disclaimer
    try:
        ws_disc = client.open(ADMIN_SHEET).worksheet("disclaimer")
        data = ws_disc.get_all_values()
        disclaimer = data[1][0] if len(data) > 1 else ""
    except Exception:
        disclaimer = ""
    # Materiales bloqueados
    try:
        ws_bloq = client.open(ADMIN_SHEET).worksheet("materiales_bloqueados")
        data = ws_bloq.get_all_values()
        bloqueados = [row[0] for row in data[1:] if row]
    except Exception:
        bloqueados = []
    return usuarios, disclaimer, bloqueados

# --- GUARDAR DISCLAIMER ---
def guardar_disclaimer(nuevo_texto):
    ws = client.open(ADMIN_SHEET).worksheet("disclaimer")
    ws.update("A2", nuevo_texto)

# --- GUARDAR MATERIALES BLOQUEADOS ---
def guardar_materiales_bloqueados(lista_referencias):
    ws = client.open(ADMIN_SHEET).worksheet("materiales_bloqueados")
    ws.resize(rows=1)  # dejamos solo fila encabezado
    if lista_referencias:
        cells = [[ref] for ref in lista_referencias]
        ws.append_rows(cells)

# --- DATOS FIJOS DEL PEDIDO ---
proveedores = {
    "1600043": "13161", "1600050": "13161", "1600051": "13161", "1600052": "13161",
    "1600053": "13161", "1600054": "13161", "1600055": "13161", "1600911": "13161",
    "1600921": "13161", "1601104": "13161", "1601161": "13161", "1601271": "13161",
    "1601306": "13161", "0400153": "10381", "0400176": "10381", "0400177": "10381",
    "0400232": "10381", "0400543": "10381", "0400548": "10381", "0400699": "10381",
    "1601001": "10381"
}

ob_values = {
    "1600043": "14001536", "1600050": "14001536", "1600051": "14001536", "1600052": "14001536",
    "1600053": "14001536", "1600054": "14001536", "1600055": "14001536", "1600911": "14001536",
    "1600921": "14001536", "1601104": "14001536", "1601161": "14001536", "1601271": "14001536",
    "1601306": "14001536", "0400153": "31005151", "0400176": "31005151", "0400177": "31005151",
    "0400232": "31005151", "0400543": "31005151", "0400548": "31005151", "0400699": "31005151",
    "1601001": "31005151"
}

articulos = [
    {"Nº artículo": "1600043", "Descripción": "TUB DESAG PVC//PVC"},
    {"Nº artículo": "1600050", "Descripción": "PYTHON A Inund 1 P"},
    {"Nº artículo": "1600051", "Descripción": "PYTHON A2 Inund 2 P"},
    {"Nº artículo": "1600052", "Descripción": "PYTHON L Contac 1P"},
    {"Nº artículo": "1600053", "Descripción": "PYTHON L2 Contac 2P"},
    {"Nº artículo": "1600054", "Descripción": "PYTHON L3 Contac 3P"},
    {"Nº artículo": "1600055", "Descripción": "TUB AGUA REF"},
    {"Nº artículo": "1600911", "Descripción": "PYTHON COL"},
    {"Nº artículo": "1600921", "Descripción": "LUPULUS 6,35 1P eventos"},
    {"Nº artículo": "1601104", "Descripción": "TUB ARM Riego arm"},
    {"Nº artículo": "1601161", "Descripción": "TUB GAS LDP"},
    {"Nº artículo": "1601271", "Descripción": "KIT ANTI-COND"},
    {"Nº artículo": "1601306", "Descripción": "FLEXLAYER 1P Vermut"},
    {"Nº artículo": "0400153", "Descripción": "DINFEX Antialgas"},
    {"Nº artículo": "0400176", "Descripción": "COMPACT 200 Limp inst."},
    {"Nº artículo": "0400177", "Descripción": "TOPFOAM Limp máq"},
    {"Nº artículo": "0400232", "Descripción": "PLUS ESPEC Limp inst."},
    {"Nº artículo": "0400543", "Descripción": "ALUTRAT Limp inst."},
    {"Nº artículo": "0400548", "Descripción": "ULTRASON Liquido"},
    {"Nº artículo": "0400699", "Descripción": "DIVOSAN TC86 SDC"},
    {"Nº artículo": "1601001", "Descripción": "GLICOL Anticong"}
]

restricciones = {
    "1600043": {"multiplo": 50, "max": 1500}, "1600050": {"multiplo": 25, "max": 2000},
    "1600051": {"multiplo": 25, "max": 500}, "1600052": {"multiplo": 25, "max": 500},
    "1600053": {"multiplo": 25, "max": 500}, "1600054": {"multiplo": 25, "max": 500},
    "1600055": {"multiplo": 50, "max": 1500}, "1600911": {"multiplo": 50, "max": 1000},
    "1600921": {"multiplo": 25, "max": 6000}, "1601104": {"multiplo": 25, "max": 50},
    "1601161": {"multiplo": 25, "max": 5000}, "1601271": {"multiplo": 25, "max": 300},
    "1601306": {"multiplo": 25, "max": 300}, "0400153": {"multiplo": 10, "max": 300},
    "0400176": {"multiplo": 10, "max": 80}, "0400177": {"multiplo": 20, "max": 80},
    "0400232": {"multiplo": 600, "max": 1800}, "0400543": {"multiplo": 20, "max": 300},
    "0400548": {"multiplo": 20, "max": 50}, "0400699": {"multiplo": 25, "max": 250},
    "1601001": {"multiplo": 25, "max": 600}
}

# --- ESTADO DE SESIÓN ---
if "mostrar_instrucciones" not in st.session_state:
    st.session_state.mostrar_instrucciones = False
if "admin_logged_in" not in st.session_state:
    st.session_state.admin_logged_in = False

# --- CARGAR DATOS DE ADMIN ---
usuarios_admin, disclaimer_text, materiales_bloqueados = cargar_admin_data()

# --- LOGIN ADMIN ---
def login_admin():
    st.sidebar.subheader("🔐 Login Admin")
    usuario = st.sidebar.text_input("Usuario")
    password = st.sidebar.text_input("Contraseña", type="password")
    if st.sidebar.button("Entrar"):
        if usuario in usuarios_admin and usuarios_admin[usuario] == password:
            st.session_state.admin_logged_in = True
            st.sidebar.success("Has iniciado sesión como admin.")
        else:
            st.sidebar.error("Usuario o contraseña incorrectos.")

def logout_admin():
    if st.sidebar.button("Cerrar sesión"):
        st.session_state.admin_logged_in = False
        st.experimental_rerun()

if not st.session_state.admin_logged_in:
    login_admin()
else:
    logout_admin()

# --- MOSTRAR DISCLAIMER ---
if disclaimer_text:
    st.markdown("### ℹ️ Información importante")
    st.info(disclaimer_text)
    st.divider()

# --- TÍTULO ---
st.title("Pedido de tuberías y químicos")



st.divider()

# --- SI ESTÁ LOGUEADO ADMIN, MOSTRAR ADMIN PANEL ---
if st.session_state.admin_logged_in:
    st.sidebar.header("⚙️ Panel Admin")

    st.sidebar.subheader("Editar Disclaimer")
    nuevo_disclaimer = st.sidebar.text_area("Texto del disclaimer:", value=disclaimer_text, height=150)
    if st.sidebar.button("Guardar Disclaimer"):
        guardar_disclaimer(nuevo_disclaimer)
        st.sidebar.success("Disclaimer actualizado. Recarga la app.")
    
    st.sidebar.subheader("Gestionar materiales bloqueados")
    lista_bloqueados = st.sidebar.text_area(
        "Introduce referencias bloqueadas, una por línea:",
        value="\n".join(materiales_bloqueados),
        height=150
    )
    if st.sidebar.button("Guardar bloqueados"):
        nuevas_refs = [line.strip() for line in lista_bloqueados.split("\n") if line.strip()]
        guardar_materiales_bloqueados(nuevas_refs)
        st.sidebar.success("Materiales bloqueados actualizados. Recarga la app.")

# --- DIRECCIONES ---
direcciones = st.secrets["direcciones"]

# --- VALIDAR DIRECCIÓN DE ENTREGA ---
dir_entrega = st.text_input("Código de Dirección de Entrega (4 cifras empezando por 8):", max_chars=4)

if not dir_entrega:
    st.warning("Debe introducir el código de almacén de envío.")
    st.stop()
elif not (dir_entrega.isdigit() and len(dir_entrega) == 4 and dir_entrega.startswith("8")):
    st.error("El código debe tener 4 cifras y empezar por 8.")
    st.stop()
elif dir_entrega not in direcciones:
    st.error("⚠️ Almacén no reconocido. Contacte con OOVV antes de hacer el pedido.")
    st.stop()
else:
    st.success(f"Dirección seleccionada: {direcciones[dir_entrega]}")

# --- SECCIÓN DE PEDIDO ---
st.subheader("Selecciona las cantidades:")

pedido = []
errores_multiplo = []

for articulo in articulos:
    codigo = articulo["Nº artículo"]
    if codigo in materiales_bloqueados:
        continue
    descripcion = articulo["Descripción"]
    proveedor = proveedores.get(codigo)
    ob = ob_values.get(codigo)
    if not proveedor or not ob:
        continue
    restric = restricciones.get(codigo, {})
    maximo = restric.get("max", 1000)
    multiplo = restric.get("multiplo", 1)
    cantidad = st.number_input(
        f"{descripcion} — Ref: {codigo} (Múltiplo: {multiplo}, Máx: {maximo})",
        min_value=0, max_value=maximo, step=multiplo, value=0,
    )
    if cantidad > 0:
        if cantidad % multiplo != 0:
            errores_multiplo.append(f"{codigo} ({descripcion}) — debe ser múltiplo de {multiplo}")
        pedido.append({
            "Fecha solicitud": datetime.date.today(),
            "OB": ob,
            "Comprador": COMPRADOR,
            "LM aux": "00004014",
            "Cód Prov": proveedor,
            "Proveedor": "",
            "Suc/planta": 8040,
            "Dir entr": dir_entrega,
            "Nº artículo": codigo,
            "Descripción": descripcion,
            "Autorizar cant": cantidad,
        })

# --- CREAR EXCEL ---
def crear_excel_protegido(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Pedido"
    ws.append(df.columns.tolist())
    for _, row in df.iterrows():
        ws.append(row.tolist())
    ws.protection.sheet = True
    ws.protection.set_password("NESTARES_24")
    excel_stream = BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0)
    return excel_stream

# --- GENERAR PEDIDO ---
if st.button("Generar Pedido"):
    if not pedido:
        st.warning("No se ha seleccionado ningún artículo.")
    elif errores_multiplo:
        st.error("⚠️ Las siguientes líneas no cumplen el múltiplo requerido:\n\n" + "\n".join(errores_multiplo))
    else:
        df = pd.DataFrame(pedido)
        excel_bytes = crear_excel_protegido(df)
        nombre_archivo = f"TubQuim_{dir_entrega}_{datetime.date.today()}.xlsx"
        st.download_button("Descargar Pedido", data=excel_bytes, file_name=nombre_archivo)
        st.success("¡Pedido generado correctamente!")
        st.session_state.mostrar_instrucciones = True

# --- INSTRUCCIONES ---
if st.session_state.mostrar_instrucciones:
    st.markdown(f"""
<div style='border: 2px solid #4CAF50; padding: 20px; border-radius: 10px; background-color: #f6fff6'>
<h4>¡Enhorabuena! Ya has generado tu pedido</h4>
<p>Ahora tienes que enviar el fichero generado por mail siguiendo estas instrucciones:</p>
<ol>
<li>Haz clic en <b>Descargar Pedido</b></li>
<li><a href='mailto:robot1@mahou-sanmiguel.com?subject=OAs%20pedidos%20materiales%20operaciones%20de%20venta&body=Adjunta%20aquí%20tu%20pedido' target='_blank'>📧 Enviar correo automáticamente</a></li>
<li>Adjunta el fichero descargado y pulsa <b>Enviar</b>.</li>
</ol>
<p><b>IMPORTANTE:</b> Usa exactamente el destinatario y el asunto indicados para que se generen correctamente las OAs.</p>
</div>

    """, unsafe_allow_html=True)
