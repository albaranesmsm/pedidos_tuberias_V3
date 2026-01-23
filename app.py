import streamlit as st
import sqlite3
import pandas as pd
st.set_page_config(page_title="PETICIÓN DE MATERIALES", layout="wide")
# --- DB ---
conn = sqlite3.connect("config.db", check_same_thread=False)
cur = conn.cursor()
# --- TABLAS ---
cur.execute("""
CREATE TABLE IF NOT EXISTS materiales (
   codigo TEXT PRIMARY KEY,
   descripcion TEXT,
   proveedor TEXT,
   ob_value TEXT,
   multiplo INTEGER,
   maximo INTEGER,
   bloqueado INTEGER DEFAULT 0
)
""")
cur.execute("""
CREATE TABLE IF NOT EXISTS materiales_almacenes (
   codigo_material TEXT,
   almacen TEXT
)
""")
cur.execute("""
CREATE TABLE IF NOT EXISTS direcciones (
   codigo TEXT PRIMARY KEY,
   descripcion TEXT
)
""")
cur.execute("""
CREATE TABLE IF NOT EXISTS disclaimers (
   id INTEGER PRIMARY KEY AUTOINCREMENT,
   texto TEXT
)
""")
conn.commit()
# --- LOGIN ADMIN ---
def login_admin():
   st.subheader("Login administrador")
   u = st.text_input("Usuario")
   p = st.text_input("Contraseña", type="password")
   if st.button("Entrar"):
       if u == st.secrets["admin"]["user"] and p == st.secrets["admin"]["password"]:
           st.session_state.admin = True
           st.success("Acceso correcto")
       else:
           st.error("Credenciales incorrectas")
# --- ADMIN PANEL ---
def admin_panel():
   st.title("Panel de Administración")
   # -- Sección Materiales --
   st.markdown("### Materiales")
   with st.form("material_form"):
       codigo = st.text_input("Código material")
       descripcion = st.text_input("Descripción")
       proveedor = st.text_input("Proveedor")
       ob = st.text_input("OB Value")
       multiplo = st.number_input("Múltiplo", min_value=1, step=1)
       maximo = st.number_input("Máximo", min_value=1, step=1)
       bloqueado = st.checkbox("Bloqueado")
       # Lista almacenes desde BD
       cur.execute("SELECT codigo FROM direcciones")
       almacenes_list = [row[0] for row in cur.fetchall()]
       almacenes = st.multiselect("Almacenes permitidos", almacenes_list)
       guardar = st.form_submit_button("Guardar / Actualizar")
   if guardar:
       cur.execute("""
       INSERT OR REPLACE INTO materiales
       VALUES (?,?,?,?,?,?,?)
       """, (codigo, descripcion, proveedor, ob, multiplo, maximo, int(bloqueado)))
       cur.execute("DELETE FROM materiales_almacenes WHERE codigo_material = ?", (codigo,))
       for a in almacenes:
           cur.execute(
               "INSERT INTO materiales_almacenes VALUES (?,?)",
               (codigo, a)
           )
       conn.commit()
       st.success("Material guardado correctamente")
   st.divider()
   # -- Sección Direcciones --
   st.markdown("### Direcciones (Almacenes)")
   with st.form("direccion_form"):
       codigo_dir = st.text_input("Código almacén (4 dígitos)")
       descripcion_dir = st.text_input("Descripción almacén")
       guardar_dir = st.form_submit_button("Guardar/Actualizar dirección")
   if guardar_dir:
       if len(codigo_dir) == 4:
           cur.execute("""
           INSERT OR REPLACE INTO direcciones VALUES (?, ?)
           """, (codigo_dir, descripcion_dir))
           conn.commit()
           st.success("Dirección guardada")
       else:
           st.error("El código debe tener 4 dígitos")
   df_direcciones = pd.read_sql("SELECT * FROM direcciones", conn)
   st.dataframe(df_direcciones, use_container_width=True)
   st.divider()
   # -- Sección Disclaimer --
   st.markdown("### Disclaimer")
   cur.execute("SELECT * FROM disclaimers ORDER BY id DESC LIMIT 1")
   row = cur.fetchone()
   texto_actual = row[1] if row else ""
   with st.form("disclaimer_form"):
       texto = st.text_area("Texto del disclaimer", value=texto_actual, height=100)
       guardar_disc = st.form_submit_button("Guardar disclaimer")
   if guardar_disc:
       if texto.strip():
           cur.execute("INSERT INTO disclaimers (texto) VALUES (?)", (texto.strip(),))
           conn.commit()
           st.success("Disclaimer guardado")
   # Mostrar último disclaimer
   st.markdown("**Texto actual del disclaimer:**")
   st.info(texto_actual if texto_actual else "No hay texto guardado.")
   st.divider()
   # -- Mostrar tabla materiales --
   st.subheader("Materiales configurados")
   df_mat = pd.read_sql("SELECT * FROM materiales", conn)
   st.dataframe(df_mat, use_container_width=True)
# --- APP USUARIO ---
def app_usuario():
   st.title("PETICIÓN DE MATERIALES")
   # Elegir almacén de BD
   cur.execute("SELECT codigo, descripcion FROM direcciones")
   dirs = cur.fetchall()
   opciones_almacen = {c: d for c, d in dirs}
   almacen = st.selectbox(
       "Almacén",
       list(opciones_almacen.keys()),
       format_func=lambda x: f"{x} - {opciones_almacen[x]}" if x in opciones_almacen else x
   )
   # Mostrar disclaimer
   cur.execute("SELECT texto FROM disclaimers ORDER BY id DESC LIMIT 1")
   row = cur.fetchone()
   if row and row[0].strip():
       st.info(row[0])
   codigo = st.text_input("Código material")
   if not codigo:
       return
   cur.execute("""
   SELECT * FROM materiales
   WHERE codigo = ? AND bloqueado = 0
   """, (codigo,))
   mat = cur.fetchone()
   if not mat:
       st.error("Material inexistente o bloqueado")
       return
   cur.execute("""
   SELECT almacen FROM materiales_almacenes
   WHERE codigo_material = ?
   """, (codigo,))
   permitidos = [x[0] for x in cur.fetchall()]
   if almacen not in permitidos:
       st.error("Material no permitido para este almacén")
       return
   _, desc, prov, ob, multiplo, maximo, _ = mat
   st.success(desc)
   st.caption(f"Proveedor: {prov} | OB: {ob}")
   cantidad = st.number_input(
       "Cantidad",
       min_value=multiplo,
       step=multiplo
   )
   if cantidad > maximo:
       st.error(f"La cantidad máxima permitida es {maximo}")
   elif cantidad % multiplo != 0:
       st.error(f"La cantidad debe ser múltiplo de {multiplo}")
   else:
       st.info("Cantidad válida ✅")
# --- MAIN ---
if "admin" not in st.session_state:
   st.session_state.admin = False
menu = st.sidebar.radio("Menú", ["App", "Admin"])
if menu == "Admin":
   if not st.session_state.admin:
       login_admin()
   else:
       admin_panel()
else:
   app_usuario()
