# ============================================================
# ingresar_usuarios_v2.py — Automatización de ingreso de usuarios
# Lee directamente desde un archivo Excel (.xlsx)
# Requiere: pip install selenium pandas openpyxl
# ============================================================

import time
import unicodedata
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

# ── Configuración ───────────────────────────────────────────
EXCEL_FILE     = "Creación de Cuenta INTRANET.xlsx"
URL_LOGIN      = "http://intranet:8001/portal/"
URL_REDIR      = "http://intranet:8001/portal/redirITRv2.php?url=http://intranet/v2/homepage.asp"
URL_FORMULARIO = "http://intranet/v2/_Informes/0602_Usuarios/10GUS2_Ingresar_usuario.asp"

# Credenciales de login
LOGIN_USER = "EVALENZUELAG"
LOGIN_PASS = "@PF.alimentos.2026"

# Valores fijos del formulario
CLAVE_USUARIO = "PF.1903!"
EMPSA_VALUE   = "1"
WAIT_TIMEOUT  = 15   # segundos máx. de espera por elemento


# ══════════════════════════════════════════════════════════════
# FUNCIONES AUXILIARES
# ══════════════════════════════════════════════════════════════

def limpiar_rut(rut: str) -> str:
    """
    Elimina puntos del RUT.
    Ejemplo: '20.293.212-6' → '20293212-6'
    """
    return str(rut).replace(".", "").strip()


def generar_login(nombre_completo: str) -> str:
    """
    Genera el login a partir del nombre completo.
    Regla: Primera letra del primer nombre + Primer apellido completo
           + Primera letra del segundo apellido. Todo en MAYÚSCULAS.

    Ejemplos:
      'Leonardo Franco Quiroz Chandia'         → 'LQUIROZC'
      'Juan Francisco Pozo Oyarzu'             → 'JPOZOO'
      'Camila Andrea Cubillos Muñoz'           → 'CCUBILLOSM'
      'IVER STIFF CASTILLO LUPAYANTE'          → 'ICASTILLOL'
      'Francisco Ignacio Jaramillo Hernández'  → 'FJARAMILLOH'
      'Sebastian Andres Flores Castillo'       → 'SFLORESC'
    """
    partes = nombre_completo.strip().split()

    if len(partes) >= 4:
        # Nombre1 [Nombre2...] Apellido1 Apellido2
        primera_nombre = partes[0][0]
        apellido1 = partes[-2]
        apellido2 = partes[-1]
    elif len(partes) == 3:
        # Nombre Apellido1 Apellido2
        primera_nombre = partes[0][0]
        apellido1 = partes[1]
        apellido2 = partes[2]
    elif len(partes) == 2:
        # Nombre Apellido (sin segundo apellido)
        primera_nombre = partes[0][0]
        apellido1 = partes[1]
        apellido2 = ""
    else:
        return partes[0].upper()

    login = primera_nombre + apellido1 + (apellido2[0] if apellido2 else "")
    return login.upper()


def a_titulo(texto: str) -> str:
    """
    Convierte texto de MAYÚSCULAS a Title Case.
    Ejemplo: 'PUERTO MONTT' → 'Puerto Montt'
    """
    return str(texto).strip().title()


def normalizar(texto: str) -> str:
    """Quita acentos y convierte a minúsculas para comparación flexible."""
    nfkd = unicodedata.normalize('NFKD', texto)
    sin_acentos = ''.join(c for c in nfkd if not unicodedata.combining(c))
    return sin_acentos.strip().lower()


def seleccionar_por_texto(select_element, texto_buscado: str) -> bool:
    """
    Selecciona una opción de un <select> buscando por texto visible
    de forma flexible (case-insensitive, sin acentos, strip).
    Primero intenta coincidencia exacta normalizada, luego parcial.
    """
    texto_norm = normalizar(texto_buscado)
    mejor_opcion = None

    # Intento 1: coincidencia exacta normalizada
    for opcion in select_element.options:
        texto_opcion = opcion.text.strip()
        if normalizar(texto_opcion) == texto_norm:
            mejor_opcion = texto_opcion
            break

    # Intento 2: coincidencia parcial (contiene)
    if not mejor_opcion:
        for opcion in select_element.options:
            texto_opcion = opcion.text.strip()
            if texto_norm in normalizar(texto_opcion) or normalizar(texto_opcion) in texto_norm:
                mejor_opcion = texto_opcion
                break

    if mejor_opcion:
        select_element.select_by_visible_text(mejor_opcion)
        return True
    else:
        print(f"   ⚠️  No se encontró opción para: '{texto_buscado}'")
        return False


# ══════════════════════════════════════════════════════════════
# FUNCIONES DEL NAVEGADOR (SELENIUM)
# ══════════════════════════════════════════════════════════════

def iniciar_sesion(driver, wait):
    """Inicia sesión en el portal de la intranet."""
    driver.get(URL_LOGIN)
    wait.until(EC.presence_of_element_located((By.ID, "mod_login_username"))).send_keys(LOGIN_USER)
    driver.find_element(By.ID, "mod_login_password").send_keys(LOGIN_PASS)
    driver.find_element(By.CSS_SELECTOR, 'input[value="Entrar"]').click()
    time.sleep(2)
    print("✅ Sesión iniciada correctamente.")


def navegar_a_formulario(driver):
    """Navega por la redirección hasta llegar al formulario de ingreso."""
    driver.get(URL_REDIR)
    time.sleep(2)
    print("✅ Redirección completada.")


def limpiar_y_escribir(driver, name: str, valor: str):
    """Limpia un input por su atributo name y escribe el valor."""
    campo = driver.find_element(By.NAME, name)
    campo.clear()
    campo.send_keys(valor)


def ingresar_usuario(driver, wait, rut, nombre, login, clave, email, depto, oficina, cargo):
    """
    Rellena y envía el formulario para un usuario.

    Parámetros ya procesados:
      rut     → sin puntos
      nombre  → nombre completo tal cual
      login   → generado automáticamente (MAYÚSCULAS)
      clave   → PF.1903!
      email   → correo institucional
      depto   → texto del departamento (se busca en el <select>)
      oficina → nombre de oficina en Title Case (se busca en el <select>)
      cargo   → descripción del cargo en MAYÚSCULAS
    """
    driver.get(URL_FORMULARIO)
    wait.until(EC.presence_of_element_located((By.NAME, "txtrut")))
    time.sleep(1)  # espera breve para que carguen los selects

    # ── Campos de texto ─────────────────────────────────────
    limpiar_y_escribir(driver, "txtrut",    rut)
    limpiar_y_escribir(driver, "txtnombre", nombre)
    limpiar_y_escribir(driver, "txtlogin",  login)
    limpiar_y_escribir(driver, "txtclave",  clave)
    limpiar_y_escribir(driver, "txtemail",  email)
    limpiar_y_escribir(driver, "txtcargo",  cargo)

    # ── Selects (desplegables) ──────────────────────────────
    # Departamento → buscar por texto visible (flexible)
    select_depto = Select(driver.find_element(By.NAME, "depto"))
    seleccionar_por_texto(select_depto, depto)
    time.sleep(0.5)

    # Oficina código → ya viene en Title Case
    select_oficod = Select(driver.find_element(By.NAME, "oficod"))
    seleccionar_por_texto(select_oficod, oficina)
    time.sleep(0.5)

    # Oficina origen → misma oficina en Title Case
    select_ofiori = Select(driver.find_element(By.NAME, "ofiori"))
    seleccionar_por_texto(select_ofiori, oficina)
    time.sleep(0.5)

    # Empresa → valor fijo "1"
    Select(driver.find_element(By.NAME, "empsa")).select_by_value(EMPSA_VALUE)

    # ── Checkboxes "Seleccionar Todo" ───────────────────────
    for tabla in ["tabla_Canales", "tabla_Lineas", "tabla_Oficinas"]:
        try:
            cb = driver.find_element(
                By.CSS_SELECTOR, f"input[onclick*=\"SelectAllCheckboxes('{tabla}'\"]"
            )
            if not cb.is_selected():
                cb.click()
        except Exception:
            # Fallback: ejecutar la función JS directamente
            driver.execute_script(f"""
                var cb = document.querySelector("input[onclick*=\\"SelectAllCheckboxes('{tabla}'\\")]");
                if (cb) {{
                    cb.checked = true;
                    SelectAllCheckboxes('{tabla}', cb);
                }} else {{
                    var checks = document.querySelectorAll("#{tabla} input[type='checkbox']");
                    checks.forEach(function(c) {{ c.checked = true; }});
                }}
            """)
    time.sleep(0.5)

    # ── Botón Grabar ────────────────────────────────────────
    btn_grabar = driver.find_element(By.CSS_SELECTOR, 'input[value=" Grabar "]')
    btn_grabar.click()
    time.sleep(3)  # espera a que se procese el formulario


# ══════════════════════════════════════════════════════════════
# LECTURA DEL EXCEL
# ══════════════════════════════════════════════════════════════

def leer_usuarios_excel(filepath: str) -> list:
    """
    Lee el archivo Excel y devuelve una lista de diccionarios
    con los campos necesarios ya procesados.

    Mapeo de columnas del Excel → campos del formulario:
      - 'Rut (sin puntos)'                                    → txtrut    (se quitan puntos)
      - 'Nombre completo del nuevo usuario'                   → txtnombre
      - (generado del nombre)                                 → txtlogin  (1ra letra nombre + apellido1 + 1ra letra apellido2)
      - 'PF.1903!'                                            → txtclave  (fijo)
      - 'Correo electrónico institucional'                    → txtemail
      - 'Departamento'                                        → depto     (select por texto)
      - 'Elija su oficina de origen (ver lista completa abajo)' → oficod/ofiori (Title Case)
      - 'Descripción del Cargo'                               → txtcargo  (MAYÚSCULAS)
    """
    df = pd.read_excel(filepath, engine="openpyxl")

    usuarios = []
    for _, row in df.iterrows():
        nombre_completo = str(row["Nombre completo del nuevo usuario"]).strip()
        if not nombre_completo or nombre_completo == "nan":
            continue

        rut_raw    = str(row["Rut (sin puntos)"]).strip()
        email      = str(row["Correo electrónico institucional"]).strip()
        oficina    = str(row["Elija su oficina de origen (ver lista completa abajo)"]).strip()
        depto      = str(row["Departamento"]).strip()
        cargo      = str(row["Descripción del Cargo"]).strip()

        usuarios.append({
            "rut":     limpiar_rut(rut_raw),
            "nombre":  nombre_completo,
            "login":   generar_login(nombre_completo),
            "clave":   CLAVE_USUARIO,
            "email":   email,
            "depto":   depto,
            "oficina": a_titulo(oficina),   # Title Case para el select
            "cargo":   cargo.upper(),        # Todo en MAYÚSCULAS
        })

    return usuarios


# ══════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════

def main():
    # ── Leer Excel ──────────────────────────────────────────
    usuarios = leer_usuarios_excel(EXCEL_FILE)
    print(f"📋 Se encontraron {len(usuarios)} usuarios en '{EXCEL_FILE}'.\n")

    # ── Vista previa de los datos procesados ────────────────
    print("─" * 90)
    print("VISTA PREVIA DE DATOS A INGRESAR:")
    print("─" * 90)
    for i, u in enumerate(usuarios, start=1):
        print(f"  [{i:>2}] RUT: {u['rut']}  |  Nombre: {u['nombre']}")
        print(f"       Login: {u['login']}  |  Email: {u['email']}  |  Clave: {u['clave']}")
        print(f"       Depto: {u['depto']}  |  Oficina: {u['oficina']}  |  Cargo: {u['cargo']}")
        print()
    print("─" * 90)

    confirmacion = input("\n¿Desea continuar con el ingreso? (s/n): ").strip().lower()
    if confirmacion != "s":
        print("❌ Proceso cancelado por el usuario.")
        return

    # ── Iniciar navegador ───────────────────────────────────
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")  # ← Descomenta para ejecutar sin ventana
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, WAIT_TIMEOUT)

    try:
        iniciar_sesion(driver, wait)
        navegar_a_formulario(driver)

        for i, u in enumerate(usuarios, start=1):
            print(f"[{i}/{len(usuarios)}] Ingresando: {u['rut']} — {u['nombre']} ({u['login']})…")
            try:
                ingresar_usuario(
                    driver, wait,
                    rut=u["rut"],
                    nombre=u["nombre"],
                    login=u["login"],
                    clave=u["clave"],
                    email=u["email"],
                    depto=u["depto"],
                    oficina=u["oficina"],
                    cargo=u["cargo"],
                )
                print(f"   ✅ Usuario {u['rut']} ingresado correctamente.\n")
            except Exception as e:
                print(f"   ❌ Error al ingresar {u['rut']}: {e}\n")
                continue

        print("🎉 Proceso finalizado.")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
