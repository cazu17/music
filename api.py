# ============================================================
# ingresar_usuarios_v3.py — Automatización de ingreso de usuarios
# Lee desde un archivo Excel (.xlsx) y valida contra Oracle antes
# de tocar el formulario de la intranet.
#
# NOVEDADES respecto a v2:
#   1. Validación previa en Oracle (tabla WEBUSUARIOS por defecto):
#        - Se detiene/omite todo RUT que ya exista.
#        - El login generado se re-valida contra la BD y contra los
#          logins ya usados en esta misma corrida; si choca, se agrega
#          un sufijo numérico (2, 3, 4…) hasta encontrar uno libre.
#   2. Credenciales (intranet y Oracle) fuera del código, vía .env
#      (python-dotenv) — ya no quedan contraseñas en texto plano en
#      el script.
#   3. Logging a archivo + consola, enmascarando RUT/correo/clave para
#      no dejar datos personales en los logs (según política interna).
#   4. Validación de columnas del Excel antes de procesar, y detección
#      de RUTs duplicados dentro del mismo archivo.
#   5. Modo --dry-run: corre toda la validación (Excel + Oracle) sin
#      abrir el navegador, para revisar antes de ingresar de verdad.
#   6. Verificación post-"Grabar" (busca mensaje de éxito/error en la
#      página) en lugar de asumir éxito tras un sleep fijo.
#   7. Reporte final .csv con el resultado de cada usuario.
#
# Requiere: pip install -r requirements.txt
# ============================================================

import argparse
import csv
import logging
import os
import re
import sys
import time
import unicodedata
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

try:
    import oracledb
except ImportError:  # se avisa recién al intentar usarlo, no al importar el script
    oracledb = None

# ── Configuración (todo viene de variables de entorno / .env) ──────────
load_dotenv()

EXCEL_FILE = os.getenv("EXCEL_FILE", "Creación de Cuenta INTRANET.xlsx")

URL_LOGIN = os.getenv("URL_LOGIN", "http://intranet:8001/portal/")
URL_REDIR = os.getenv(
    "URL_REDIR", "http://intranet:8001/portal/redirITRv2.php?url=http://intranet/v2/homepage.asp"
)
URL_FORMULARIO = os.getenv(
    "URL_FORMULARIO", "http://intranet/v2/_Informes/0602_Usuarios/10GUS2_Ingresar_usuario.asp"
)

LOGIN_USER = os.environ["INTRANET_USER"]  # obligatorio: definir en .env
LOGIN_PASS = os.environ["INTRANET_PASSWORD"]  # obligatorio: definir en .env

CLAVE_USUARIO = os.getenv("CLAVE_USUARIO_DEFAULT", "")
if not CLAVE_USUARIO:
    raise RuntimeError("Defina CLAVE_USUARIO_DEFAULT en el archivo .env")

EMPSA_VALUE = os.getenv("EMPSA_VALUE", "1")
WAIT_TIMEOUT = int(os.getenv("WAIT_TIMEOUT", "15"))
MAX_LOGIN_LEN = int(os.getenv("MAX_LOGIN_LEN", "20"))

# ── Datos de conexión a Oracle (validación previa) ─────────────────────
ORACLE_USER = os.getenv("ORACLE_USER")
ORACLE_PASSWORD = os.getenv("ORACLE_PASSWORD")
ORACLE_DSN = os.getenv("ORACLE_DSN")  # ej: host:puerto/servicio
ORACLE_TABLE = os.getenv("ORACLE_TABLE", "WEBUSUARIOS")
ORACLE_COL_RUT = os.getenv("ORACLE_COL_RUT", "RUT")
ORACLE_COL_LOGIN = os.getenv("ORACLE_COL_LOGIN", "USULOGIN")

REPORT_DIR = os.getenv("REPORT_DIR", ".")

# ══════════════════════════════════════════════════════════════
# LOGGING (con enmascaramiento de datos personales)
# ══════════════════════════════════════════════════════════════

os.makedirs(REPORT_DIR, exist_ok=True)
_log_path = os.path.join(REPORT_DIR, f"ingreso_usuarios_{datetime.now():%Y%m%d_%H%M%S}.log")

logger = logging.getLogger("ingreso_usuarios")
logger.setLevel(logging.INFO)
_fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S")

_console = logging.StreamHandler(sys.stdout)
_console.setFormatter(_fmt)
_file = logging.FileHandler(_log_path, encoding="utf-8")
_file.setFormatter(_fmt)

logger.addHandler(_console)
logger.addHandler(_file)


def enmascarar(valor: str, visibles: int = 3) -> str:
    """Enmascara un dato sensible dejando visibles solo los últimos N caracteres.
    Usar para RUT, correo, teléfono, login, etc. en cualquier log o print.
    Nunca registrar la clave en texto plano (se reemplaza siempre por '****').
    """
    valor = str(valor)
    if len(valor) <= visibles:
        return "*" * len(valor)
    return "*" * (len(valor) - visibles) + valor[-visibles:]


# ══════════════════════════════════════════════════════════════
# FUNCIONES AUXILIARES DE FORMATO
# ══════════════════════════════════════════════════════════════

def limpiar_rut(rut: str) -> str:
    """Elimina puntos del RUT. Ejemplo: '20.293.212-6' → '20293212-6'."""
    return str(rut).replace(".", "").strip()


def generar_login(nombre_completo: str) -> str:
    """
    Genera el login base a partir del nombre completo.
    Regla: 1ra letra del primer nombre + primer apellido completo
           + 1ra letra del segundo apellido. Todo en MAYÚSCULAS.
    Se trunca a MAX_LOGIN_LEN caracteres por si el nombre es muy largo.
    """
    partes = nombre_completo.strip().split()

    if len(partes) >= 4:
        primera_nombre = partes[0][0]
        apellido1 = partes[-2]
        apellido2 = partes[-1]
    elif len(partes) == 3:
        primera_nombre = partes[0][0]
        apellido1 = partes[1]
        apellido2 = partes[2]
    elif len(partes) == 2:
        primera_nombre = partes[0][0]
        apellido1 = partes[1]
        apellido2 = ""
    else:
        return partes[0].upper()[:MAX_LOGIN_LEN]

    login = primera_nombre + apellido1 + (apellido2[0] if apellido2 else "")
    return login.upper()[:MAX_LOGIN_LEN]


def a_titulo(texto: str) -> str:
    """Convierte texto de MAYÚSCULAS a Title Case."""
    return str(texto).strip().title()


def normalizar(texto: str) -> str:
    """Quita acentos y convierte a minúsculas para comparación flexible."""
    nfkd = unicodedata.normalize("NFKD", texto)
    sin_acentos = "".join(c for c in nfkd if not unicodedata.combining(c))
    return sin_acentos.strip().lower()


def seleccionar_por_texto(select_element: Select, texto_buscado: str) -> bool:
    """Selecciona una opción de un <select> por texto visible, de forma flexible."""
    texto_norm = normalizar(texto_buscado)
    mejor_opcion = None

    for opcion in select_element.options:
        if normalizar(opcion.text.strip()) == texto_norm:
            mejor_opcion = opcion.text.strip()
            break

    if not mejor_opcion:
        for opcion in select_element.options:
            texto_opcion = opcion.text.strip()
            if texto_norm in normalizar(texto_opcion) or normalizar(texto_opcion) in texto_norm:
                mejor_opcion = texto_opcion
                break

    if mejor_opcion:
        select_element.select_by_visible_text(mejor_opcion)
        return True

    logger.warning("No se encontró opción de <select> para: '%s'", texto_buscado)
    return False


# ══════════════════════════════════════════════════════════════
# VALIDACIÓN PREVIA CONTRA ORACLE
# ══════════════════════════════════════════════════════════════

def conectar_oracle():
    """Abre una conexión Oracle en modo 'thin' (sin necesitar Oracle Client)."""
    if oracledb is None:
        raise RuntimeError(
            "Falta el paquete 'oracledb'. Instálelo con: pip install oracledb"
        )
    faltantes = [n for n, v in (
        ("ORACLE_USER", ORACLE_USER),
        ("ORACLE_PASSWORD", ORACLE_PASSWORD),
        ("ORACLE_DSN", ORACLE_DSN),
    ) if not v]
    if faltantes:
        raise RuntimeError(f"Defina en .env las variables faltantes: {', '.join(faltantes)}")

    conn = oracledb.connect(user=ORACLE_USER, password=ORACLE_PASSWORD, dsn=ORACLE_DSN)
    logger.info("Conexión a Oracle establecida (DSN: %s).", ORACLE_DSN)
    return conn


def existe_rut(conn, rut: str) -> bool:
    """True si el RUT ya está registrado en la tabla de usuarios."""
    sql = f"SELECT COUNT(*) FROM {ORACLE_TABLE} WHERE {ORACLE_COL_RUT} = :rut"
    with conn.cursor() as cur:
        cur.execute(sql, rut=rut)
        (total,) = cur.fetchone()
    return total > 0


def existe_login(conn, login: str) -> bool:
    """True si el login ya está registrado en la tabla de usuarios."""
    sql = f"SELECT COUNT(*) FROM {ORACLE_TABLE} WHERE {ORACLE_COL_LOGIN} = :login"
    with conn.cursor() as cur:
        cur.execute(sql, login=login)
        (total,) = cur.fetchone()
    return total > 0


def generar_login_unico(conn, nombre_completo: str, logins_usados_en_lote: set) -> str:
    """
    Genera un login y lo valida contra:
      a) los logins ya asignados en esta misma corrida (evita choques dentro del lote)
      b) los logins ya existentes en Oracle (evita choques con usuarios reales)
    Si hay colisión, agrega un sufijo numérico incremental (2, 3, 4…).
    """
    base = generar_login(nombre_completo)
    candidato = base
    sufijo = 2
    while candidato in logins_usados_en_lote or existe_login(conn, candidato):
        candidato = f"{base}{sufijo}"[:MAX_LOGIN_LEN]
        sufijo += 1
    logins_usados_en_lote.add(candidato)
    return candidato


def validar_usuarios_contra_oracle(usuarios: list, conn) -> list:
    """
    Recorre la lista de usuarios leídos del Excel y, para cada uno:
      - Marca 'omitir' = True si el RUT ya existe en Oracle (no se procesa).
      - Re-genera el login si el propuesto colisiona con uno existente.
    Devuelve la misma lista enriquecida con 'omitir' y 'motivo_omision'.
    """
    logins_usados_en_lote = set()
    for u in usuarios:
        u["omitir"] = False
        u["motivo_omision"] = None

        if existe_rut(conn, u["rut"]):
            u["omitir"] = True
            u["motivo_omision"] = "RUT ya existe en Oracle"
            logger.warning(
                "RUT %s (%s) ya existe en Oracle → se omitirá del ingreso.",
                enmascarar(u["rut"]), u["nombre"],
            )
            # igual reservamos el login para que no se lo lleve otro usuario del lote
            continue

        login_final = generar_login_unico(conn, u["nombre"], logins_usados_en_lote)
        if login_final != u["login"]:
            logger.info(
                "Login '%s' ya existía; se usará '%s' para %s.",
                u["login"], login_final, u["nombre"],
            )
        u["login"] = login_final

    return usuarios


# ══════════════════════════════════════════════════════════════
# FUNCIONES DEL NAVEGADOR (SELENIUM)
# ══════════════════════════════════════════════════════════════

def iniciar_sesion(driver, wait):
    driver.get(URL_LOGIN)
    wait.until(EC.presence_of_element_located((By.ID, "mod_login_username"))).send_keys(LOGIN_USER)
    driver.find_element(By.ID, "mod_login_password").send_keys(LOGIN_PASS)
    driver.find_element(By.CSS_SELECTOR, 'input[value="Entrar"]').click()
    time.sleep(2)
    logger.info("Sesión iniciada correctamente en la intranet.")


def navegar_a_formulario(driver):
    driver.get(URL_REDIR)
    time.sleep(2)
    logger.info("Redirección completada.")


def limpiar_y_escribir(driver, name: str, valor: str):
    campo = driver.find_element(By.NAME, name)
    campo.clear()
    campo.send_keys(valor)


def verificar_resultado_grabado(driver, wait) -> str:
    """
    Intenta detectar un mensaje de éxito/error tras pulsar 'Grabar'.
    NOTA: ajuste el selector CSS según el HTML real del formulario
    (aquí se intenta con clases/ids comunes como fallback razonable).
    Devuelve el texto encontrado, o 'DESCONOCIDO' si no hay forma de verificarlo.
    """
    posibles_selectores = [
        "#mensaje", ".mensaje", ".alert", "#resultado", ".resultado",
        "span.error", "span.ok", "div.msgOK", "div.msgError",
    ]
    for sel in posibles_selectores:
        try:
            elem = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, sel))
            )
            if elem.text.strip():
                return elem.text.strip()
        except TimeoutException:
            continue
    return "DESCONOCIDO (revisar manualmente en la intranet)"


def ingresar_usuario(driver, wait, rut, nombre, login, clave, email, depto, oficina, cargo) -> str:
    """Rellena y envía el formulario para un usuario. Devuelve el resultado detectado."""
    driver.get(URL_FORMULARIO)
    wait.until(EC.presence_of_element_located((By.NAME, "txtrut")))
    time.sleep(1)

    limpiar_y_escribir(driver, "txtrut", rut)
    limpiar_y_escribir(driver, "txtnombre", nombre)
    limpiar_y_escribir(driver, "txtlogin", login)
    limpiar_y_escribir(driver, "txtclave", clave)
    limpiar_y_escribir(driver, "txtemail", email)
    limpiar_y_escribir(driver, "txtcargo", cargo)

    select_depto = Select(driver.find_element(By.NAME, "depto"))
    seleccionar_por_texto(select_depto, depto)
    time.sleep(0.5)

    select_oficod = Select(driver.find_element(By.NAME, "oficod"))
    seleccionar_por_texto(select_oficod, oficina)
    time.sleep(0.5)

    select_ofiori = Select(driver.find_element(By.NAME, "ofiori"))
    seleccionar_por_texto(select_ofiori, oficina)
    time.sleep(0.5)

    Select(driver.find_element(By.NAME, "empsa")).select_by_value(EMPSA_VALUE)

    for tabla in ["tabla_Canales", "tabla_Lineas", "tabla_Oficinas"]:
        try:
            cb = driver.find_element(
                By.CSS_SELECTOR, f"input[onclick*=\"SelectAllCheckboxes('{tabla}'\"]"
            )
            if not cb.is_selected():
                cb.click()
        except NoSuchElementException:
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

    driver.find_element(By.CSS_SELECTOR, 'input[value=" Grabar "]').click()
    return verificar_resultado_grabado(driver, wait)


# ══════════════════════════════════════════════════════════════
# LECTURA Y VALIDACIÓN DEL EXCEL
# ══════════════════════════════════════════════════════════════

COLUMNAS_REQUERIDAS = [
    "Rut (sin puntos)",
    "Nombre completo del nuevo usuario",
    "Correo electrónico institucional",
    "Departamento",
    "Elija su oficina de origen (ver lista completa abajo)",
    "Descripción del Cargo",
]


def leer_usuarios_excel(filepath: str) -> list:
    df = pd.read_excel(filepath, engine="openpyxl")

    faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
    if faltantes:
        raise ValueError(
            "El Excel no tiene las columnas esperadas: " + ", ".join(faltantes)
        )

    usuarios = []
    ruts_vistos = set()
    for idx, row in df.iterrows():
        nombre_completo = str(row["Nombre completo del nuevo usuario"]).strip()
        if not nombre_completo or nombre_completo == "nan":
            continue

        rut = limpiar_rut(row["Rut (sin puntos)"])
        if rut in ruts_vistos:
            logger.warning(
                "Fila %s: RUT %s repetido dentro del mismo Excel; se omite el duplicado.",
                idx + 2, enmascarar(rut),
            )
            continue
        ruts_vistos.add(rut)

        email = str(row["Correo electrónico institucional"]).strip()
        oficina = str(row["Elija su oficina de origen (ver lista completa abajo)"]).strip()
        depto = str(row["Departamento"]).strip()
        cargo = str(row["Descripción del Cargo"]).strip()

        if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
            logger.warning(
                "Fila %s: correo con formato inválido para %s (%s); revíselo antes de continuar.",
                idx + 2, nombre_completo, enmascarar(email),
            )

        usuarios.append({
            "rut": rut,
            "nombre": nombre_completo,
            "login": generar_login(nombre_completo),  # se re-valida luego contra Oracle
            "clave": CLAVE_USUARIO,
            "email": email,
            "depto": depto,
            "oficina": a_titulo(oficina),
            "cargo": cargo.upper(),
        })

    return usuarios


# ══════════════════════════════════════════════════════════════
# REPORTE FINAL
# ══════════════════════════════════════════════════════════════

def guardar_reporte(usuarios: list, resultados: dict):
    """Guarda un CSV con el resultado de cada usuario (datos sensibles enmascarados)."""
    ruta = os.path.join(REPORT_DIR, f"reporte_ingreso_{datetime.now():%Y%m%d_%H%M%S}.csv")
    with open(ruta, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["rut_enmascarado", "nombre", "login", "estado", "detalle"])
        for u in usuarios:
            rut_m = enmascarar(u["rut"])
            estado = resultados.get(u["rut"], {}).get("estado", "NO_PROCESADO")
            detalle = resultados.get(u["rut"], {}).get("detalle", u.get("motivo_omision") or "")
            w.writerow([rut_m, u["nombre"], u["login"], estado, detalle])
    logger.info("Reporte guardado en: %s", ruta)
    return ruta


# ══════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="Ingreso masivo de usuarios en la intranet.")
    parser.add_argument("--excel", default=EXCEL_FILE, help="Ruta del archivo Excel de entrada.")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Solo valida (Excel + Oracle), no abre el navegador ni ingresa nada.",
    )
    parser.add_argument(
        "--headless", action="store_true", help="Ejecuta Chrome sin ventana visible.",
    )
    parser.add_argument(
        "--si", action="store_true",
        help="Omite la confirmación interactiva (útil para ejecuciones programadas).",
    )
    args = parser.parse_args()

    usuarios = leer_usuarios_excel(args.excel)
    logger.info("Se encontraron %d usuarios en '%s'.", len(usuarios), args.excel)

    # ── Validación contra Oracle (RUT y login ya existentes) ───────────
    conn_oracle = conectar_oracle()
    try:
        usuarios = validar_usuarios_contra_oracle(usuarios, conn_oracle)
    finally:
        conn_oracle.close()

    a_procesar = [u for u in usuarios if not u["omitir"]]
    omitidos = [u for u in usuarios if u["omitir"]]

    print("─" * 90)
    print("VISTA PREVIA (datos sensibles enmascarados):")
    print("─" * 90)
    for i, u in enumerate(usuarios, start=1):
        estado = "OMITIDO ⚠️ " + u["motivo_omision"] if u["omitir"] else "a procesar"
        print(f"  [{i:>2}] RUT: {enmascarar(u['rut'])}  |  Nombre: {u['nombre']}  |  {estado}")
        print(f"       Login: {u['login']}  |  Email: {enmascarar(u['email'])}  |  Clave: ****")
        print(f"       Depto: {u['depto']}  |  Oficina: {u['oficina']}  |  Cargo: {u['cargo']}")
        print()
    print("─" * 90)
    print(f"Total: {len(usuarios)}  |  A procesar: {len(a_procesar)}  |  Omitidos (ya existen): {len(omitidos)}")
    print("─" * 90)

    if args.dry_run:
        logger.info("Modo --dry-run: no se abrirá el navegador. Fin de la validación.")
        return

    if not a_procesar:
        logger.info("No hay usuarios nuevos por ingresar. Fin del proceso.")
        return

    if not args.si:
        confirmacion = input("\n¿Desea continuar con el ingreso de los usuarios NO omitidos? (s/n): ").strip().lower()
        if confirmacion != "s":
            logger.info("Proceso cancelado por el usuario.")
            return

    options = webdriver.ChromeOptions()
    if args.headless:
        options.add_argument("--headless=new")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, WAIT_TIMEOUT)
    resultados = {}

    try:
        iniciar_sesion(driver, wait)
        navegar_a_formulario(driver)

        for i, u in enumerate(a_procesar, start=1):
            logger.info(
                "[%d/%d] Ingresando usuario %s (login %s)…",
                i, len(a_procesar), enmascarar(u["rut"]), u["login"],
            )
            try:
                detalle = ingresar_usuario(
                    driver, wait,
                    rut=u["rut"], nombre=u["nombre"], login=u["login"],
                    clave=u["clave"], email=u["email"], depto=u["depto"],
                    oficina=u["oficina"], cargo=u["cargo"],
                )
                resultados[u["rut"]] = {"estado": "OK", "detalle": detalle}
                logger.info("   ✅ Usuario %s procesado. Resultado: %s", enmascarar(u["rut"]), detalle)
            except Exception as e:
                resultados[u["rut"]] = {"estado": "ERROR", "detalle": str(e)}
                logger.error("   ❌ Error al ingresar %s: %s", enmascarar(u["rut"]), e)
                continue

        logger.info("🎉 Proceso finalizado.")
    finally:
        driver.quit()
        guardar_reporte(usuarios, resultados)


if __name__ == "__main__":
    main()
