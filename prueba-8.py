import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import time
import re
import unicodedata

try:
    import pandas as pd
except ImportError:
    pd = None

# ====================== INTENTO DE IMPORTAR OCR (opcional) ======================
OCR_AVAILABLE = False
try:
    from PIL import Image, ImageFilter, ImageEnhance, ImageOps
    from io import BytesIO
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = r'C:\Users\fserrano\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
    OCR_AVAILABLE = True
    print("✅ OCR (pytesseract) cargado correctamente")
except ImportError:
    print("⚠️  pytesseract NO está instalado → se usará modo MANUAL (CAPTCHA a mano)")
except Exception as e:
    print(f"⚠️  Error al cargar OCR: {e} → modo MANUAL")

load_dotenv()

URL_LOGIN = "https://www.sucamec.gob.pe/sel/faces/login.xhtml?faces-redirect=true"
EXCEL_PATH = os.getenv("EXCEL_PATH", os.path.join("data", "programaciones-armas.xlsx"))

CREDENCIALES = {
    "tipo_documento_valor": os.getenv("TIPO_DOC", "RUC"),
    "numero_documento": os.getenv("NUMERO_DOCUMENTO", ""),
    "usuario": os.getenv("USUARIO_SEL", ""),
    "contrasena": os.getenv("CLAVE_SEL", ""),
}

SEL = {
    "tab_tradicional": 'a[href="#tabViewLogin:j_idt33"]',
    "tipo_doc_select": "#tabViewLogin\\:tradicionalForm\\:tipoDoc_input",
    "numero_documento": "#tabViewLogin\\:tradicionalForm\\:documento",
    "usuario": "#tabViewLogin\\:tradicionalForm\\:usuario",
    "clave": "#tabViewLogin\\:tradicionalForm\\:clave",
    "captcha_img": "#tabViewLogin\\:tradicionalForm\\:imgCaptcha",
    "captcha_input": "#tabViewLogin\\:tradicionalForm\\:textoCaptcha",
    "boton_refresh": "#tabViewLogin\\:tradicionalForm\\:botonCaptcha",
    "ingresar": "#tabViewLogin\\:tradicionalForm\\:ingresar",

    # ── Menú PanelMenu PrimeFaces ─────────────────────────────────────────────
    # Header del acordeón CITAS  →  el <h3> que contiene el <a>CITAS</a>
    # Hacemos clic en él para expandir/colapsar el panel
    "menu_citas_header": '#j_idt11\\:menuPrincipal .ui-panelmenu-header:has(a:text-is("CITAS"))',

    # Panel de contenido que se despliega al hacer clic en el header CITAS
    # id fijo según el HTML: j_idt11:menuPrincipal_7
    "menu_citas_panel": '#j_idt11\\:menuPrincipal_7',

    # Ítem "RESERVAS DE CITAS" — usa el onclick con menuid='7_1'
    # Selector más robusto: busca dentro del panel CITAS el span con ese texto
    "submenu_reservas": '#j_idt11\\:menuPrincipal_7 span.ui-menuitem-text:text-is("RESERVAS DE CITAS")',

    # ── SelectOneMenu: tipo de cita en Gestión de Citas ──────────────────────
    "tipo_cita_trigger": '#gestionCitasForm\\:j_idt32 .ui-selectonemenu-trigger',
    "tipo_cita_panel": '#gestionCitasForm\\:j_idt32_panel',
    "tipo_cita_label": '#gestionCitasForm\\:j_idt32_label',
    "tipo_cita_opcion_poligono": '#gestionCitasForm\\:j_idt32_panel li[data-label="EXAMEN PARA POLÍGONO DE TIRO"]',

    # ── Reserva de Cupos (tabGestion:creaCitaPolJurForm) ───────────────────
    "reserva_form": '#tabGestion\\:creaCitaPolJurForm',
    "sede_trigger": '#tabGestion\\:creaCitaPolJurForm\\:sedeId .ui-selectonemenu-trigger',
    "sede_panel": '#tabGestion\\:creaCitaPolJurForm\\:sedeId_panel',
    "sede_label": '#tabGestion\\:creaCitaPolJurForm\\:sedeId_label',
    "fecha_trigger": '#tabGestion\\:creaCitaPolJurForm\\:listaDiasId .ui-selectonemenu-trigger',
    "fecha_panel": '#tabGestion\\:creaCitaPolJurForm\\:listaDiasId_panel',
    "fecha_label": '#tabGestion\\:creaCitaPolJurForm\\:listaDiasId_label',

    # ── Tabla de programación de cupos ──────────────────────────────────────
    "tabla_programacion": '#tabGestion\\:creaCitaPolJurForm\\:dtProgramacion',
    "tabla_programacion_rows": '#tabGestion\\:creaCitaPolJurForm\\:dtProgramacion_data tr',
    "boton_siguiente": '#tabGestion\\:creaCitaPolJurForm button:has-text("Siguiente")',

    # ── Paso 2 del Wizard ───────────────────────────────────────────────────
    "tipo_operacion_trigger": '#tabGestion\\:creaCitaPolJurForm\\:tipoOpe .ui-selectonemenu-trigger',
    "tipo_operacion_panel": '#tabGestion\\:creaCitaPolJurForm\\:tipoOpe_panel',
    "tipo_operacion_items": '#tabGestion\\:creaCitaPolJurForm\\:tipoOpe_panel li.ui-selectonemenu-item',
    "tipo_operacion_label": '#tabGestion\\:creaCitaPolJurForm\\:tipoOpe_label',
    "doc_vig_input": '#tabGestion\\:creaCitaPolJurForm\\:nroDocVig_input',
    "doc_vig_panel": '#tabGestion\\:creaCitaPolJurForm\\:nroDocVig_panel',
    "doc_vig_items": '#tabGestion\\:creaCitaPolJurForm\\:nroDocVig_panel li.ui-autocomplete-item',
    "seleccione_solicitud_trigger": '#tabGestion\\:creaCitaPolJurForm\\:seleccioneSolicitud .ui-selectonemenu-trigger',
    "seleccione_solicitud_panel": '#tabGestion\\:creaCitaPolJurForm\\:seleccioneSolicitud_panel',
    "seleccione_solicitud_si": '#tabGestion\\:creaCitaPolJurForm\\:seleccioneSolicitud_panel li[id$="_1"]',
    "seleccione_solicitud_label": '#tabGestion\\:creaCitaPolJurForm\\:seleccioneSolicitud_label',
    "nro_solicitud_trigger": '#tabGestion\\:creaCitaPolJurForm\\:nroSolicitud .ui-selectonemenu-trigger',
    "nro_solicitud_panel": '#tabGestion\\:creaCitaPolJurForm\\:nroSolicitud_panel',
    "nro_solicitud_items": '#tabGestion\\:creaCitaPolJurForm\\:nroSolicitud_panel li.ui-selectonemenu-item',
    "nro_solicitud_label": '#tabGestion\\:creaCitaPolJurForm\\:nroSolicitud_label',
}


# ============================================================
# OCR helpers  (sin cambios)
# ============================================================

def corregir_captcha_ocr(texto_raw: str) -> str:
    if not texto_raw:
        return ""
    texto = texto_raw.strip().upper().replace(" ", "").replace("\n", "").replace("\r", "")
    texto = ''.join(c for c in texto if c.isalnum())
    return texto


def validar_captcha_texto(texto: str) -> bool:
    if not texto or len(texto) != 5:
        return False
    return texto.isalnum()


def escribir_input_jsf(page, selector: str, valor: str):
    for intento in range(4):
        campo = page.locator(selector)
        campo.wait_for(state="visible", timeout=12000)

        # Intento principal: tipeo humano con delay alto para evitar pérdida de dígitos.
        campo.click()
        campo.press("Control+A")
        campo.press("Backspace")
        campo.type(valor, delay=65)
        campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
        page.wait_for_timeout(140)

        actual = campo.input_value().strip()
        if actual != valor:
            # Fallback fuerte: asignación directa del value y eventos JSF.
            campo.evaluate(
                '''(el, val) => {
                    el.focus();
                    el.value = val;
                    el.setAttribute("value", val);
                    el.dispatchEvent(new Event("input", { bubbles: true }));
                    el.dispatchEvent(new Event("change", { bubbles: true }));
                }''',
                valor
            )
            page.wait_for_timeout(120)
            actual = campo.input_value().strip()

        if actual == valor:
            # Dispara blur al final para el comportamiento JSF de validación.
            campo.evaluate('el => el.blur()')
            page.wait_for_timeout(220)
            try:
                confirmado = page.locator(selector).input_value().strip()
            except Exception:
                confirmado = ""
            if confirmado == valor:
                return
            actual = confirmado

        print(f"   ⚠️ Campo {selector}: esperado '{valor}', tiene '{actual}' → reintentando ({intento+1}/4)")
        page.wait_for_timeout(260)

    raise Exception(f"No se pudo fijar correctamente el valor del campo {selector}")


def escribir_input_rapido(page, selector: str, valor: str):
    campo = page.locator(selector)
    campo.wait_for(state="visible", timeout=10000)
    campo.click()
    campo.fill(valor)
    campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
    campo.blur()
    if campo.input_value() != valor:
        campo.click()
        campo.press("Control+A")
        campo.press("Backspace")
        campo.type(valor, delay=10)
        campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
        campo.blur()


def solve_captcha_manual(page):
    print("\n🔴 MODO MANUAL ACTIVADO")
    print("Llena el código de verificación en la ventana del navegador")
    input("✅ Cuando hayas escrito el CAPTCHA → presiona ENTER aquí para continuar...")


def preprocesar_imagen_captcha(img_bytes: bytes, variante: int = 0) -> 'Image':
    img = Image.open(BytesIO(img_bytes))
    img = img.convert('L')
    if variante == 0:
        img = ImageEnhance.Contrast(img).enhance(3.5)
        w, h = img.size
        img = img.resize((w * 4, h * 4), Image.LANCZOS)
        img = img.filter(ImageFilter.MedianFilter(size=3))
        img = ImageOps.invert(img)
        img = img.point(lambda p: 255 if p > 130 else 0)
        img = ImageEnhance.Sharpness(img).enhance(3.0)
    elif variante == 1:
        img = ImageEnhance.Contrast(img).enhance(2.5)
        w, h = img.size
        img = img.resize((w * 3, h * 3), Image.LANCZOS)
        img = img.filter(ImageFilter.MedianFilter(size=5))
        img = img.point(lambda p: 255 if p > 160 else 0)
        img = ImageEnhance.Sharpness(img).enhance(2.0)
    else:
        img = ImageEnhance.Contrast(img).enhance(4.0)
        w, h = img.size
        img = img.resize((w * 5, h * 5), Image.LANCZOS)
        img = img.filter(ImageFilter.GaussianBlur(radius=0.5))
        img = ImageOps.invert(img)
        img = img.point(lambda p: 255 if p > 110 else 0)
        img = ImageEnhance.Sharpness(img).enhance(4.0)
    return img


def solve_captcha_ocr(page):
    if not OCR_AVAILABLE:
        return None
    MAX_INTENTOS = 6
    PSM_MODES = [7, 8, 13]
    NUM_VARIANTES = 3
    for intento in range(MAX_INTENTOS):
        try:
            print(f"🔍 Intentando resolver CAPTCHA (intento {intento+1}/{MAX_INTENTOS})...")
            page.wait_for_timeout(200)
            img_bytes = page.locator(SEL["captcha_img"]).screenshot(type="png")
            mejor_texto = None
            for variante in range(NUM_VARIANTES):
                img = preprocesar_imagen_captcha(img_bytes, variante=variante)
                for psm in PSM_MODES:
                    config = f'--psm {psm} --oem 3 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ --dpi 300'
                    texto_raw = pytesseract.image_to_string(img, config=config, lang='eng').strip()
                    texto = corregir_captcha_ocr(texto_raw)
                    if validar_captcha_texto(texto):
                        print(f"   → Variante {variante}, PSM {psm}: '{texto_raw}' → '{texto}' ✓")
                        mejor_texto = texto
                        break
                    else:
                        print(f"   → Variante {variante}, PSM {psm}: '{texto_raw}' → '{texto}' (len={len(texto)}) ✗")
                if mejor_texto:
                    break
            if mejor_texto:
                print(f"   ✓ CAPTCHA válido → Usando: {mejor_texto}")
                return mejor_texto
            print("   ✗ Ninguna combinación dio resultado → Refrescando CAPTCHA...")
            print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
            page.locator(SEL["boton_refresh"]).click(force=True)
            page.wait_for_timeout(500)
        except Exception as e:
            print(f"   Error en intento {intento+1}: {str(e)}")
            page.wait_for_timeout(300)
    print(f"❌ No se pudo resolver automáticamente después de {MAX_INTENTOS} intentos → modo manual")
    return None


def normalizar_fecha_excel(valor_fecha: str) -> str:
    """Convierte fechas de Excel al formato dd/mm/yyyy esperado por SEL."""
    texto = str(valor_fecha or "").strip()
    if not texto:
        return ""

    # Caso típico de Excel: 2026-03-31 00:00:00
    dt = pd.to_datetime(texto, errors="coerce", dayfirst=False)
    if pd.notna(dt):
        return dt.strftime("%d/%m/%Y")

    # Si ya viene como dd/mm/yyyy o similar, lo conservamos sin hora
    texto = texto.split(" ")[0]
    return texto


def normalizar_hora_fragmento(valor_hora: str) -> str:
    """Normaliza una hora a HH:MM (ej: 8:5 -> 08:05)."""
    texto = str(valor_hora or "").strip().replace(".", ":")
    if ":" not in texto:
        return texto
    partes = texto.split(":")
    if len(partes) != 2:
        return texto
    try:
        hh = int(partes[0])
        mm = int(partes[1])
    except ValueError:
        return texto
    return f"{hh:02d}:{mm:02d}"


def normalizar_hora_rango(valor_rango: str) -> str:
    """Normaliza rango de hora a HH:MM-HH:MM."""
    texto = str(valor_rango or "").strip()
    if not texto:
        return ""
    texto = texto.replace("–", "-").replace("—", "-").replace(" a ", "-").replace(" ", "")
    partes = texto.split("-")
    if len(partes) != 2:
        return texto
    inicio = normalizar_hora_fragmento(partes[0])
    fin = normalizar_hora_fragmento(partes[1])
    return f"{inicio}-{fin}"


def convertir_a_entero(texto: str) -> int:
    numeros = re.findall(r"\d+", str(texto or ""))
    return int(numeros[0]) if numeros else 0


def normalizar_texto_comparable(texto: str) -> str:
    base = str(texto or "").strip().upper()
    base = unicodedata.normalize("NFKD", base)
    base = "".join(c for c in base if not unicodedata.combining(c))
    base = re.sub(r"\s+", " ", base)
    return base


def extraer_token_solicitud(valor: str) -> str:
    """Obtiene el número principal de solicitud para comparar dentro del label del combo."""
    texto = str(valor or "")
    grupos = re.findall(r"\d+", texto)
    if not grupos:
        return ""
    token = grupos[0].lstrip("0")
    return token if token else "0"


def normalizar_tipo_arma_excel(valor: str) -> str:
    """Normaliza valor de tipo_arma del Excel para comparaciones."""
    base = normalizar_texto_comparable(valor)
    equivalencias = {
        "LARG": "LARGA",
        "LARGA": "LARGA",
        "CORTA": "CORTA",
        "PISTOLA": "PISTOLA",
        "REVOLVER": "REVOLVER",
        "CARABINA": "CARABINA",
        "ESCOPETA": "ESCOPETA",
    }
    return equivalencias.get(base, base)


def inferir_objetivo_arma_desde_excel(valor: str) -> str:
    """
    Interpreta texto libre de tipo_arma y devuelve una clave usable.
    Ejemplos válidos: "CORTA", "CORTA PISTOLA", "LARGA ESCOPETA".
    """
    base = normalizar_texto_comparable(valor)
    if not base:
        return ""

    # Priorizamos el arma específica si está presente.
    if "ESCOPETA" in base:
        return "ESCOPETA"
    if "CARABINA" in base:
        return "CARABINA"
    if "REVOLVER" in base:
        return "REVOLVER"
    if "PISTOLA" in base:
        return "PISTOLA"

    # Si no hay arma específica, devolvemos tipo general.
    if "LARG" in base:
        return "LARGA"
    if "CORT" in base:
        return "CORTA"

    return normalizar_tipo_arma_excel(base)


def fecha_comparable(valor_fecha: str) -> str:
    """Convierte fecha de Excel a una cadena comparable dd/mm/yyyy."""
    return normalizar_fecha_excel(valor_fecha)


def cargar_primer_registro_pendiente_desde_excel(ruta_excel: str) -> dict:
    """
    Lee el Excel y devuelve el primer registro con estado 'Pendiente'.
    Campos mínimos requeridos para este paso: sede y fecha.
    """
    if pd is None:
        raise Exception("Falta dependencia 'pandas'. Instala con: pip install pandas openpyxl")

    if not os.path.exists(ruta_excel):
        raise Exception(f"No se encontró el Excel en: {ruta_excel}")

    df = pd.read_excel(ruta_excel, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    columnas_requeridas = {"sede", "fecha", "hora_rango", "tipo_operacion", "nro_solicitud", "tipo_arma", "arma", "estado"}
    faltantes = [c for c in columnas_requeridas if c not in df.columns]
    if faltantes:
        raise Exception(f"Faltan columnas requeridas en Excel: {faltantes}")

    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()

    pendientes = df[df["estado"].str.upper() == "PENDIENTE"]
    if pendientes.empty:
        raise Exception("No hay registros con estado 'Pendiente' en el Excel")

    indice_primer_pendiente = pendientes.index[0]
    registro = pendientes.loc[indice_primer_pendiente].to_dict()

    fecha_col_programacion = "fecha_programacion" if "fecha_programacion" in df.columns else "fecha"

    sede = registro.get("sede", "").strip()
    fecha = normalizar_fecha_excel(registro.get("fecha", ""))
    hora_rango = normalizar_hora_rango(registro.get("hora_rango", ""))
    tipo_operacion = registro.get("tipo_operacion", "").strip()
    nro_solicitud = registro.get("nro_solicitud", "").strip()
    doc_vigilante = registro.get("doc_vigilante", registro.get("dni", "")).strip()
    tipo_arma_base = inferir_objetivo_arma_desde_excel(registro.get("tipo_arma", ""))
    arma_base = inferir_objetivo_arma_desde_excel(registro.get("arma", ""))

    if not sede or not fecha or not hora_rango:
        raise Exception("El registro pendiente no tiene 'sede', 'fecha' o 'hora_rango' con valor")
    if not tipo_operacion or not nro_solicitud or not doc_vigilante:
        raise Exception("El registro pendiente no tiene 'tipo_operacion', 'doc_vigilante/dni' o 'nro_solicitud'")
    if not tipo_arma_base:
        raise Exception("El registro pendiente no tiene 'tipo_arma'")
    if not arma_base:
        raise Exception("El registro pendiente no tiene 'arma'")

    # Agrupa posibles duplicados de la misma programación (mismo usuario + misma fecha_programacion/fecha).
    fecha_base = fecha_comparable(registro.get(fecha_col_programacion, registro.get("fecha", "")))
    doc_base = doc_vigilante
    pendientes_aux = pendientes.copy()
    pendientes_aux["fecha_norm"] = pendientes_aux[fecha_col_programacion].apply(fecha_comparable)
    pendientes_aux["doc_norm"] = pendientes_aux.apply(
        lambda r: str(r.get("doc_vigilante", "") or r.get("dni", "")).strip(), axis=1
    )
    relacionados = pendientes_aux[
        (pendientes_aux["fecha_norm"] == fecha_base) &
        (pendientes_aux["doc_norm"] == doc_base)
    ]

    # Validación adicional solicitada: revisar explícitamente el siguiente registro.
    siguiente_mismo_doc_y_fecha = False
    siguiente_idx = indice_primer_pendiente + 1
    if siguiente_idx in df.index:
        fila_sig = df.loc[siguiente_idx]
        estado_sig = str(fila_sig.get("estado", "")).strip().upper()
        doc_sig = str(fila_sig.get("doc_vigilante", "") or fila_sig.get("dni", "")).strip()
        fecha_sig = fecha_comparable(fila_sig.get(fecha_col_programacion, fila_sig.get("fecha", "")))
        if estado_sig == "PENDIENTE" and doc_sig == doc_base and fecha_sig == fecha_base:
            siguiente_mismo_doc_y_fecha = True

    tipos_arma_excel = []
    armas_excel = []
    objetivos_arma = []
    armas_especificas = {"PISTOLA", "REVOLVER", "CARABINA", "ESCOPETA"}

    for _, fila in relacionados.iterrows():
        tipo_raw = str(fila.get("tipo_arma", "")).strip()
        arma_raw = str(fila.get("arma", "")).strip()
        tipo_inferido = inferir_objetivo_arma_desde_excel(tipo_raw)
        arma_inferida = inferir_objetivo_arma_desde_excel(arma_raw)

        if not arma_inferida:
            arma_inferida = inferir_objetivo_arma_desde_excel(tipo_raw)

        tipo_norm_texto = normalizar_texto_comparable(tipo_raw)
        if arma_inferida in {"PISTOLA", "REVOLVER"}:
            tipo_fila = "CORTA"
        elif arma_inferida in {"CARABINA", "ESCOPETA"}:
            tipo_fila = "LARGA"
        elif "CORT" in tipo_norm_texto or tipo_inferido == "CORTA":
            tipo_fila = "CORTA"
        elif "LARG" in tipo_norm_texto or tipo_inferido == "LARGA":
            tipo_fila = "LARGA"
        else:
            continue

        if arma_inferida in armas_especificas:
            arma_objetivo = arma_inferida
        else:
            arma_objetivo = "PISTOLA" if tipo_fila == "CORTA" else "CARABINA"

        if tipo_fila not in tipos_arma_excel:
            tipos_arma_excel.append(tipo_fila)
        if arma_objetivo not in armas_excel:
            armas_excel.append(arma_objetivo)

        par_objetivo = (tipo_fila, arma_objetivo)
        if par_objetivo not in objetivos_arma:
            objetivos_arma.append(par_objetivo)

    if not objetivos_arma:
        # Fallback mínimo usando el primer registro, manteniendo origen en Excel.
        if arma_base in {"PISTOLA", "REVOLVER"}:
            tipo_base = "CORTA"
            arma_objetivo = arma_base
        elif arma_base in {"CARABINA", "ESCOPETA"}:
            tipo_base = "LARGA"
            arma_objetivo = arma_base
        elif tipo_arma_base == "LARGA":
            tipo_base = "LARGA"
            arma_objetivo = "CARABINA"
        else:
            tipo_base = "CORTA"
            arma_objetivo = "PISTOLA"

        objetivos_arma = [(tipo_base, arma_objetivo)]
        tipos_arma_excel = [tipo_base]
        armas_excel = [arma_objetivo]

    tipos_arma_objetivo = [t for t, _ in objetivos_arma]

    registro["fecha"] = fecha
    registro["hora_rango"] = hora_rango
    registro["doc_vigilante"] = doc_vigilante
    registro["objetivos_arma"] = objetivos_arma
    registro["tipos_arma_objetivo"] = tipos_arma_objetivo
    registro["armas_objetivo"] = armas_excel

    print("📄 Registro tomado desde Excel:")
    print(f"   • id_registro: {registro.get('id_registro', '')}")
    print(f"   • sede: {sede}")
    print(f"   • fecha: {fecha}")
    print(f"   • hora_rango: {hora_rango}")
    print(f"   • tipo_operacion: {tipo_operacion}")
    print(f"   • doc_vigilante: {doc_vigilante}")
    print(f"   • nro_solicitud: {nro_solicitud}")
    print(f"   • fecha_col_programacion: {fecha_col_programacion}")
    print(f"   • siguiente_mismo_doc_y_fecha: {siguiente_mismo_doc_y_fecha}")
    print(f"   • tipo_arma (excel): {tipos_arma_excel}")
    print(f"   • arma (excel): {armas_excel}")
    print(f"   • objetivos_arma: {objetivos_arma}")
    print(f"   • tipos_arma_objetivo: {tipos_arma_objetivo}")
    return registro


def seleccionar_en_selectonemenu(page, trigger_selector: str, panel_selector: str, label_selector: str, valor: str, nombre_campo: str):
    """Selecciona una opción PrimeFaces SelectOneMenu por data-label o texto visible."""
    trigger = page.locator(trigger_selector)
    trigger.wait_for(state="visible", timeout=12000)
    trigger.click()

    panel = page.locator(panel_selector)
    panel.wait_for(state="visible", timeout=7000)

    opcion = panel.locator(f'li.ui-selectonemenu-item[data-label="{valor}"]')
    try:
        opcion.wait_for(state="visible", timeout=2000)
    except PlaywrightTimeoutError:
        opcion = panel.locator("li.ui-selectonemenu-item").filter(has_text=valor)
        opcion.wait_for(state="visible", timeout=5000)

    opcion.first.click()
    page.wait_for_timeout(250)

    texto_label = page.locator(label_selector).inner_text().strip()
    if texto_label.upper() != valor.upper():
        raise Exception(
            f"No se confirmó la selección de {nombre_campo}. Esperado: '{valor}' | Actual: '{texto_label}'"
        )
    print(f"   ✓ {nombre_campo} seleccionado: {texto_label}")


# ============================================================
# NAVEGACIÓN: CITAS → RESERVAS DE CITAS
# ============================================================

def navegar_reservas_citas(page):
    """
    El menú de SUCAMEC es un PrimeFaces PanelMenu (acordeón).
    NO usa hover — se expande haciendo CLIC en el <h3> header.

    Flujo:
      1. Clic en el <h3> de "CITAS" para expandir el panel.
      2. Esperar a que el panel interno sea visible (display:block).
      3. Clic en el <a> de "RESERVAS DE CITAS" (dispara el submit JSF).
      4. Esperar a que la nueva vista cargue.
    """
    print("\n📋 Navegando a CITAS → RESERVAS DE CITAS...")

    # 1. Esperar carga base
    try:
        page.wait_for_load_state("domcontentloaded", timeout=8000)
    except Exception:
        pass

    def vista_reservas_lista(timeout_ms: int = 3500) -> bool:
        """Confirma que ya estamos en la vista donde aparece el combo 'Cita para'."""
        try:
            page.locator("form#gestionCitasForm").wait_for(state="visible", timeout=timeout_ms)
            page.locator(SEL["tipo_cita_trigger"]).wait_for(state="visible", timeout=timeout_ms)
            return True
        except Exception:
            return False

    # FAST PATH: clic directo al item menuid=7_1 dentro del panel lateral j_idt10.
    # Es más rápido porque evita expandir manualmente el acordeón CITAS.
    url_antes = page.url
    try:
        page.locator("#j_idt10").wait_for(state="visible", timeout=4000)
        click_directo = page.evaluate(
            '''() => {
                const link = document.querySelector('#j_idt10 a[onclick*="7_1"][onclick*="menuPrincipal"]');
                if (!link) return false;
                link.click();
                return true;
            }'''
        )
        if click_directo:
            print("   ⚡ Fast-path: click directo en 'RESERVAS DE CITAS' (menuid 7_1)")
            try:
                page.wait_for_load_state("networkidle", timeout=7000)
            except Exception:
                pass
            if ("GestionCitas.xhtml" in page.url) or (page.url != url_antes) or vista_reservas_lista(5000):
                print(f"✅ Navegación completada (fast-path) → URL: {page.url}")
                return
            print("   ⚠️ Fast-path no confirmó navegación → usando flujo estándar")
    except Exception:
        pass

    # ── PASO 1: Clic en el header "CITAS" del PanelMenu ──────────────────────
    # El header es el <h3> que contiene <a href="#" tabindex="-1">CITAS</a>
    # Usamos el <a> interno como punto de clic (más preciso).
    header_citas = page.locator(
        '#j_idt11\\:menuPrincipal .ui-panelmenu-header a[tabindex="-1"]'
    ).filter(has_text="CITAS")

    try:
        header_citas.wait_for(state="visible", timeout=5000)
    except PlaywrightTimeoutError:
        raise Exception("No se encontró el header 'CITAS' en el PanelMenu")

    header_citas.click()
    print("   ✓ Clic en header 'CITAS' → expandiendo panel...")

    # ── PASO 2: Esperar a que el panel de CITAS sea visible ──────────────────
    # El panel tiene id fijo: j_idt11:menuPrincipal_7
    # PrimeFaces lo muestra quitando la clase ui-helper-hidden y poniendo display:block
    panel_citas = page.locator('#j_idt11\\:menuPrincipal_7')
    try:
        # Esperar a que el panel sea visible (PrimeFaces hace toggle de display)
        panel_citas.wait_for(state="visible", timeout=2500)
        print("   ✓ Panel CITAS desplegado")
    except PlaywrightTimeoutError:
        # En algunas versiones de PF el panel ya está en el DOM pero con display:none
        # Forzamos visibilidad vía JS como fallback
        print("   ⚠️ Panel no visible por Playwright → forzando visibilidad vía JS")
        page.evaluate("""
            const panel = document.getElementById('j_idt11:menuPrincipal_7');
            if (panel) {
                panel.classList.remove('ui-helper-hidden');
                panel.style.display = 'block';
            }
        """)
        page.wait_for_timeout(180)

    # ── PASO 3: Clic en "RESERVAS DE CITAS" ──────────────────────────────────
    # Buscamos el <a> que contiene el span con texto "RESERVAS DE CITAS"
    # dentro del panel de CITAS ya desplegado.
    reservas_link = panel_citas.locator(
        'a.ui-menuitem-link:has(span.ui-menuitem-text:text-is("RESERVAS DE CITAS"))'
    )
    try:
        reservas_link.wait_for(state="visible", timeout=2500)
    except PlaywrightTimeoutError:
        # Fallback: buscar directamente por el onclick con menuid 7_1
        print("   ⚠️ Link no visible → usando fallback por menuid 7_1")
        reservas_link = page.locator(
            'a[onclick*="7_1"][onclick*="menuPrincipal"]'
        )
        reservas_link.wait_for(state="visible", timeout=3000)

    reservas_link.click()
    print("   ✓ Clic en 'RESERVAS DE CITAS'")

    # ── PASO 4: Esperar a que la nueva vista cargue ───────────────────────────
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass

    if not vista_reservas_lista(6000):
        raise Exception("No se confirmó la vista de 'Reservas de Citas' tras la navegación")

    print(f"✅ Navegación completada → URL: {page.url}")


def seleccionar_tipo_cita_poligono(page):
    """
    En la vista de Gestión de Citas, abre el SelectOneMenu de tipo de cita
    y selecciona la opción "EXAMEN PARA POLÍGONO DE TIRO".
    """
    print("\n🎯 Seleccionando tipo de cita: EXAMEN PARA POLÍGONO DE TIRO...")

    # Esperar que la vista de gestión esté lista
    page.locator("form#gestionCitasForm").wait_for(state="visible", timeout=12000)

    # 1) Abrir el combo (trigger)
    trigger = page.locator(SEL["tipo_cita_trigger"])
    try:
        trigger.wait_for(state="visible", timeout=6000)
        trigger.click()
    except PlaywrightTimeoutError:
        # Fallback: clic en el label del select para abrir panel
        print("   ⚠️ Trigger no visible → usando fallback sobre label")
        label = page.locator(SEL["tipo_cita_label"])
        label.wait_for(state="visible", timeout=6000)
        label.click()

    # 2) Esperar panel de opciones
    panel = page.locator(SEL["tipo_cita_panel"])
    panel.wait_for(state="visible", timeout=6000)

    # 3) Seleccionar opción de polígono
    opcion = page.locator(SEL["tipo_cita_opcion_poligono"])
    try:
        opcion.wait_for(state="visible", timeout=4000)
    except PlaywrightTimeoutError:
        print("   ⚠️ Opción por data-label no visible → buscando por texto")
        opcion = panel.locator("li.ui-selectonemenu-item").filter(has_text="EXAMEN PARA POLÍGONO DE TIRO")
        opcion.wait_for(state="visible", timeout=4000)

    opcion.click()

    # 4) Validar que el label del combo refleje la selección
    label = page.locator(SEL["tipo_cita_label"])
    page.wait_for_timeout(250)
    texto_label = label.inner_text().strip().upper()
    if "POLÍGONO DE TIRO" not in texto_label and "POLIGONO DE TIRO" not in texto_label:
        raise Exception(f"No se confirmó la selección en el combo. Label actual: '{texto_label}'")

    print(f"   ✓ Tipo de cita seleccionado: {texto_label}")


def seleccionar_sede_y_fecha_desde_registro(page, registro: dict):
    """
    En Reserva de Cupos, selecciona Sede y Fecha según el registro del Excel.
    """
    sede = registro["sede"].strip()
    fecha = registro["fecha"].strip()

    print("\n🧭 Completando Reserva de Cupos con datos del Excel...")
    page.locator(SEL["reserva_form"]).wait_for(state="visible", timeout=15000)

    seleccionar_en_selectonemenu(
        page,
        trigger_selector=SEL["sede_trigger"],
        panel_selector=SEL["sede_panel"],
        label_selector=SEL["sede_label"],
        valor=sede,
        nombre_campo="Sede"
    )

    # Al cambiar sede, PrimeFaces suele refrescar opciones de fecha por AJAX.
    page.wait_for_timeout(700)

    seleccionar_en_selectonemenu(
        page,
        trigger_selector=SEL["fecha_trigger"],
        panel_selector=SEL["fecha_panel"],
        label_selector=SEL["fecha_label"],
        valor=fecha,
        nombre_campo="Fecha"
    )


def seleccionar_hora_con_cupo_y_avanzar(page, registro: dict):
    """
    Busca la hora del Excel en la tabla de cupos, valida cupos > 0,
    selecciona el radiobutton de la fila y presiona 'Siguiente'.
    """
    hora_objetivo = normalizar_hora_rango(registro.get("hora_rango", ""))
    if not hora_objetivo:
        raise Exception("El registro no tiene 'hora_rango' válido")

    print(f"\n🕒 Buscando hora en tabla: {hora_objetivo}")

    tabla = page.locator(SEL["tabla_programacion"])
    tabla.wait_for(state="visible", timeout=15000)

    filas = page.locator(SEL["tabla_programacion_rows"])
    total_filas = filas.count()
    if total_filas == 0:
        raise Exception("La tabla de programación no tiene filas para la fecha/sede seleccionadas")

    fila_objetivo = None
    cupos_objetivo = 0
    resumen = []

    for i in range(total_filas):
        fila = filas.nth(i)
        celdas = fila.locator("td")
        if celdas.count() < 3:
            continue

        hora_tabla = normalizar_hora_rango(celdas.nth(0).inner_text().strip())
        cupos_texto = celdas.nth(1).inner_text().strip()
        cupos = convertir_a_entero(cupos_texto)
        resumen.append(f"{hora_tabla} ({cupos})")

        if hora_tabla == hora_objetivo:
            fila_objetivo = fila
            cupos_objetivo = cupos
            break

    if fila_objetivo is None:
        raise Exception(
            "No se encontró la hora objetivo en la tabla. "
            f"Objetivo: '{hora_objetivo}' | Disponibles: {', '.join(resumen)}"
        )

    if cupos_objetivo <= 0:
        raise Exception(f"La hora '{hora_objetivo}' no tiene cupos disponibles (Cupos Libres={cupos_objetivo})")

    radio_box = fila_objetivo.locator("td.ui-selection-column div.ui-radiobutton-box")
    if radio_box.count() == 0:
        raise Exception("No se encontró radiobutton en la fila de la hora objetivo")

    radio_box.first.click()
    page.wait_for_timeout(250)

    clase_radio = (radio_box.first.get_attribute("class") or "")
    aria_fila = (fila_objetivo.get_attribute("aria-selected") or "").lower()
    if "ui-state-active" not in clase_radio and aria_fila != "true":
        raise Exception("No se confirmó la selección del radiobutton de la hora")

    print(f"   ✓ Hora seleccionada: {hora_objetivo} (Cupos Libres={cupos_objetivo})")

    boton_siguiente = page.locator(SEL["boton_siguiente"])
    boton_siguiente.wait_for(state="visible", timeout=7000)
    boton_siguiente.click()
    print("   ✓ Click en botón 'Siguiente'")


def seleccionar_opcion_flexible_en_panel(page, panel_selector: str, texto_objetivo: str, nombre_campo: str):
    """Selecciona un li dentro de un panel PrimeFaces por coincidencia flexible de texto."""
    panel = page.locator(panel_selector)
    panel.wait_for(state="visible", timeout=7000)

    items = panel.locator("li.ui-selectonemenu-item")
    total = items.count()
    if total == 0:
        raise Exception(f"No hay opciones disponibles en {nombre_campo}")

    objetivo_norm = normalizar_texto_comparable(texto_objetivo)
    for i in range(total):
        item = items.nth(i)
        label = (item.get_attribute("data-label") or item.inner_text() or "").strip()
        label_norm = normalizar_texto_comparable(label)
        if objetivo_norm == label_norm or objetivo_norm in label_norm or label_norm in objetivo_norm:
            item.click()
            return label

    opciones = []
    for i in range(total):
        item = items.nth(i)
        opciones.append((item.get_attribute("data-label") or item.inner_text() or "").strip())
    raise Exception(
        f"No se encontró coincidencia para {nombre_campo}. "
        f"Objetivo: '{texto_objetivo}' | Opciones: {opciones}"
    )


def completar_paso_2_desde_registro(page, registro: dict):
    """
    Paso 2: tipo operación, doc. vigilante (autocomplete), seleccionar SI,
    y elegir número de solicitud por coincidencia con nro_solicitud del Excel.
    """
    tipo_operacion = registro.get("tipo_operacion", "").strip()
    doc_vigilante = registro.get("doc_vigilante", "").strip()
    nro_solicitud_excel = registro.get("nro_solicitud", "").strip()
    token_solicitud = extraer_token_solicitud(nro_solicitud_excel)

    print("\n🧩 Completando Paso 2 con datos del Excel...")

    # 2.1 Tipo de operación
    page.locator(SEL["tipo_operacion_trigger"]).wait_for(state="visible", timeout=12000)
    page.locator(SEL["tipo_operacion_trigger"]).click()
    page.locator(SEL["tipo_operacion_panel"]).wait_for(state="visible", timeout=7000)

    opcion_tipo = None
    items_tipo = page.locator(SEL["tipo_operacion_items"])
    total_tipo = items_tipo.count()
    objetivo_tipo = normalizar_texto_comparable(tipo_operacion)
    for i in range(total_tipo):
        item = items_tipo.nth(i)
        label = (item.get_attribute("data-label") or item.inner_text() or "").strip()
        label_norm = normalizar_texto_comparable(label)
        if objetivo_tipo == label_norm or objetivo_tipo in label_norm or label_norm in objetivo_tipo:
            item.click()
            opcion_tipo = label
            break
    if not opcion_tipo:
        raise Exception(f"No se encontró Tipo Operación '{tipo_operacion}' en el combo")

    page.wait_for_timeout(250)
    label_tipo = page.locator(SEL["tipo_operacion_label"]).inner_text().strip()
    if not label_tipo or label_tipo == "---":
        raise Exception("No se confirmó la selección de Tipo Operación")
    print(f"   ✓ Tipo Operación seleccionado: {opcion_tipo}")

    # 2.2 Documento de vigilante (autocomplete)
    doc_input = page.locator(SEL["doc_vig_input"])
    doc_input.wait_for(state="visible", timeout=12000)
    doc_input.click()
    doc_input.fill("")
    doc_input.type(doc_vigilante, delay=20)

    panel_doc = page.locator(SEL["doc_vig_panel"])
    items_doc = page.locator(SEL["doc_vig_items"])

    elegido = False
    try:
        panel_doc.wait_for(state="visible", timeout=2500)
    except PlaywrightTimeoutError:
        # Fallback: algunos autocompletes solo abren panel si se navega por teclado.
        doc_input.press("ArrowDown")
        page.wait_for_timeout(350)

    if panel_doc.is_visible():
        try:
            items_doc.first.wait_for(state="visible", timeout=2500)
        except PlaywrightTimeoutError:
            page.wait_for_timeout(700)

        total_doc = items_doc.count()
        for i in range(total_doc):
            item = items_doc.nth(i)
            data_label = (item.get_attribute("data-item-label") or "").strip()
            data_value = (item.get_attribute("data-item-value") or "").strip()
            texto_item = item.inner_text().strip()
            if doc_vigilante in data_label or doc_vigilante in data_value or doc_vigilante in texto_item:
                item.click()
                elegido = True
                break

        if not elegido and total_doc > 0:
            items_doc.first.click()
            elegido = True

    if not elegido:
        # Fallback final: forzar blur/change por si el valor exacto ya es aceptado por JSF.
        doc_input.evaluate(
            'el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); el.blur(); }'
        )

    page.wait_for_timeout(300)
    valor_doc = doc_input.input_value().strip()
    if doc_vigilante not in valor_doc:
        raise Exception(f"No se confirmó el documento vigilante. Esperado contiene '{doc_vigilante}' | Actual '{valor_doc}'")
    print(f"   ✓ Documento vigilante seleccionado: {valor_doc}")

    # 2.3 Seleccione Solicitud -> SI (siempre)
    page.locator(SEL["seleccione_solicitud_trigger"]).wait_for(state="visible", timeout=12000)
    page.locator(SEL["seleccione_solicitud_trigger"]).click()
    page.locator(SEL["seleccione_solicitud_panel"]).wait_for(state="visible", timeout=7000)
    page.locator(SEL["seleccione_solicitud_si"]).first.click()
    page.wait_for_timeout(350)
    label_si = page.locator(SEL["seleccione_solicitud_label"]).inner_text().strip().upper()
    if label_si.replace(" ", "") != "SI":
        raise Exception(f"No se confirmó Seleccione Solicitud = SI. Actual: '{label_si}'")
    print("   ✓ Seleccione Solicitud: SI")

    # 2.4 Nro Solicitud por coincidencia parcial (ej. 90086)
    if not token_solicitud:
        raise Exception(f"No se pudo extraer token numérico de nro_solicitud: '{nro_solicitud_excel}'")

    page.locator(SEL["nro_solicitud_trigger"]).wait_for(state="visible", timeout=12000)
    page.locator(SEL["nro_solicitud_trigger"]).click()

    panel_nro = page.locator(SEL["nro_solicitud_panel"])
    panel_nro.wait_for(state="visible", timeout=7000)
    items_nro = page.locator(SEL["nro_solicitud_items"])
    total_nro = items_nro.count()
    if total_nro == 0:
        raise Exception("No hay opciones en el combo de Nro Solicitud")

    seleccionado_label = None
    for i in range(total_nro):
        item = items_nro.nth(i)
        label = (item.get_attribute("data-label") or item.inner_text() or "").strip()
        # Comparamos contra todos los bloques numéricos del label para encontrar el Nro Empoce
        bloques = re.findall(r"\d+", label)
        bloques_norm = [b.lstrip("0") or "0" for b in bloques]
        if token_solicitud in bloques_norm:
            item.click()
            seleccionado_label = label
            break

    if not seleccionado_label:
        disponibles = []
        for i in range(total_nro):
            item = items_nro.nth(i)
            disponibles.append((item.get_attribute("data-label") or item.inner_text() or "").strip())
        raise Exception(
            f"No se encontró Nro Solicitud con token '{token_solicitud}'. Opciones: {disponibles}"
        )

    page.wait_for_timeout(300)
    label_nro = page.locator(SEL["nro_solicitud_label"]).inner_text().strip()
    bloques_final = [b.lstrip("0") or "0" for b in re.findall(r"\d+", label_nro)]
    if token_solicitud not in bloques_final:
        raise Exception(
            f"No se confirmó Nro Solicitud. Esperado token '{token_solicitud}' | Actual '{label_nro}'"
        )
    print(f"   ✓ Nro Solicitud seleccionado: {label_nro}")


def completar_tabla_tipos_arma_y_avanzar(page, registro: dict):
    """
    En Fase 2 completa la tabla dtTipoLic según tipo_arma del Excel y
    pulsa 'Siguiente' (botonSiguiente3).

    Reglas:
      - Si hay más de un registro del mismo usuario+fecha, se infiere misma programación
        y se aplican todos los tipos/armas encontrados.
      - Si hay solo un registro, se aplica solo ese.
    """
    print("\n🔫 Completando tabla de tipos de arma (Fase 2)...")

    objetivos_excel = registro.get("objetivos_arma", []) or []
    objetivos = []
    for item in objetivos_excel:
        if isinstance(item, (list, tuple)) and len(item) == 2:
            tipo_fila = normalizar_tipo_arma_excel(item[0])
            arma_objetivo = normalizar_tipo_arma_excel(item[1])
            if tipo_fila and arma_objetivo and (tipo_fila, arma_objetivo) not in objetivos:
                objetivos.append((tipo_fila, arma_objetivo))

    if not objetivos:
        raise Exception("No se recibieron objetivos de arma válidos desde Excel (tipo_arma + arma)")

    # PrimeFaces puede renderizar filas sin el sufijo _data y en modo editable por celda.
    filas = page.locator('#tabGestion\\:creaCitaPolJurForm\\:dtTipoLic tbody tr')
    try:
        filas.first.wait_for(state="visible", timeout=9000)
    except PlaywrightTimeoutError:
        filas = page.locator('table[id^="tabGestion:creaCitaPolJurForm:dtTipoLic"] tbody tr')
        try:
            filas.first.wait_for(state="visible", timeout=4000)
        except PlaywrightTimeoutError:
            raise Exception("No se encontró la tabla de tipos de arma (dtTipoLic)")

    total_filas = filas.count()
    if total_filas == 0:
        raise Exception("La tabla dtTipoLic no tiene filas")

    aplicados = []
    for tipo_fila, arma_objetivo in objetivos:
        fila_match = None
        for i in range(total_filas):
            fila = filas.nth(i)
            celdas = fila.locator('td[role="gridcell"]')
            if celdas.count() == 0:
                celdas = fila.locator("td")

            textos = []
            for j in range(celdas.count()):
                texto_celda = normalizar_texto_comparable(celdas.nth(j).inner_text().strip())
                if texto_celda:
                    textos.append(texto_celda)

            tipo_texto = " ".join(textos)
            if tipo_fila in tipo_texto:
                fila_match = fila
                break

        if fila_match is None:
            raise Exception(f"No se encontró fila para tipo de arma '{tipo_fila}' en dtTipoLic")

        # La columna "Arma" es editable; activamos la celda para mostrar el select.
        celdas_editables = fila_match.locator("td.ui-editable-column")
        if celdas_editables.count() > 0:
            celdas_editables.last.click()
            page.wait_for_timeout(180)

        combo = fila_match.locator("select")
        if combo.count() == 0:
            raise Exception(f"No se encontró combo de Arma para tipo '{tipo_fila}'")

        combo.first.wait_for(state="visible", timeout=3500)
        combo.first.select_option(label=arma_objetivo)
        page.wait_for_timeout(350)

        try:
            page.wait_for_load_state("networkidle", timeout=3500)
        except Exception:
            pass

        seleccionado = combo.first.evaluate(
            "el => el.options[el.selectedIndex] ? el.options[el.selectedIndex].text.trim() : ''"
        )
        if normalizar_texto_comparable(seleccionado) != normalizar_texto_comparable(arma_objetivo):
            raise Exception(
                f"No se confirmó Arma para '{tipo_fila}'. Esperado '{arma_objetivo}' | Actual '{seleccionado}'"
            )

        aplicados.append(f"{tipo_fila} -> {seleccionado}")
        print(f"   ✓ {tipo_fila}: {seleccionado}")

    if not aplicados:
        raise Exception("No se aplicó ninguna selección de arma en dtTipoLic")

    boton_siguiente_3 = page.locator('#tabGestion\\:creaCitaPolJurForm\\:botonSiguiente3')
    boton_siguiente_3.wait_for(state="visible", timeout=8000)
    boton_siguiente_3.click()
    print("   ✓ Click en botón 'Siguiente' de Fase 2 (botonSiguiente3)")

    try:
        page.wait_for_load_state("networkidle", timeout=7000)
    except Exception:
        pass


# ============================================================
# FLUJO PRINCIPAL
# ============================================================

def llenar_login_sel():
    print("🚀 INICIANDO SCRIPT SEL - Login Automático")

    inicio_total_flujo = time.time()
    duracion_total_flujo = None

    registro_excel = cargar_primer_registro_pendiente_desde_excel(EXCEL_PATH)

    playwright = sync_playwright().start()
    browser = None
    login_exitoso = False

    try:
        for intento_global in range(3):
            start_time = time.time()
            print(f"\n🔄 Intento global {intento_global+1}/3")

            if browser is not None:
                try:
                    browser.close()
                except Exception:
                    pass

            browser = playwright.chromium.launch(
                headless=False,
                slow_mo=0,
                args=[
                    "--start-maximized",
                    "--disable-infobars",
                    "--window-size=1920,1080",
                    "--window-position=0,0"
                ]
            )
            context = browser.new_context(viewport=None, ignore_https_errors=True)
            page = context.new_page()
            page.evaluate("() => { window.moveTo(0, 0); window.resizeTo(screen.width, screen.height); }")

            try:
                page.goto(URL_LOGIN, wait_until="domcontentloaded", timeout=45000)
                print("1. Página de login cargada")

                tab = page.locator(SEL["tab_tradicional"])
                tab.wait_for(state="visible", timeout=8000)
                tab.click()
                print("2. Pestaña 'Autenticación Tradicional' seleccionada")

                page.locator(SEL["numero_documento"]).wait_for(state="visible", timeout=8000)

                page.select_option(SEL["tipo_doc_select"], value=CREDENCIALES["tipo_documento_valor"])
                # El onchange del tipo de documento dispara un update AJAX sobre el input documento.
                # Esperamos que ese update termine antes de escribir para evitar valores parciales.
                page.wait_for_timeout(450)
                page.locator(SEL["numero_documento"]).wait_for(state="visible", timeout=8000)
                escribir_input_jsf(page, SEL["numero_documento"], CREDENCIALES["numero_documento"])
                escribir_input_rapido(page, SEL["usuario"], CREDENCIALES["usuario"])
                escribir_input_rapido(page, SEL["clave"], CREDENCIALES["contrasena"])
                print("✅ Credenciales llenadas")

                captcha_text = solve_captcha_ocr(page)
                if captcha_text and len(captcha_text) == 5:
                    escribir_input_rapido(page, SEL["captcha_input"], captcha_text)
                    print(f"✅ CAPTCHA automático: {captcha_text}")
                else:
                    solve_captcha_manual(page)

                print("🔘 Enviando login...")
                page.locator(SEL["ingresar"]).click(timeout=10000)

                print("⏳ Validando acceso...")
                url_ok = False
                for _ in range(50):
                    if "/aplicacion/" in page.url:
                        url_ok = True
                        break
                    page.wait_for_timeout(200)

                if url_ok:
                    total_time = time.time() - start_time
                    print(f"🎉 ¡ACCESO EXITOSO!")
                    print(f"   → URL: {page.url}")
                    print(f"⏱️ Tiempo total login: {total_time:.2f} segundos")
                    login_exitoso = True

                    # ── NAVEGAR A CITAS → RESERVAS DE CITAS ──────────────────
                    navegar_reservas_citas(page)

                    # ── SELECCIONAR TIPO DE CITA: EXAMEN PARA POLÍGONO ───────
                    seleccionar_tipo_cita_poligono(page)

                    # ── COMPLETAR RESERVA DE CUPOS: SEDE Y FECHA (Excel) ─────
                    seleccionar_sede_y_fecha_desde_registro(page, registro_excel)

                    # ── SELECCIONAR HORA (si hay cupos) Y AVANZAR ─────────────
                    seleccionar_hora_con_cupo_y_avanzar(page, registro_excel)

                    # ── COMPLETAR PASO 2 (TIPO OP, DNI, SI, NRO SOLICITUD) ───
                    completar_paso_2_desde_registro(page, registro_excel)

                    # ── FASE 2 FINAL: TABLA TIPO ARMA + SIGUIENTE ────────────
                    completar_tabla_tipos_arma_y_avanzar(page, registro_excel)

                    duracion_total_flujo = time.time() - inicio_total_flujo
                    print(f"\n⏱️ Tiempo total del flujo (inicio → fin Paso 2): {duracion_total_flujo:.2f} segundos")

                    break
                else:
                    print(f"❌ Login falló - URL NO cambió a /aplicacion/")
                    print(f"   → URL actual: {page.url}")
                    raise Exception("CAPTCHA incorrecto o credenciales inválidas")

            except Exception as e:
                print(f"❌ Intento {intento_global+1} falló: {e}")
                if intento_global < 2:
                    print("   Reintentando...")
                    time.sleep(1)
                else:
                    print("   Se agotaron los 3 intentos")

        if login_exitoso:
            print("\n✅ Flujo completado. Navegador abierto para uso manual.")
            if duracion_total_flujo is not None:
                print(f"   ⏱️ Duración final del flujo: {duracion_total_flujo:.2f} segundos")
            print("   Presiona Ctrl+C o cierra la ventana cuando termines.")
            try:
                while True:
                    time.sleep(60)
            except KeyboardInterrupt:
                print("\n🛑 Interrupción manual. Cerrando navegador...")
        else:
            print("\n❌ No se pudo completar el login después de todos los intentos.")
            input("   Presiona ENTER para cerrar el navegador...")

    finally:
        try:
            if browser is not None:
                browser.close() 
        except Exception:
            pass
        try:
            playwright.stop()
        except Exception:
            pass
        print("Navegador cerrado.")


if __name__ == "__main__":
    llenar_login_sel()