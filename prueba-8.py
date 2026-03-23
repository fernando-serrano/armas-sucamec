import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import time

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
    campo = page.locator(selector)
    campo.wait_for(state="visible", timeout=10000)
    for intento in range(3):
        campo.click()
        campo.press("Control+A")
        campo.press("Backspace")
        campo.type(valor, delay=10)
        campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
        campo.blur()
        if campo.input_value() == valor:
            return
        print(f"   ⚠️ Campo {selector}: esperado '{valor}', tiene '{campo.input_value()}' → reintentando ({intento+1}/3)")
        page.wait_for_timeout(200)
    campo.click()
    campo.fill(valor)
    campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
    campo.blur()


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

    # 1. Esperar carga completa de inicio.xhtml
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass

    # ── PASO 1: Clic en el header "CITAS" del PanelMenu ──────────────────────
    # El header es el <h3> que contiene <a href="#" tabindex="-1">CITAS</a>
    # Usamos el <a> interno como punto de clic (más preciso).
    header_citas = page.locator(
        '#j_idt11\\:menuPrincipal .ui-panelmenu-header a[tabindex="-1"]'
    ).filter(has_text="CITAS")

    try:
        header_citas.wait_for(state="visible", timeout=8000)
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
        panel_citas.wait_for(state="visible", timeout=5000)
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
        page.wait_for_timeout(300)

    # ── PASO 3: Clic en "RESERVAS DE CITAS" ──────────────────────────────────
    # Buscamos el <a> que contiene el span con texto "RESERVAS DE CITAS"
    # dentro del panel de CITAS ya desplegado.
    reservas_link = panel_citas.locator(
        'a.ui-menuitem-link:has(span.ui-menuitem-text:text-is("RESERVAS DE CITAS"))'
    )
    try:
        reservas_link.wait_for(state="visible", timeout=5000)
    except PlaywrightTimeoutError:
        # Fallback: buscar directamente por el onclick con menuid 7_1
        print("   ⚠️ Link no visible → usando fallback por menuid 7_1")
        reservas_link = page.locator(
            'a[onclick*="7_1"][onclick*="menuPrincipal"]'
        )
        reservas_link.wait_for(state="visible", timeout=5000)

    reservas_link.click()
    print("   ✓ Clic en 'RESERVAS DE CITAS'")

    # ── PASO 4: Esperar a que la nueva vista cargue ───────────────────────────
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass

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


# ============================================================
# FLUJO PRINCIPAL
# ============================================================

def llenar_login_sel():
    print("🚀 INICIANDO SCRIPT SEL - Login Automático")

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