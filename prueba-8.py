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
}

def corregir_captcha_ocr(texto_raw: str) -> str:
    """
    Limpia el texto OCR del CAPTCHA:
    - Normaliza a mayúsculas
    - Elimina espacios y caracteres no alfanuméricos (basura OCR)
    
    NOTA: No se aplican sustituciones letra↔dígito porque el CAPTCHA
    de SUCAMEC usa AMBOS (letras y dígitos mezclados), por lo tanto
    no se puede determinar la dirección correcta de corrección.
    La precisión depende del preprocesamiento de imagen + config de Tesseract.
    """
    if not texto_raw:
        return ""
    
    texto = texto_raw.strip().upper().replace(" ", "").replace("\n", "").replace("\r", "")
    
    # Eliminar caracteres que no son alfanuméricos (basura OCR)
    texto = ''.join(c for c in texto if c.isalnum())
    
    return texto


def validar_captcha_texto(texto: str) -> bool:
    """
    Valida que el texto del CAPTCHA cumpla los requisitos:
    - Exactamente 5 caracteres
    - Solo alfanuméricos (0-9, A-Z)
    """
    if not texto or len(texto) != 5:
        return False
    return texto.isalnum()


def escribir_input_jsf(page, selector: str, valor: str):
    """
    Escribe en un input JSF usando type() (genera eventos de teclado reales).
    Verifica que el valor se haya escrito y reintenta si está vacío.
    """
    campo = page.locator(selector)
    campo.wait_for(state="visible", timeout=10000)
    
    for intento in range(3):
        campo.click()
        campo.press("Control+A")
        campo.press("Backspace")
        campo.type(valor, delay=10)
        campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
        campo.blur()
        
        # Verificar que el valor se escribió correctamente
        valor_actual = campo.input_value()
        if valor_actual == valor:
            return
        
        print(f"   ⚠️ Campo {selector}: esperado '{valor}', tiene '{valor_actual}' → reintentando ({intento+1}/3)")
        page.wait_for_timeout(200)
    
    # Último intento con fill() como fallback
    campo.click()
    campo.fill(valor)
    campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
    campo.blur()


def escribir_input_rapido(page, selector: str, valor: str):
    """
    Escribe en un input usando fill() (rápido).
    Verifica que el valor se haya escrito.
    """
    campo = page.locator(selector)
    campo.wait_for(state="visible", timeout=10000)
    campo.click()
    campo.fill(valor)
    campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
    campo.blur()
    
    # Verificar valor
    valor_actual = campo.input_value()
    if valor_actual != valor:
        # Reintentar con type()
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
    """
    Preprocesa la imagen del CAPTCHA para mejorar la precisión del OCR.
    Aplica múltiples técnicas de procesamiento de imagen.
    
    variante: permite probar distintas configuraciones de preprocesamiento
              para mejorar la tasa de acierto en CAPTCHAs difíciles.
    """
    img = Image.open(BytesIO(img_bytes))
    
    # Convertir a escala de grises
    img = img.convert('L')
    
    if variante == 0:
        # Variante principal: alto contraste + escala grande
        img = ImageEnhance.Contrast(img).enhance(3.5)
        w, h = img.size
        img = img.resize((w * 4, h * 4), Image.LANCZOS)
        img = img.filter(ImageFilter.MedianFilter(size=3))
        img = ImageOps.invert(img)
        img = img.point(lambda p: 255 if p > 130 else 0)
        img = ImageEnhance.Sharpness(img).enhance(3.0)
    elif variante == 1:
        # Variante alternativa: contraste moderado + sin inversión
        img = ImageEnhance.Contrast(img).enhance(2.5)
        w, h = img.size
        img = img.resize((w * 3, h * 3), Image.LANCZOS)
        img = img.filter(ImageFilter.MedianFilter(size=5))
        img = img.point(lambda p: 255 if p > 160 else 0)
        img = ImageEnhance.Sharpness(img).enhance(2.0)
    else:
        # Variante 2: escala mayor + umbral más bajo
        img = ImageEnhance.Contrast(img).enhance(4.0)
        w, h = img.size
        img = img.resize((w * 5, h * 5), Image.LANCZOS)
        img = img.filter(ImageFilter.GaussianBlur(radius=0.5))
        img = ImageOps.invert(img)
        img = img.point(lambda p: 255 if p > 110 else 0)
        img = ImageEnhance.Sharpness(img).enhance(4.0)
    
    return img


def solve_captcha_ocr(page):
    """
    Intenta resolver el CAPTCHA usando OCR con múltiples intentos.
    Prueba múltiples variantes de preprocesamiento y configuraciones PSM
    para maximizar la tasa de acierto.
    """
    if not OCR_AVAILABLE:
        return None
    
    MAX_INTENTOS = 6
    PSM_MODES = [7, 8, 13]  # línea, palabra, carácter crudo
    NUM_VARIANTES = 3       # variantes de preprocesamiento
    
    for intento in range(MAX_INTENTOS):
        try:
            print(f"🔍 Intentando resolver CAPTCHA (intento {intento+1}/{MAX_INTENTOS})...")
            
            # Espera mínima para que la imagen cargue
            page.wait_for_timeout(200)
            
            # Capturar screenshot del CAPTCHA
            img_bytes = page.locator(SEL["captcha_img"]).screenshot(type="png")
            
            # Probar múltiples combinaciones de preprocesamiento + PSM
            mejor_texto = None
            for variante in range(NUM_VARIANTES):
                img = preprocesar_imagen_captcha(img_bytes, variante=variante)
                
                for psm in PSM_MODES:
                    config = f'--psm {psm} --oem 3 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ --dpi 300'
                    
                    texto_raw = pytesseract.image_to_string(
                        img, config=config, lang='eng'
                    ).strip()
                    
                    texto = corregir_captcha_ocr(texto_raw)
                    
                    if validar_captcha_texto(texto):
                        print(f"   → Variante {variante}, PSM {psm}: '{texto_raw}' → '{texto}' ✓")
                        mejor_texto = texto
                        break  # encontró válido en este PSM
                    else:
                        print(f"   → Variante {variante}, PSM {psm}: '{texto_raw}' → '{texto}' (len={len(texto)}) ✗")
                
                if mejor_texto:
                    break  # encontró válido en esta variante
            
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


def llenar_login_sel():
    """
    Flujo principal de login automatizado en SUCAMEC-SEL.
    El navegador permanece abierto tras el login exitoso.
    """
    print("🚀 INICIANDO SCRIPT SEL - Login Automático")
    
    # Iniciar Playwright FUERA del with para controlar cuándo se cierra
    playwright = sync_playwright().start()
    browser = None
    login_exitoso = False
    
    try:
        for intento_global in range(3):
            start_time = time.time()
            print(f"\n🔄 Intento global {intento_global+1}/3")
            
            # Si hay un browser previo de un intento fallido, cerrarlo
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
            context = browser.new_context(
                viewport=None,
                ignore_https_errors=True
            )
            page = context.new_page()
            
            # Maximizar ventana
            page.evaluate("() => { window.moveTo(0, 0); window.resizeTo(screen.width, screen.height); }")
            
            try:
                page.goto(URL_LOGIN, wait_until="domcontentloaded", timeout=45000)
                print("1. Página de login cargada")
                
                # Pestaña Autenticación Tradicional – siempre hacer clic
                # (necesario en carga inicial y en reintentos tras CAPTCHA fallido)
                tab = page.locator(SEL["tab_tradicional"])
                tab.wait_for(state="visible", timeout=8000)
                tab.click()
                print("2. Pestaña 'Autenticación Tradicional' seleccionada")
                
                # Esperar a que el formulario esté visible
                page.locator(SEL["numero_documento"]).wait_for(state="visible", timeout=8000)
                
                # Llenar credenciales
                page.select_option(SEL["tipo_doc_select"], value=CREDENCIALES["tipo_documento_valor"])
                escribir_input_jsf(page, SEL["numero_documento"], CREDENCIALES["numero_documento"])
                escribir_input_rapido(page, SEL["usuario"], CREDENCIALES["usuario"])
                escribir_input_rapido(page, SEL["clave"], CREDENCIALES["contrasena"])
                
                print("✅ Credenciales llenadas")
                
                # CAPTCHA
                captcha_text = solve_captcha_ocr(page)
                if captcha_text and len(captcha_text) == 5:
                    escribir_input_rapido(page, SEL["captcha_input"], captcha_text)
                    print(f"✅ CAPTCHA automático: {captcha_text}")
                else:
                    solve_captcha_manual(page)
                
                # Enviar login
                print("🔘 Enviando login...")
                page.locator(SEL["ingresar"]).click(timeout=10000)
                
                # === VALIDACIÓN DE ACCESO POR URL ===
                print("⏳ Validando acceso...")
                
                # Polling rápido: verificar URL cada 200ms, máximo 10s
                url_ok = False
                for _ in range(50):  # 50 × 200ms = 10s
                    if "/aplicacion/" in page.url:
                        url_ok = True
                        break
                    page.wait_for_timeout(200)
                
                if url_ok:
                    total_time = time.time() - start_time
                    print(f"🎉 ¡ACCESO EXITOSO!")
                    print(f"   → URL: {page.url}")
                    print(f"⏱️ Tiempo total: {total_time:.2f} segundos")
                    login_exitoso = True
                    break  # Salir del loop de reintentos
                else:
                    # Login falló (CAPTCHA incorrecto u otra razón)
                    print(f"❌ Login falló - la URL NO cambió a /aplicacion/")
                    print(f"   → URL actual: {page.url}")
                    raise Exception("CAPTCHA incorrecto o credenciales inválidas")
                
            except Exception as e:
                print(f"❌ Intento {intento_global+1} falló: {e}")
                if intento_global < 2:
                    print("   Reintentando...")
                    time.sleep(1)
                else:
                    print("   Se agotaron los 3 intentos")
        
        # ========================================================
        # MANTENER NAVEGADOR ABIERTO después del login exitoso
        # ========================================================
        if login_exitoso:
            print("\n✅ Flujo completado. El navegador quedará abierto para que puedas usarlo.")
            print("   Presiona Ctrl+C en la terminal o cierra la ventana manualmente cuando termines.")
            
            try:
                while True:
                    time.sleep(60)
            except KeyboardInterrupt:
                print("\n🛑 Interrupción manual detectada. Cerrando navegador...")
        else:
            print("\n❌ No se pudo completar el login después de todos los intentos.")
            input("   Presiona ENTER para cerrar el navegador...")
    
    finally:
        # Limpieza segura
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