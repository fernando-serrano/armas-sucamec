import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ====================== INTENTO DE IMPORTAR OCR (opcional) ======================
OCR_AVAILABLE = False
try:
    from PIL import Image, ImageFilter, ImageEnhance
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
    "captcha_img": "#tabViewLogin\\:tradicionalForm\\:imgCaptcha",          # ← para OCR
    "captcha_input": "#tabViewLogin\\:tradicionalForm\\:textoCaptcha",
    "boton_refresh": "#tabViewLogin\\:tradicionalForm\\:botonCaptcha",      # ← refrescar
    "ingresar": "#tabViewLogin\\:tradicionalForm\\:ingresar",
}

def escribir_input_jsf(page, selector: str, valor: str, delay: int = 35):
    campo = page.locator(selector)
    campo.wait_for(state="visible", timeout=10000)
    campo.click()
    campo.press("Control+A")
    campo.press("Backspace")
    campo.type(valor, delay=delay)
    campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
    campo.blur()
    page.wait_for_timeout(300)

def escribir_input_rapido(page, selector: str, valor: str):
    campo = page.locator(selector)
    campo.wait_for(state="visible", timeout=10000)
    campo.click()
    campo.press("Control+A")
    campo.press("Backspace")
    campo.fill(valor)
    campo.evaluate('el => { el.dispatchEvent(new Event("input", {bubbles:true})); el.dispatchEvent(new Event("change", {bubbles:true})); }')
    campo.blur()
    page.wait_for_timeout(150)

def solve_captcha_manual(page):
    print("\n🔴 MODO MANUAL ACTIVADO")
    print("Llena el código de verificación en la ventana del navegador")
    input("✅ Cuando hayas escrito el CAPTCHA → presiona ENTER aquí para continuar...")

# ... (imports y resto igual)

def solve_captcha_ocr(page):
    if not OCR_AVAILABLE:
        return None
    
    for intento in range(4):
        try:
            print(f"🔍 Intentando resolver CAPTCHA automáticamente (intento {intento+1}/4)...")
            
            img_bytes = page.locator(SEL["captcha_img"]).screenshot(type="png")
            img = Image.open(BytesIO(img_bytes))
            
            # (tu preprocesamiento actual aquí, sin cambios)
            img = img.convert('L')
            img = ImageEnhance.Contrast(img).enhance(3.0)
            img = img.filter(ImageFilter.MedianFilter(size=3))
            img = img.filter(ImageFilter.GaussianBlur(radius=1.0))
            from PIL import ImageOps
            img = ImageOps.invert(img)
            img = img.point(lambda p: 255 if p > 140 else 0)
            img = ImageEnhance.Sharpness(img).enhance(2.5)
            
            custom_config = r'--psm 8 --oem 3 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ --dpi 300'
            
            texto = pytesseract.image_to_string(
                img,
                config=custom_config,
                lang='eng'
            ).strip().upper().replace(" ", "").replace("\n", "").replace("\r", "")
            
            print(f"   → Detectado: '{texto}' (longitud: {len(texto)})")
            
            if len(texto) == 5 and texto.isalnum():
                print(f"   ✓ CAPTCHA válido → Usando: {texto}")
                return texto
                
            print("   ✗ No válido → Refrescando CAPTCHA...")
            print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
            page.locator(SEL["boton_refresh"]).click(force=True)
            page.wait_for_timeout(2500)
            
        except Exception as e:
            print(f"   Error en intento {intento+1}: {str(e)}")
            page.wait_for_timeout(1500)
    
    print("❌ No se pudo resolver automáticamente después de 4 intentos → modo manual")
    return None


def llenar_login_sel():
    # ... (parte inicial igual: browser, page, goto, tab tradicional, llenado credenciales)
    
    print("✅ Credenciales llenadas correctamente")
    
    # CAPTCHA
    captcha_text = solve_captcha_ocr(page) if OCR_AVAILABLE else None
    
    if captcha_text and len(captcha_text) == 5:
        escribir_input_rapido(page, SEL["captcha_input"], captcha_text)
        print(f"✅ CAPTCHA automático ingresado: {captcha_text}")
    else:
        solve_captcha_manual(page)   # tu función manual
    
    print("🔘 Enviando login...")
    page.locator(SEL["ingresar"]).click(timeout=10000)
    print("   → Clic en 'Ingresar al Sistema' realizado")
    
    # ================================================
    # NUEVA PARTE: Esperar y confirmar acceso exitoso
    # ================================================
    try:
        print("⏳ Esperando redirección o carga de la página principal (máx. 15 segundos)...")
        
        # Opción 1: Esperar cambio de URL (la más confiable si redirige)
        original_url = page.url
        page.wait_for_url("**sucamec.gob.pe/sel/faces/**dashboard** OR **principal** OR **home** OR **!**/login**", 
                          timeout=15000, wait_until="domcontentloaded")
        
        if page.url != original_url:
            print("🎉 ¡ACCESO EXITOSO!")
            print(f"   → Redirigido a: {page.url}")
            print("   Estás dentro del sistema SEL.")
        else:
            print("⚠️ No se detectó cambio de URL → posible error de login o página sin redirección.")
        
        # Opción alternativa (si no redirige y queda en la misma URL pero cambia contenido):
        # Descomenta y ajusta el selector si sabes uno que aparece SOLO después del login
        # page.wait_for_selector("#algun-selector-del-dashboard", timeout=10000)
        # print("🎉 Elemento del dashboard detectado → ACCESO EXITOSO!")
        
        page.wait_for_timeout(3000)  # tiempo extra para ver la página
        
    except Exception as e:
        print(f"⚠️ No se pudo confirmar acceso exitoso: {str(e)}")
        print("   → Revisa manualmente si entraste o si hubo error (captcha equivocado, credenciales, etc.)")
    
    # ... (finally con input para cerrar)

def llenar_login_sel():
    print("🚀 INICIANDO SCRIPT SEL - Login Automático")
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            slow_mo=100,                    # ← más lento para que veas todo
            args=["--start-maximized", "--disable-infobars"]
        )
        context = browser.new_context(no_viewport=True)
        page = context.new_page()
        
        try:
            print("1. Abriendo página de login...")
            page.goto(URL_LOGIN, wait_until="domcontentloaded", timeout=45000)
            print("2. Página cargada")

            # Pestaña tradicional
            if not page.locator(SEL["numero_documento"]).is_visible(timeout=5000):
                print("3. Cambiando a pestaña Tradicional...")
                page.locator(SEL["tab_tradicional"]).click(timeout=8000)

            # Llenar campos
            page.select_option(SEL["tipo_doc_select"], value=CREDENCIALES["tipo_documento_valor"])
            escribir_input_jsf(page, SEL["numero_documento"], CREDENCIALES["numero_documento"])
            escribir_input_rapido(page, SEL["usuario"], CREDENCIALES["usuario"])
            escribir_input_rapido(page, SEL["clave"], CREDENCIALES["contrasena"])
            
            print("✅ Credenciales llenadas correctamente")

            # === CAPTCHA ===
            captcha_text = solve_captcha_ocr(page) if OCR_AVAILABLE else None
            
            if captcha_text and len(captcha_text) == 5:
                escribir_input_rapido(page, SEL["captcha_input"], captcha_text)
                print(f"✅ CAPTCHA automático: {captcha_text}")
            else:
                solve_captcha_manual(page)   # pausa para que lo hagas a mano

            # Click final
            page.locator(SEL["ingresar"]).click(timeout=10000)
            print("✅ Clic en 'Ingresar al Sistema' enviado")
            page.wait_for_timeout(7000)

        except Exception as e:
            print(f"❌ ERROR: {type(e).__name__} - {e}")
            import traceback
            traceback.print_exc()
            page.screenshot(path="error_screenshot.png")
            print("📸 Screenshot guardado: error_screenshot.png")
        finally:
            input("\nPresiona ENTER para cerrar el navegador...")
            browser.close()

if __name__ == "__main__":
    llenar_login_sel()