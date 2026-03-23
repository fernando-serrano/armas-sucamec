from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, args=["--start-maximized"])
    page = browser.new_page()
    page.goto("https://www.google.com")
    print("Navegador debería estar abierto ahora")
    input("Presiona ENTER para cerrar...")
    browser.close()