from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

URL = "https://marketplace.microsoft.com/en-us/marketplace/partner-dir?filter=products%3DDeveloperTools%252CDynamics365Enterprise%3Bservices%3DIntegration%252CLicensing%252CManaged%2520Services%2520(MSP)%252CDeployment%2520or%2520Migration%252CConsulting%3Bsort%3D0%3BpageSize%3D18%3BonlyThisCountry%3Dtrue%3Bcountry%3DMX%3Bradius%3D100%3Blocname%3DMexico%3BlocationNotRequired%3Dtrue"

def find_results_frame(page, timeout_ms=20000):
    """
    Find the frame that actually contains the partner cards (p.name).
    """
    page.wait_for_timeout(2000)
    waited = 0
    step = 500

    while waited < timeout_ms:
        for fr in page.frames:
            try:
                # If the frame contains p.name, it's the one we need
                if fr.locator("p.name").count() > 0:
                    return fr
            except Exception:
                pass
        page.wait_for_timeout(step)
        waited += step

    return None

def extract_names(frame):
    names = frame.locator("p.name").all_inner_texts()
    return [n.strip() for n in names if n and n.strip()]

def find_next_button(frame):
    """
    Find a "Next" control inside the SAME frame.
    We try multiple selectors because UI varies.
    """
    candidates = [
        "button[aria-label*='Next']",
        "a[aria-label*='Next']",
        "button[aria-label*='Siguiente']",
        "a[aria-label*='Siguiente']",
        "button:has-text('Next')",
        "a:has-text('Next')",
        "button:has-text('Siguiente')",
        "a:has-text('Siguiente')",
        "button:has-text('›')",
        "a:has-text('›')",
        "button:has-text('>')",
        "a:has-text('>')",
    ]
    for sel in candidates:
        loc = frame.locator(sel)
        if loc.count() > 0:
            btn = loc.first
            # Check disabled states
            try:
                if btn.is_disabled():
                    return None
            except Exception:
                aria_disabled = btn.get_attribute("aria-disabled")
                if aria_disabled == "true":
                    return None
            return btn
    return None

def main():
    all_names = set()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        print("Opening URL...", flush=True)
        page.goto(URL, wait_until="domcontentloaded")

        # Find the frame that contains results
        frame = find_results_frame(page)
        if not frame:
            print("ERROR: No encontré el frame con 'p.name'.", flush=True)
            print("Tip: revisa si aparece un pop-up de cookies o si los resultados no cargan.", flush=True)
            browser.close()
            return

        print("OK: Encontré el frame de resultados:", frame.url, flush=True)

        page_num = 1
        while True:
            # Ensure results exist
            try:
                frame.wait_for_selector("p.name", timeout=15000)
            except PWTimeout:
                print("No aparecen p.name en el frame. Deteniendo.", flush=True)
                break

            names = extract_names(frame)
            before = len(all_names)
            all_names.update(names)
            after = len(all_names)

            print(f"[Page {page_num}] names_on_page={len(names)} total_unique={after} (+{after - before})", flush=True)

            # Try next page inside frame
            next_btn = find_next_button(frame)
            if not next_btn:
                print("No encontré Next (o está deshabilitado). Terminé.", flush=True)
                break

            # Detect page change by first item
            first_before = names[0] if names else ""

            next_btn.click()
            page.wait_for_timeout(2500)

            # Wait until the first name changes (avoid infinite loop)
            changed = False
            for _ in range(30):
                new_names = extract_names(frame)
                first_after = new_names[0] if new_names else ""
                if first_after and first_after != first_before:
                    changed = True
                    break
                page.wait_for_timeout(500)

            if not changed:
                print("Click en Next, pero los resultados no cambiaron. Detengo para evitar loop.", flush=True)
                break

            page_num += 1

        out = sorted(all_names)
        with open("partners_names.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(out))

        print(f"\nTOTAL FINAL: {len(out)}", flush=True)
        print("Archivo generado: partners_names.txt (en el directorio donde ejecutaste el script)", flush=True)

        browser.close()

if __name__ == "__main__":
    main()
