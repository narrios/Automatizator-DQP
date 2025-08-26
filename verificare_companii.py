import json
import os
import random
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import pandas as pd
import re
import time
import threading
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

# === Load country phone codes safely ===
try:
    with open("all_country_phone_codes.json", "r", encoding="utf-8") as f:
        country_codes = json.load(f)
except Exception:
    country_codes = {}

# === Phone patterns and helpers ===
PHONE_TEXT_RE = re.compile(r'\+?\d[\d\s\-\.\(\)\/]{6,}\d')
PHONE_TEL_RE  = re.compile(r'href=["\']\s*tel:([+0-9\-\s\.\(\)\/]+)["\']', re.I)

# --- Ban words around numbers that are NOT phones (tax IDs, banking, GDPR etc.)
CONTEXT_BAN_RE = re.compile(
    r"""(?ix)
    \b(
        vat|tva|vat\s*no|vat\s*nr|vat\s*id|vat\s*reg|intracommunautaire|
        fiscal\s*code|tax\s*id|tax\s*no|tin|ein|ssn|nid|
        reg(\.|istration)?\s*(no|nr|number)|company\s*reg|crn|brn|roc|uen|uen\s*no|
        iban|swift|bic|bank\s*account|acct\s*no|account\s*no|
        gdpr|regolamento\s*(ue|eu)|ue\s*n\.?|directive|2016\/679|
        ust\-?id|steuernummer|handelsreg|hrb|uid\-?nr|mwst|
        siren|siret|rcs|
        p\.?\s*iva|partita\s*iva|codice\s*fiscale|cf\b|
        nif\b|cif\b|nie\b|nipc|nif\s*pt|
        cui\b|ro[\s\-]?\d{2,}|idno|idnp|
        company\s*no|companies\s*house|utr|vrn|pps\s*no|
        nip\b|regon|pesel|
        i[čc]o\b|i[čc]\s*dph|dic\b|ič\s*dph|
        oib\b|mbs\b|pi[bp]\b|jmbg|mbr|em[bj]g|nuiss|nuis|
        bulstat|eik\b|
        afm\b|vat\s*cy|
        kvk|kbo|tva|tva\s*be|btw\s*nr|
        єдрпоу|едрпоу|edrpou|unp\b|инн\b|кпп\b|ogrn|ogrnip|okpo|bin\b|iin\b|
        vergi|tckn|vkn|
        ח\.פ|ע\.מ|מספר\s*עוסק|
        gstin|pan\b|tan\b|cin\b|udyam|
        cnic|ntn|bin\b|
        uscc|brn|统一社会信用代码|統一編號|ubn\b|
        ssm|roc\s*my|uen|npwp|tin\s*ph|bir\s*no|srn|brn\s*my|dti|
        abn|acn|ird\s*no|gst\s*no|
        fein|ein|itin|ssn|sin|cra|bn\b|
        rfc\b|curp|rut\b|ruc\b|cuit|cuil|dni\b|cedula|nit\b|
        trn\b|cr\b|cac|\bbrn\b|cipc|ck\s*no
    )\b
    """,
    re.IGNORECASE
)

# Things often confused with phones (e.g. GDPR 2016/679, IBAN)
SUSPECT_PATTERN_RE = re.compile(
    r"""(?ix)
    (
        \b20\d{2}/\d{2,4}\b |
        \b[A-Z]{2}\d{2}[A-Z0-9]{11,30}\b
    )
    """
)

def _cleanup_phone_str(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r'(?<!^)\+', '', s)   # internal '+'
    s = re.sub(r'[^\d+]', '', s)     # keep digits and leading '+'
    return s

def _digits_count(s: str) -> int:
    return len(re.sub(r'\D', '', s or ''))

def extrage_numere(text: str):
    if not text:
        return []
    pattern = re.compile(r'\+?\d[\d\-\s\.\(\)\/]{6,}\d')
    results = []
    for m in pattern.finditer(text):
        raw = m.group(0).strip()
        dcnt = _digits_count(raw)
        if dcnt < 7 or dcnt > 15:
            continue
        start, end = m.start(), m.end()
        ctx = text[max(0, start-50): min(len(text), end+50)]
        if CONTEXT_BAN_RE.search(ctx):
            continue
        if SUSPECT_PATTERN_RE.search(ctx) or SUSPECT_PATTERN_RE.search(raw):
            continue
        results.append(_cleanup_phone_str(raw))
    return results

def normalize_with_country_code(phone: str, country: str) -> str:
    digits = re.sub(r'\D', '', phone or '')
    if not digits:
        return ''
    prefix = country_codes.get(country) or ''
    prefix_digits = re.sub(r'\D', '', prefix)
    if prefix_digits and digits.startswith('00' + prefix_digits):
        return digits[2:]
    if prefix_digits and digits.startswith(prefix_digits):
        return digits
    if prefix_digits:
        return prefix_digits + digits.lstrip('0')
    return digits

# === Lazy Chrome driver (prevents crash at import) ===
driver = None

def ensure_driver(consola=None):
    global driver
    if driver is not None:
        return driver
    try:
        options = uc.ChromeOptions()
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--lang=en")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36")
        driver = uc.Chrome(options=options)
        return driver
    except Exception as e:
        msg = f"Nu pot porni browserul: {e}"
        if consola:
            consola.insert(tk.END, "❌ " + msg + "\n"); consola.see(tk.END); consola.update()
        messagebox.showerror("Eroare Chrome", msg)
        return None

# --- Consent helper + panel fallback ---
def accept_google_consent(d):
    try:
        time.sleep(0.5)
        iframes = d.find_elements(By.CSS_SELECTOR, "iframe[src*='consent']")
        if iframes:
            d.switch_to.frame(iframes[0])
        xpaths = [
            "//button[@id='L2AGLb']",
            "//button[normalize-space()='I agree']",
            "//button[contains(., 'Accept all')]",
            "//button[contains(., 'Sunt de acord')]",
            "//button[contains(., 'Acceptă tot')]",
            "//button[contains(., 'Ich stimme zu')]",
            "//button[contains(., 'Aceptar todo')]",
        ]
        for xp in xpaths:
            btns = d.find_elements(By.XPATH, xp)
            if btns:
                try:
                    btns[0].click()
                except WebDriverException:
                    d.execute_script("arguments[0].click();", btns[0])
                break
    except Exception:
        pass
    finally:
        try:
            d.switch_to.default_content()
        except Exception:
            pass

def find_knowledge_panel(d, timeout=7):
    wait = WebDriverWait(d, timeout)
    selectors = {
        "#rhs",
        "div[role='complementary']",
        "div[data-attrid='title']",
    }
    for sel in selectors:
        try:
            return wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
        except TimeoutException:
            continue
    return None

def gaseste_cartela_google(query, consola=None):
    d = ensure_driver(consola=consola)
    if d is None:
        return {"found": False}

    url = f"https://www.google.com/search?q={query.replace(' ', '+')}&hl=en&gl=us&pws=0"
    d.get(url)
    accept_google_consent(d)

    # Detect “unusual traffic”
    try:
        page_text = d.find_element(By.TAG_NAME, "body").text[:4000].lower()
        if "unusual traffic" in page_text or "/sorry/" in d.current_url:
            if consola:
                consola.insert(tk.END, "        ⚠️ Google a cerut verificare (captcha). Sar peste.\n")
                consola.see(tk.END); consola.update()
            return {"found": False}
    except Exception:
        pass

    wait = WebDriverWait(d, 5)

    def _extract_name_from_maps():
        try:
            el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.DUwDvf")))
            name = el.text.strip()
            if name: return name
        except: pass
        try:
            el = d.find_element(By.CSS_SELECTOR, '[role="heading"][aria-level="1"]')
            name = el.text.strip()
            if name: return name
        except: pass
        try:
            el = d.find_element(By.CSS_SELECTOR, 'div[data-attrid="title"] span')
            name = el.text.strip()
            if name: return name
        except: pass
        try:
            meta = d.find_element(By.CSS_SELECTOR, 'meta[property="og:title"]')
            name = (meta.get_attribute("content") or "").strip()
            if name: return name
        except: pass
        return "N/A"

    try:
        panel = find_knowledge_panel(d, timeout=3)
        if not panel:
            if consola:
                consola.insert(tk.END, "        ℹ️ Nu există knowledge panel pentru interogarea asta.\n")
                consola.see(tk.END); consola.update()
            return {"found": False}

        content = panel.text
        numere = extrage_numere(content)

        # --- Closure status strictly by SPANs on panel ---
        closure_status = "Active"
        try:
            if panel.find_elements(By.XPATH, ".//span[normalize-space()='Permanently closed']"):
                closure_status = "Permanently closed"
            elif panel.find_elements(By.XPATH, ".//span[normalize-space()='Temporarily closed']"):
                closure_status = "Temporarily closed"
            else:
                if panel.find_elements(By.XPATH, ".//span[contains(normalize-space(.), 'Permanently closed')]"):
                    closure_status = "Permanently closed"
                elif panel.find_elements(By.XPATH, ".//span[contains(normalize-space(.), 'Temporarily closed')]"):
                    closure_status = "Temporarily closed"
        except Exception:
            pass

        # WEBSITE
        website = None
        try:
            site_elem = panel.find_element(By.XPATH, ".//a[.//span[text()='Site'] or .//span[text()='Website']]")
            website = site_elem.get_attribute("href")
        except Exception:
            try:
                links = panel.find_elements(By.XPATH, ".//a[contains(@href,'http')]")
                for a in links:
                    href = a.get_attribute("href") or ""
                    if "facebook.com" in href: 
                        continue
                    if "google." in href or "support.google" in href or "/maps/" in href:
                        continue
                    website = href
                    break
            except Exception:
                website = None

        # FACEBOOK
        try:
            fb_elem = panel.find_element(By.XPATH, ".//a[contains(@href,'facebook.com')]")
            facebook = fb_elem.get_attribute("href")
        except Exception:
            facebook = None

        # Company name via Maps (if available)
        company_name_found = "N/A"
        opened_maps = False
        try:
            maps_links = panel.find_elements(By.XPATH, ".//a[contains(@href,'/maps/place/')]")
            if maps_links:
                maps_href = maps_links[0].get_attribute("href")
                d.get(maps_href)
                opened_maps = True
                company_name_found = _extract_name_from_maps()

                # also try closure on Maps if still Open
                if closure_status == "Open":
                    try:
                        if d.find_elements(By.XPATH, "//span[normalize-space()='Permanently closed']"):
                            closure_status = "Permanently closed"
                        elif d.find_elements(By.XPATH, "//span[normalize-space()='Temporarily closed']"):
                            closure_status = "Temporarily closed"
                        elif d.find_elements(By.XPATH, "//span[contains(normalize-space(.), 'Permanently closed')]"):
                            closure_status = "Permanently closed"
                        elif d.find_elements(By.XPATH, "//span[contains(normalize-space(.), 'Temporarily closed')]"):
                            closure_status = "Temporarily closed"
                    except Exception:
                        pass
        except Exception:
            pass

        if company_name_found == "N/A" and not opened_maps:
            try:
                el = panel.find_element(By.CSS_SELECTOR, '[role="heading"]')
                company_name_found = el.text.strip() or "N/A"
            except Exception:
                try:
                    el = panel.find_element(By.CSS_SELECTOR, "div[data-attrid='title'] span")
                    company_name_found = el.text.strip() or "N/A"
                except Exception:
                    company_name_found = "N/A"

        if opened_maps:
            d.back()
            accept_google_consent(d)
            try:
                find_knowledge_panel(d, timeout=8)
            except Exception:
                pass

        return {
            "found": True,
            "site": website,
            "facebook": facebook,
            "phones": numere,
            "company_name_found": company_name_found,
            "closure_status": closure_status
        }
    except TimeoutException:
        if consola:
            consola.insert(tk.END, "⏳ Timeout la găsirea panoului.\n"); consola.see(tk.END); consola.update()
        return {"found": False}
    except Exception as e:
        if consola:
            consola.insert(tk.END, f"❌ Eroare la citirea panelului: {type(e).__name__}\n")
            consola.see(tk.END); consola.update()
        return {"found": False}

def extrage_numere_de_pe_pagina(url, consola=None):
    d = ensure_driver(consola=consola)
    if d is None:
        return []
    try:
        d.get(url)
        time.sleep(random.uniform(2, 5))
        d.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(random.uniform(1, 2))

        html = d.page_source
        try:
            text = d.find_element(By.TAG_NAME, "body").text
        except Exception:
            text = ''

        nums = set()
        for m in extrage_numere(text):
            nums.add(_cleanup_phone_str(m))
        for m in PHONE_TEL_RE.findall(html):
            candidate = _cleanup_phone_str(m)
            if 7 <= _digits_count(candidate) <= 15 and not SUSPECT_PATTERN_RE.search(candidate):
                nums.add(candidate)

        return list(nums)
    except Exception as e:
        if consola:
            consola.insert(tk.END, f"   ❌ Page parse error: {e}\n"); consola.see(tk.END); consola.update()
        return []

def save_dataframe_safely(df: pd.DataFrame, default_name="rezultate_companii.xlsx", consola=None):
    def _log(msg):
        if consola is not None:
            consola.insert(tk.END, msg + "\n"); consola.see(tk.END); consola.update()
    try:
        with pd.ExcelWriter(default_name, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        _log(f"💾 Saved: {os.path.abspath(default_name)}")
        return os.path.abspath(default_name)
    except PermissionError:
        _log("⚠️ File is open/locked. Trying with a timestamped filename...")
    except Exception as e:
        _log(f"⚠️ Could not save to default name: {e}. Trying timestamped filename...")

    name, ext = os.path.splitext(default_name)
    ts_name = f"{name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
    try:
        with pd.ExcelWriter(ts_name, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        _log(f"💾 Saved: {os.path.abspath(ts_name)}")
        return os.path.abspath(ts_name)
    except Exception as e:
        _log(f"⚠️ Could not save to timestamped name: {e}")

    _log("📁 Opening Save As… dialog.")
    path = filedialog.asksaveasfilename(
        title="Save results as...",
        defaultextension=".xlsx",
        initialfile=default_name,
        filetypes=[("Excel files", "*.xlsx")],
    )
    if not path:
        _log("🛑 Save cancelled by user.")
        return None
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        _log(f"💾 Saved: {os.path.abspath(path)}")
        return os.path.abspath(path)
    except Exception as e:
        _log(f"❌ Final save failed: {e}")
        return None

def interfata():
    root = tk.Tk()
    root.title("Google Business Card Checker")

    frame = tk.Frame(root)
    frame.pack(pady=10)
    filepath_var = tk.StringVar()
    stop_flag = tk.BooleanVar(value=False)

    if not country_codes:
        messagebox.showerror("Eroare",
            "Nu pot încărca all_country_phone_codes.json.\n"
            "Verifică fișierul și repornește aplicația.")

    global consola
    consola = scrolledtext.ScrolledText(root, width=120, height=30)
    consola.pack(padx=10, pady=10)

    def incarca_fisier():
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            filepath_var.set(filepath)
            consola.insert(tk.END, f"Selected file: {filepath}\n")
            consola.see(tk.END); consola.update()

    def oprire():
        stop_flag.set(True)
        consola.insert(tk.END, "🛑 Stop requested. Finishing current company...\n")
        consola.see(tk.END); consola.update()

    def _pretty_e164(n: str, country: str) -> str:
        prefix_digits = re.sub(r'\D', '', (country_codes.get(country) or ''))
        return ('+' + n) if prefix_digits and n.startswith(prefix_digits) else n

    def proceseaza_fisier(filepath, consola, stop_flag):
        try:
            df = pd.read_excel(filepath)
            rezultate = []

            for idx, (_, row) in enumerate(df.iterrows(), start=1):
                if stop_flag.get():
                    consola.insert(tk.END, '\n🛑 Process was stopped by the user.\n')
                    consola.see(tk.END); consola.update()
                    break

                companie = str(row["Company Name"])
                id_link = str(row.get("Company ID (Link)", ""))
                adresa = str(row.get("Address Line One", "") or "")
                zip_code = str(row.get("ZIP", "") or "")
                city = str(row.get("City", "") or "")
                tara = str(row.get("Country", "") or "")
                phone_col = str(row.get("Phone(s)", "") or "")
                nota = str(row.get("DQP Employee Note", "") or "")

                # Initial phones -> canonical E.164 (no '+'), strip '(x/y)' notes
                phones_initiale = set()
                for p in re.split(r'[;,]', phone_col or ''):
                    p = p.strip()
                    if not p:
                        continue
                    p_curat = re.sub(r'\([^)]*\)', '', p).strip()
                    if not p_curat:
                        continue
                    nrm = normalize_with_country_code(p_curat, tara)
                    if nrm:
                        phones_initiale.add(nrm)

                consola.insert(tk.END, f"\n📦 [{idx}] {companie}:\n")
                consola.see(tk.END); consola.update()

                variante_cautare = [
                    f"{companie} {adresa} {zip_code} {city} {tara}",
                    f"{companie} {adresa}",
                    f"{companie} {city}",
                    f"{companie} {tara}"
                ]
                rezultat_valid = None

                for query in variante_cautare:
                    consola.insert(tk.END, f"   🔍 Searching: {query}\n")
                    consola.see(tk.END); consola.update()
                    rezultat = gaseste_cartela_google(query, consola=consola)
                    if not rezultat.get("found"):
                        continue

                    consola.insert(tk.END, "   ✅ Google card found\n")
                    consola.see(tk.END)
                    consola.update()
                    rezultat_valid = rezultat
                    break

                if not rezultat_valid:
                    consola.insert(tk.END, "   ❌ No matching Google business card found\n")
                    consola.see(tk.END); consola.update()
                    rezultate.append({
                        "Company ID": id_link,
                        "Company Name": companie,
                        "Initial Phones": phone_col,
                        "Matched Company Name": "N/A",
                        "Unique Phones Found": "Google card not found",
                        "Google Phone(s)": "N/A",
                        "Facebook Phone(s)": "N/A",
                        "Website Phone(s)": "N/A",
                        "Closure Status": "N/A"
                    })
                    continue

                # Collect phones from all sources
                toate_numerele = set()

                # Google phones
                google_phones_raw = sorted(set(rezultat_valid.get("phones", [])))
                telefoane_google = ', '.join(google_phones_raw) if google_phones_raw else "N/A"
                toate_numerele.update(
                    n for n in (normalize_with_country_code(p, tara) for p in google_phones_raw)
                    if n
                )

                site = rezultat_valid.get("site")
                facebook = rezultat_valid.get("facebook")
                closure_status = rezultat_valid.get("closure_status", "N/A")

                telefoane_site = "N/A"
                if site:
                    site_nums_raw = extrage_numere_de_pe_pagina(site, consola=consola)
                    if site_nums_raw:
                        telefoane_site = ', '.join(sorted(set(site_nums_raw)))
                        toate_numerele.update(
                            n for n in (normalize_with_country_code(x, tara) for x in site_nums_raw)
                            if n
                        )

                telefoane_fb = "N/A"
                if facebook:
                    fb_nums_raw = extrage_numere_de_pe_pagina(facebook, consola=consola)
                    if fb_nums_raw:
                        telefoane_fb = ', '.join(sorted(set(fb_nums_raw)))
                        toate_numerele.update(
                            n for n in (normalize_with_country_code(x, tara) for x in fb_nums_raw)
                            if n
                        )

                # Numbers from employee note
                numere_nota = set(
                    n for n in (normalize_with_country_code(x, tara) for x in extrage_numere(nota))
                    if n
                )

                # Additional = all found - already present - from note
                numere_adaugate = toate_numerele - phones_initiale - numere_nota
                text_aditional = ', '.join(sorted(_pretty_e164(n, tara) for n in numere_adaugate)) \
                                 if numere_adaugate else "No additional phone found."

                matched_name = (rezultat_valid.get("company_name_found") or "").strip() or "N/A"
                consola.insert(tk.END, f"   🏷️ Matched Name: {matched_name}\n")
                consola.insert(tk.END, f"   🏪 Closure Status: {closure_status}\n")
                consola.insert(tk.END, f"   ➤ Additional phones: {text_aditional}\n")
                consola.see(tk.END); consola.update()

                rezultate.append({
                    "Company ID": id_link,
                    "Company Name": companie,
                    "Initial Phones": phone_col,
                    "Matched Company Name": matched_name,
                    "Unique Phones Found": text_aditional,
                    "Google Phone(s)": telefoane_google,
                    "Facebook Phone(s)": telefoane_fb,
                    "Website Phone(s)": telefoane_site,
                    "Closure Status": closure_status
                })

            rezultat_df = pd.DataFrame(rezultate)
            saved_path = save_dataframe_safely(rezultat_df, "rezultate_companii.xlsx", consola=consola)

            if saved_path:
                consola.insert(tk.END, f"\n✅ Done. Results saved to:\n{saved_path}\n")
                consola.see(tk.END)
                messagebox.showinfo("Complete!", f"Processing complete.\nSaved to:\n{saved_path}")
            else:
                consola.insert(tk.END, "\n❌ Done, but could not save the Excel file.\n")
                consola.see(tk.END)
                messagebox.showerror("Save failed", "Processing finished, but the Excel file could not be saved.\nTry closing any open Excel files and run again.")
        except Exception as e:
            import traceback
            consola.insert(tk.END, "❌ A apărut o eroare:\n" + traceback.format_exc() + "\n")
            consola.see(tk.END); consola.update()
            messagebox.showerror("Eroare", str(e))
        finally:
            consola.insert(tk.END, "ℹ️ Worker thread finished.\n")
            consola.see(tk.END); consola.update()

    def start_procesare():
        if not filepath_var.get():
            messagebox.showerror("Error", "Select an Excel file first.")
            return
        stop_flag.set(False)
        threading.Thread(target=proceseaza_fisier, args=(filepath_var.get(), consola, stop_flag), daemon=True).start()

    def on_close():
        try:
            d = ensure_driver(consola=consola)
            if d is not None:
                d.quit()
        except Exception:
            pass
        root.destroy()

    tk.Button(frame, text="Load Excel File", command=incarca_fisier).pack(side=tk.LEFT, padx=5)
    tk.Button(frame, text="Start", command=start_procesare).pack(side=tk.LEFT, padx=5)
    tk.Button(frame, text="Stop", command=oprire).pack(side=tk.LEFT, padx=5)

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    interfata()
