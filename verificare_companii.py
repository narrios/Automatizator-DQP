import json
import os
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

# === Load country phone codes safely ===
try:
    with open("all_country_phone_codes.json", "r", encoding="utf-8") as f:
        country_codes = json.load(f)
except Exception as e:
    country_codes = {}
    # vom afișa eroarea în UI la start

# === Phone patterns and helpers ===
PHONE_TEXT_RE = re.compile(r'\+?\d[\d\s\-\.\(\)\/]{6,}\d')
PHONE_TEL_RE  = re.compile(r'href=["\']\s*tel:([+0-9\-\s\.\(\)\/]+)["\']', re.I)

# --- Ban words around numbers that are NOT phones (global tax IDs, reg. numbers, banking, GDPR etc.)
CONTEXT_BAN_RE = re.compile(
    r"""(?ix)
    \b(
        # EU / generic
        vat|tva|vat\s*no|vat\s*nr|vat\s*id|vat\s*reg|intracommunautaire|
        fiscal\s*code|tax\s*id|tax\s*no|tin|ein|ssn|nid|
        reg(\.|istration)?\s*(no|nr|number)|company\s*reg|crn|brn|roc|uen|uen\s*no|
        # Banking
        iban|swift|bic|bank\s*account|acct\s*no|account\s*no|
        # GDPR / laws
        gdpr|regolamento\s*(ue|eu)|ue\s*n\.?|directive|2016\/679|
        # DE/AT/CH
        ust\-?id|steuernummer|handelsreg|hrb|uid\-?nr|mwst|
        # FR
        siren|siret|rcs|
        # IT
        p\.?\s*iva|partita\s*iva|codice\s*fiscale|cf\b|
        # ES/PT
        nif\b|cif\b|nie\b|nipc|nif\s*pt|
        # RO/MO
        cui\b|ro[\s\-]?\d{2,}|idno|idnp|
        # UK/IE
        company\s*no|companies\s*house|utr|vrn|pps\s*no|
        # PL
        nip\b|regon|pesel|
        # CZ/SK
        i[čc]o\b|i[čc]\s*dph|dic\b|ič\s*dph|
        # HR/SI/RS/BA/ME/MK/AL
        oib\b|mbs\b|pi[bp]\b|jmbg|mbr|em[bj]g|nuiss|nuis|
        # BG
        bulstat|eik\b|
        # GR/CY
        afm\b|vat\s*cy|
        # NL/BE/LU
        kvk|kbo|tva|tva\s*be|btw\s*nr|
        # UA/BY/RU/KZ
        єдрпоу|едрпоу|edrpou|unp\b|инн\b|кпп\b|ogrn|ogrnip|okpo|bin\b|iin\b|
        # TR
        vergi|tckn|vkn|
        # IL
        ח\.פ|ע\.מ|מספר\s*עוסק|
        # IN
        gstin|pan\b|tan\b|cin\b|udyam|
        # PK/BD/LK/NP
        cnic|ntn|bin\b|
        # CN/HK/TW
        uscc|brn|统一社会信用代码|統一編號|ubn\b|
        # SG/MY/PH/ID/TH/VN
        ssm|roc\s*my|uen|npwp|siret\s*id|tin\s*ph|bir\s*no|srn|brn\s*my|dti|
        # AU/NZ
        abn|acn|ird\s*no|gst\s*no|
        # US/CA
        fein|ein|itin|ssn|sin|cra|bn\b|
        # MX/Central/South America
        rfc\b|curp|rut\b|ruc\b|cuit|cuil|dni\b|cedula|nit\b|
        # Africa / Middle East (examples)
        trn\b|cr\b|cac|\bbrn\b|cipc|ck\s*no
    )\b
    """,
    re.IGNORECASE
)

# Some numeric patterns often confused with phones (e.g., GDPR 2016/679, IBAN)
SUSPECT_PATTERN_RE = re.compile(
    r"""(?ix)
    (
        \b20\d{2}/\d{2,4}\b |           # e.g., 2016/679 (GDPR), 2023/1234
        \b[A-Z]{2}\d{2}[A-Z0-9]{11,30}\b # IBAN (approx.)
    )
    """
)

def accept_google_consent(d):
    """Închide dialogul de consimțământ Google (cookies/terms), dacă apare."""
    try:
        time.sleep(0.5)
        # dacă e într-un iframe
        iframes = d.find_elements(By.CSS_SELECTOR, "iframe[src*='consent']")
        if iframes:
            d.switch_to.frame(iframes[0])
        # caută butoane cu texte uzuale în mai multe limbi
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
                except:
                    d.execute_script("arguments[0].click();", btns[0])
                break
    except Exception:
        pass
    finally:
        try:
            d.switch_to.default_content()
        except Exception:
            pass


def is_captcha_page(d) -> bool:
    """Detectează pagina /sorry sau mesajul 'unusual traffic' (Google CAPTCHA)."""
    try:
        if "/sorry/" in d.current_url.lower():
            return True
        html = d.page_source.lower()
        if "our systems have detected unusual traffic" in html:
            return True
        return False
    except Exception:
        return False

def accept_google_consent(d):
    """Închide dialogul de consimțământ Google (cookies/terms), dacă apare."""
    try:
        time.sleep(0.5)
        # dacă e într-un iframe de consimțământ
        iframes = d.find_elements(By.CSS_SELECTOR, "iframe[src*='consent']")
        if iframes:
            d.switch_to.frame(iframes[0])

        # butoane uzuale pe diverse limbi
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
                except:
                    d.execute_script("arguments[0].click();", btns[0])
                break
    except Exception:
        pass
    finally:
        try:
            d.switch_to.default_content()
        except Exception:
            pass

def find_knowledge_panel(d, timeout=6):
    """Caută și returnează elementul knowledge panel din dreapta (SERP)."""
    wait = WebDriverWait(d, timeout)
    selectors = [
        "#rhs",                       # layout clasic
        "div[role='complementary']",  # unele layout-uri
        "div[data-attrid='title']"    # fallback pe titlu
    ]
    for sel in selectors:
        try:
            return wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
        except:
            continue
    return None

def switch_to_last_window(d):
    """Comută pe ultimul tab/fereastră deschisă în driver."""
    try:
        handles = d.window_handles
        if handles:
            d.switch_to.window(handles[-1])
    except Exception:
        pass

def _extract_name_from_maps(d, wait, consola=None):
        try:
            # 1) H1 clasic
            el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.DUwDvf")))
            name = el.text.strip()
            if name:
                return name
        except:
            pass
        try:
            el = d.find_element(By.CSS_SELECTOR, '[role="heading"][aria-level="1"]')
            name = el.text.strip()
            if name:
                return name
        except:
            pass
        try:
            el = d.find_element(By.CSS_SELECTOR, 'div[data-attrid="title"] span')
            name = el.text.strip()
            if name:
                return name
        except:
            pass
        try:
            meta = d.find_element(By.CSS_SELECTOR, 'meta[property="og:title"]')
            name = meta.get_attribute("content").strip()
            name = re.sub(r"\s*[-–]\s*Google Maps\s*$", "", name)
            if name:
                return name
        except:
            pass
        # Fallback pe <title>
        try:
            title = d.title.strip()
            title = re.sub(r"\s*[-–]\s*Google Maps\s*$", "", title)
            if title:
                return title
        except:
            pass
        return "N/A"

def _cleanup_phone_str(s: str) -> str:
    """Keep a leading '+' if present, remove all other non-digits/non-leading plus."""
    s = (s or "").strip()
    s = re.sub(r'(?<!^)\+', '', s)       # internal '+' -> remove
    s = re.sub(r'[^\d+]', '', s)         # keep only digits and a possible leading '+'
    return s

def _digits_count(s: str) -> int:
    return len(re.sub(r'\D', '', s or ''))

def extrage_numere(text: str):
    """
    Extract phone-like strings from visible text, excluding
    tax IDs / banking refs / GDPR refs in the nearby context.
    Enforces 7..15 digits and removes IBAN/GDPR-like patterns.
    """
    if not text:
        return []

    pattern = re.compile(r'\+?\d[\d\-\s\.\(\)\/]{6,}\d')
    results = []

    for m in pattern.finditer(text):
        raw = m.group(0).strip()
        dcnt = _digits_count(raw)
        if dcnt < 7 or dcnt > 15:
            continue

        # context ~ 50 chars around the match
        start, end = m.start(), m.end()
        left = max(0, start - 50)
        right = min(len(text), end + 50)
        ctx = text[left:right]

        if CONTEXT_BAN_RE.search(ctx):
            continue
        if SUSPECT_PATTERN_RE.search(ctx) or SUSPECT_PATTERN_RE.search(raw):
            continue

        results.append(_cleanup_phone_str(raw))

    return results

def normalize_with_country_code(phone: str, country: str) -> str:
    """
    Return canonical phone for comparison/dedup in E.164 WITHOUT '+'.
    Examples:
      '+34 982 25 42 87' -> '34982254287'
      '982 25 42 87' (ES) -> '34982254287'
    """
    digits = re.sub(r'\D', '', phone or '')
    if not digits:
        return ''

    prefix = country_codes.get(country) or ''
    prefix_digits = re.sub(r'\D', '', prefix)

    # 00 + country code
    if prefix_digits and digits.startswith('00' + prefix_digits):
        return digits[2:]  # drop leading '00'

    # already has country code
    if prefix_digits and digits.startswith(prefix_digits):
        return digits

    # local/lacking country code -> prepend, drop trunk '0'
    if prefix_digits:
        local = digits.lstrip('0')
        return prefix_digits + local

    return digits

# === Lazy Chrome driver (prevents crash at import) ===
driver = None  # will be created on-demand

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
        #options.add_argument("--headless")
        driver = uc.Chrome(options=options)
        return driver
    except Exception as e:
        msg = f"The browser can not be opened: {e}"
        if consola:
            consola.insert(tk.END, "❌ " + msg + "\n")
            consola.see(tk.END)
            consola.update()
        messagebox.showerror("Chrome error", msg)
        return None

def gaseste_cartela_google(query, consola=None):
    global driver

    d = ensure_driver(consola=consola)
    if d is None:
        return {"found": False}

    url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
    d.get(url)
    accept_google_consent(d)

    # captcha handling
    if is_captcha_page(d):
        if consola:
            consola.insert(tk.END, "⚠️ Captcha detected. Restarting browser...\n")
            consola.see(tk.END); consola.update()
        try:
            d.quit()
        except:
            pass
        driver = None
        d = ensure_driver(consola=consola)
        if d is None:
            return {"found": False, "captcha": True}
        d.get(url)
        if is_captcha_page(d):
            if consola:
                consola.insert(tk.END, "❌ Captcha still present after restart. Skipping.\n")
                consola.see(tk.END); consola.update()
            return {"found": False, "captcha": True}

    wait = WebDriverWait(d, 6)

    try:
        panel = find_knowledge_panel(d, timeout=3)
        if not panel:
            if consola:
                consola.insert(tk.END, "ℹ️ Google card does not exist.\n")
                consola.see(tk.END); consola.update()
            return {"found": False}

        content = panel.text
        numere = extrage_numere(content)

        # Closure status
        closure_status = "Active"
        try:
            if panel.find_elements(By.XPATH, ".//span[normalize-space()='Permanently closed']"):
                closure_status = "Permanently closed"
            elif panel.find_elements(By.XPATH, ".//span[normalize-space()='Temporarily closed']"):
                closure_status = "Temporarily closed"
        except:
            pass

        # Website
        website = None
        try:
            site_elem = panel.find_element(By.XPATH, ".//a[.//span[text()='Site'] or .//span[text()='Website']]")
            website = site_elem.get_attribute("href")
        except:
            try:
                links = panel.find_elements(By.XPATH, ".//a[contains(@href,'http')]")
                for a in links:
                    href = a.get_attribute("href") or ""
                    if "facebook.com" in href: continue
                    if "google." in href or "support.google" in href or "/maps/" in href: continue
                    website = href; break
            except:
                website = None

        # Facebook
        try:
            fb_elem = panel.find_element(By.XPATH, ".//a[contains(@href,'facebook.com')]")
            facebook = fb_elem.get_attribute("href")
        except:
            facebook = None

        # Google Maps profile
        company_name_found = "N/A"
        try:
            maps_links = panel.find_elements(By.XPATH, ".//a[contains(@href,'/maps/place/')]")
            if maps_links:
                maps_href = maps_links[0].get_attribute("href")
                d.get(maps_href)
                switch_to_last_window(d)
                accept_google_consent(d)

                WebDriverWait(d, 6).until(
                    lambda drv: "/maps/place" in drv.current_url or 
                                drv.find_elements(By.CSS_SELECTOR, "meta[property='og:title']"))

                company_name_found = _extract_name_from_maps(d, wait, consola=consola)

                # verifică și statusul pe Maps
                if closure_status == "Active":
                    try:
                        if d.find_elements(By.XPATH, "//span[normalize-space()='Permanently closed']"):
                            closure_status = "Permanently closed"
                        elif d.find_elements(By.XPATH, "//span[normalize-space()='Temporarily closed']"):
                            closure_status = "Temporarily closed"
                    except:
                        pass
        except Exception as e:
            if consola:
                consola.insert(tk.END, f"   ℹ️ Maps open error: {e}\n")
                consola.see(tk.END); consola.update()

        return {
            "found": True,
            "site": website,
            "facebook": facebook,
            "phones": numere,
            "company_name_found": company_name_found,
            "closure_status": closure_status
        }

    except Exception as e:
        if consola:
            consola.insert(tk.END, f"❌ Error reading the Google Card: {type(e).__name__}\n")
            consola.see(tk.END); consola.update()
        return {"found": False}

def extrage_numere_de_pe_pagina(url, consola=None):
    d = ensure_driver(consola=consola)
    if d is None:
        return []
    try:
        d.get(url)
        # small scroll to load footer/lazy content
        time.sleep(1.5)
        d.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1.0)

        html = d.page_source
        try:
            text = d.find_element(By.TAG_NAME, "body").text
        except:
            text = ''

        nums = set()
        # 1) from visible text (context-filtered)
        for m in extrage_numere(text):
            nums.add(_cleanup_phone_str(m))
        # 2) from tel: hrefs (length/suspect filters only)
        for m in PHONE_TEL_RE.findall(html):
            candidate = _cleanup_phone_str(m)
            if 7 <= _digits_count(candidate) <= 15 and not SUSPECT_PATTERN_RE.search(candidate):
                nums.add(candidate)

        return list(nums)
    except Exception as e:
        if consola:
            consola.insert(tk.END, f"   ❌ Page parse error: {e}\n")
            consola.see(tk.END)
            consola.update()
        return []

def save_dataframe_safely(df: pd.DataFrame, default_name="rezultate_companii.xlsx", consola=None):
    """
    Save df safely:
      1) try default name;
      2) if locked, save with timestamp;
      3) if still failing, open Save As…
    """
    def _log(msg):
        if consola is not None:
            consola.insert(tk.END, msg + "\n")
            consola.see(tk.END)
            consola.update()

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

    # if phone codes failed to load, show an error but allow UI to open
    if not country_codes:
        messagebox.showerror("Eroare",
            "all_country_phone_codes.json can not be loaded.\n"
            "Check the file and restart the application.")
    # Console
    global consola
    consola = scrolledtext.ScrolledText(root, width=120, height=30)
    consola.pack(padx=10, pady=10)

    def incarca_fisier():
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filepath:
            filepath_var.set(filepath)
            consola.insert(tk.END, f"Selected file: {filepath}\n")
            consola.see(tk.END)
            consola.update()

    def oprire():
        stop_flag.set(True)
        consola.insert(tk.END, "🛑 Stop requested. Finishing current company...\n")
        consola.see(tk.END)
        consola.update()

    def _pretty_e164(n: str, country: str) -> str:
        prefix_digits = re.sub(r'\D', '', (country_codes.get(country) or ''))
        return ('+' + n) if prefix_digits and n.startswith(prefix_digits) else n

    def proceseaza_fisier(filepath, consola, stop_flag):
        try:
            df = pd.read_excel(filepath)
            rezultate = []

            backup_interval = 5   # salvează la fiecare 5 companii
            counter = 0

            for idx, (_, row) in enumerate(df.iterrows(), start=1):
                if stop_flag.get():
                    consola.insert(tk.END, '\n🛑 Process was stopped by the user.\n')
                    consola.see(tk.END)
                    consola.update()
                    break

                companie = str(row["Company Name"])
                id_link = str(row.get("Company ID (Link)", ""))
                adresa = str(row.get("Address Line One", "") or "")
                zip_code = str(row.get("ZIP", "") or "")
                city = str(row.get("City", "") or "")
                tara = str(row.get("Country", "") or "")
                phone_col = str(row.get("Phone(s)", "") or "")
                nota = str(row.get("DQP Employee Note", "") or "")

                # Initial phones -> canonical E.164 (no '+'), strip any '(x/y)' suffixes
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
                consola.see(tk.END)
                consola.update()

                variante_cautare = [
                    f"{companie} {adresa} {zip_code} {city} {tara}",
                    f"{companie} {city} {tara}",
                ]
                rezultat_valid = None

                for query in variante_cautare:
                    consola.insert(tk.END, f"   🔍 Searching: {query}\n")
                    consola.see(tk.END)
                    consola.update()
                    rezultat = gaseste_cartela_google(query, consola=consola)
                    if not rezultat.get("found"):
                        continue

                    google_norm = set(
                        n for n in (normalize_with_country_code(p, tara) for p in rezultat.get("phones", []))
                        if n
                    )

                    rezultat = gaseste_cartela_google(query, consola=consola)
                    if not rezultat.get("found"):
                        continue

                    consola.insert(tk.END, "   ✅ Google card found\n")
                    consola.see(tk.END); consola.update()
                    rezultat_valid = rezultat
                    break

                if not rezultat_valid:
                    consola.insert(tk.END, "   ❌ No matching Google business card found\n")
                    consola.see(tk.END)
                    consola.update()
                    rezultate.append({
                        "Company ID": id_link,
                        "Company Name": companie,
                        "Initial Phones": phone_col,
                        "DQP Employee Note": nota,
                        "Matched Company Name": "N/A",
                        "Unique Phones Found": "Google card not found",
                        "Google Phone(s)": "N/A",
                        "Facebook Phone(s)": "N/A",
                        "Website Phone(s)": "N/A",
                        "Closure Status": "N/A"
                    })
                    continue

                # Collect phones from all sources (Google card + site + Facebook), normalized for dedup
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

                # Numbers from employee note (normalize too)
                numere_nota = set(
                    n for n in (normalize_with_country_code(x, tara) for x in extrage_numere(nota))
                    if n
                )

                # Additional = all found - already present - from note
                numere_adaugate = toate_numerele - phones_initiale - numere_nota
                text_aditional = ', '.join(sorted(
                    ('+' + n) if re.sub(r'\D', '', (country_codes.get(tara) or '')) and n.startswith(re.sub(r'\D','',(country_codes.get(tara) or ''))) else n
                    for n in numere_adaugate
                )) if numere_adaugate else "No additional phone found."

                matched_name = (rezultat_valid.get("company_name_found") or "").strip() or "N/A"
                consola.insert(tk.END, f"   🏷️ Matched Name: {matched_name}\n")
                consola.insert(tk.END, f"   🏪 Closure Status: {closure_status}\n")
                consola.insert(tk.END, f"   ➤ Additional phones: {text_aditional}\n")
                consola.see(tk.END)
                consola.update()

                rezultate.append({
                    "Company ID": id_link,
                    "Company Name": companie,
                    "Initial Phones": phone_col,
                    "DQP Employee Note": nota,
                    "Matched Company Name": matched_name,
                    "Unique Phones Found": text_aditional,
                    "Google Phone(s)": telefoane_google,
                    "Facebook Phone(s)": telefoane_fb,
                    "Website Phone(s)": telefoane_site,
                    "Closure Status": closure_status
                })

                counter += 1
                if counter % backup_interval == 0:
                    # Suprascrie backup-ul la fiecare 5 companii
                    rezultat_df = pd.DataFrame(rezultate)
                    rezultat_df.to_excel("rezultate_companii_backup.xlsx", index=False)
                    consola.insert(tk.END, f"💾 Backup has been overwrited after {counter} companies.\n")
                    consola.see(tk.END)
                    consola.update()

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
            consola.insert(tk.END, "❌ An error appered:\n" + traceback.format_exc() + "\n")
            consola.see(tk.END)
            consola.update()
            messagebox.showerror("Error", str(e))
        finally:
            consola.insert(tk.END, "ℹ️ Worker thread finished.\n")
            consola.see(tk.END)
            consola.update()

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
        except:
            pass
        root.destroy()

    # UI buttons
    tk.Button(frame, text="Load Excel File", command=incarca_fisier).pack(side=tk.LEFT, padx=5)
    tk.Button(frame, text="Start", command=start_procesare).pack(side=tk.LEFT, padx=5)
    tk.Button(frame, text="Stop", command=oprire).pack(side=tk.LEFT, padx=5)

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()

if __name__ == "__main__":
    interfata()