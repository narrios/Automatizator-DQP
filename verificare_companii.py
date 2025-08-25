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

# === Load country phone codes (e.g., {"Spain": "+34", "Romania": "+40", ...}) ===
with open("all_country_phone_codes.json", "r", encoding="utf-8") as f:
    country_codes = json.load(f)

# === Phone patterns and helpers ===
PHONE_TEXT_RE = re.compile(r'\+?\d[\d\s\-\.\(\)\/]{6,}\d')
PHONE_TEL_RE  = re.compile(r'href=["\']\s*tel:([+0-9\-\s\.\(\)\/]+)["\']', re.I)

def _cleanup_phone_str(s: str) -> str:
    """Keep a leading '+' if present, remove all other non-digits/non-leading plus."""
    s = (s or "").strip()
    s = re.sub(r'(?<!^)\+', '', s)       # internal '+' -> remove
    s = re.sub(r'[^\d+]', '', s)         # keep only digits and a possible leading '+'
    return s

def extrage_numere(text: str):
    """Extract phone-like strings from visible text."""
    if not text:
        return []
    return [_cleanup_phone_str(m) for m in PHONE_TEXT_RE.findall(text)]

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

    # handle 00 + countrycode (international)
    if prefix_digits and digits.startswith('00' + prefix_digits):
        return digits[2:]  # drop leading '00'

    # already has country code
    if prefix_digits and digits.startswith(prefix_digits):
        return digits

    # local/lacking country code -> prepend country code, drop trunk '0'
    if prefix_digits:
        local = digits.lstrip('0')
        return prefix_digits + local

    # fallback when no country prefix known
    return digits

# === Browser setup ===
options = uc.ChromeOptions()
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--lang=en")
options.add_argument("--window-size=1920,1080")
driver = uc.Chrome(options=options)

def gaseste_cartela_google(query):
    driver.get(f"https://www.google.com/search?q={query.replace(' ', '+')}")
    wait = WebDriverWait(driver, 10)

    def _extract_name_from_maps():
        # We are on Google Maps business profile page
        try:
            el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.DUwDvf")))
            name = el.text.strip()
            if name:
                return name
        except:
            pass
        try:
            el = driver.find_element(By.CSS_SELECTOR, '[role="heading"][aria-level="1"]')
            name = el.text.strip()
            if name:
                return name
        except:
            pass
        try:
            el = driver.find_element(By.CSS_SELECTOR, 'div[data-attrid="title"] span')
            name = el.text.strip()
            if name:
                return name
        except:
            pass
        try:
            meta = driver.find_element(By.CSS_SELECTOR, 'meta[property="og:title"]')
            name = meta.get_attribute("content").strip()
            if name:
                return name
        except:
            pass
        return "N/A"

    try:
        # Wait for the right-side panel (knowledge panel / business card)
        panel = wait.until(EC.presence_of_element_located((By.ID, "rhs")))
        content = panel.text
        numere = extrage_numere(content)

        # WEBSITE — try "Site" span first, then fallbacks
        website = None
        try:
            site_elem = panel.find_element(By.XPATH, ".//a[.//span[text()='Site']]")
            website = site_elem.get_attribute("href")
        except:
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
            except:
                website = None

        # FACEBOOK
        try:
            fb_elem = panel.find_element(By.XPATH, ".//a[contains(@href,'facebook.com')]")
            facebook = fb_elem.get_attribute("href")
        except:
            facebook = None

        # ==== COMPANY NAME: try opening the Maps profile and extract from there ====
        company_name_found = "N/A"
        opened_maps = False

        try:
            maps_links = panel.find_elements(By.XPATH, ".//a[contains(@href,'/maps/place/')]")
            if maps_links:
                maps_href = maps_links[0].get_attribute("href")
                driver.get(maps_href)   # open in same tab
                opened_maps = True
                company_name_found = _extract_name_from_maps()
        except:
            pass

        # If not on Maps, fallback to panel headings
        if company_name_found == "N/A" and not opened_maps:
            try:
                el = panel.find_element(By.CSS_SELECTOR, '[role="heading"]')
                company_name_found = el.text.strip() or "N/A"
            except:
                try:
                    el = panel.find_element(By.CSS_SELECTOR, "div[data-attrid='title'] span")
                    company_name_found = el.text.strip() or "N/A"
                except:
                    company_name_found = "N/A"

        # If we opened Maps, go back to SERP to keep the flow consistent
        if opened_maps:
            driver.back()
            try:
                wait.until(EC.presence_of_element_located((By.ID, "rhs")))
            except:
                pass

        return {
            "found": True,
            "site": website,
            "facebook": facebook,
            "phones": numere,
            "company_name_found": company_name_found
        }
    except:
        return {"found": False}

def extrage_numere_de_pe_pagina(url):
    """Extract phone numbers from a page (visible text + tel: links)."""
    try:
        driver.get(url)
        # small scroll to load footer/lazy content
        time.sleep(1.5)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1.5)

        html = driver.page_source
        try:
            text = driver.find_element(By.TAG_NAME, "body").text
        except:
            text = ''

        nums = set()
        # 1) from visible text
        for m in PHONE_TEXT_RE.findall(text):
            nums.add(_cleanup_phone_str(m))
        # 2) from tel: hrefs
        for m in PHONE_TEL_RE.findall(html):
            nums.add(_cleanup_phone_str(m))

        return list(nums)
    except:
        return []

def save_dataframe_safely(df: pd.DataFrame, default_name="rezultate_companii.xlsx", consola=None):
    """
    Încearcă să salveze df în `default_name`.
    Dacă fișierul este blocat, salvează cu timestamp.
    Dacă tot nu reușește, deschide un Save As… pentru a alege manual.
    Returnează calea finală salvată sau None dacă a eșuat.
    """
    def _log(msg):
        if consola is not None:
            consola.insert(tk.END, msg + "\n")
            consola.see(tk.END)
            consola.update()

    # 1) încercare pe numele implicit, în folderul curent
    try:
        with pd.ExcelWriter(default_name, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        _log(f"💾 Saved: {os.path.abspath(default_name)}")
        return os.path.abspath(default_name)
    except PermissionError:
        _log("⚠️ File is open/locked. Trying with a timestamped filename...")
    except Exception as e:
        _log(f"⚠️ Could not save to default name: {e}. Trying timestamped filename...")

    # 2) nume cu timestamp
    name, ext = os.path.splitext(default_name)
    ts_name = f"{name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
    try:
        with pd.ExcelWriter(ts_name, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        _log(f"💾 Saved: {os.path.abspath(ts_name)}")
        return os.path.abspath(ts_name)
    except Exception as e:
        _log(f"⚠️ Could not save to timestamped name: {e}")

    # 3) Save As… (alegi manual)
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
        """For display only: add '+' if number starts with the country code."""
        prefix_digits = re.sub(r'\D', '', (country_codes.get(country) or ''))
        return ('+' + n) if prefix_digits and n.startswith(prefix_digits) else n

    def proceseaza_fisier(filepath, consola, stop_flag):
        df = pd.read_excel(filepath)
        rezultate = []

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

            # Build initial phone set in canonical E.164 (no '+')
            phones_initiale = set(
                n for n in (normalize_with_country_code(p.strip(), tara) for p in re.split(r'[;,]', phone_col or ''))
                if n
            )

            consola.insert(tk.END, f"\n📦 [{idx}] {companie}:\n")
            consola.see(tk.END)
            consola.update()

            variante_cautare = [
                f"{companie} {adresa} {zip_code} {city} {tara}",
                f"{companie} {adresa}",
                f"{companie} {city}",
                f"{companie} {tara}"
            ]
            rezultat_valid = None

            for query in variante_cautare:
                consola.insert(tk.END, f"   🔍 Searching: {query}\n")
                consola.see(tk.END)
                consola.update()
                rezultat = gaseste_cartela_google(query)
                if not rezultat.get("found"):
                    continue

                google_norm = set(
                    n for n in (normalize_with_country_code(p, tara) for p in rezultat.get("phones", []))
                    if n
                )

                if phones_initiale.intersection(google_norm):
                    consola.insert(tk.END, "   ✅ Phone matched with Google card\n")
                    consola.see(tk.END)
                    consola.update()
                    rezultat_valid = rezultat
                    break
                else:
                    consola.insert(tk.END, "   ⚠️ Google card found but phone did not match\n")
                    consola.see(tk.END)
                    consola.update()

            if not rezultat_valid:
                consola.insert(tk.END, "   ❌ No matching Google business card found\n")
                consola.see(tk.END)
                consola.update()
                rezultate.append({
                    "Company ID": id_link,
                    "Company Name": companie,
                    "Initial Phones": phone_col,
                    "Matched Company Name": "N/A",
                    "Unique Phones Found": "Google card not found",
                    "Google Phone(s)": "N/A",
                    "Facebook Phone(s)": "N/A",
                    "Website Phone(s)": "N/A"
                })
                continue

            # Collect phones from all sources (Google card + site + Facebook), normalized for dedup
            toate_numerele = set()

            # Google phones (raw for display, normalized for dedup)
            google_phones_raw = sorted(set(rezultat_valid.get("phones", [])))
            telefoane_google = ', '.join(google_phones_raw) if google_phones_raw else "N/A"
            toate_numerele.update(
                n for n in (normalize_with_country_code(p, tara) for p in google_phones_raw)
                if n
            )

            site = rezultat_valid.get("site")
            facebook = rezultat_valid.get("facebook")

            telefoane_site = "N/A"
            if site:
                site_nums_raw = extrage_numere_de_pe_pagina(site)
                if site_nums_raw:
                    telefoane_site = ', '.join(sorted(set(site_nums_raw)))
                    toate_numerele.update(
                        n for n in (normalize_with_country_code(x, tara) for x in site_nums_raw)
                        if n
                    )

            telefoane_fb = "N/A"
            if facebook:
                fb_nums_raw = extrage_numere_de_pe_pagina(facebook)
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
            text_aditional = ', '.join(sorted(_pretty_e164(n, tara) for n in numere_adaugate)) \
                             if numere_adaugate else "No additional phone found."

            matched_name = (rezultat_valid.get("company_name_found") or "").strip() or "N/A"
            consola.insert(tk.END, f"   🏷️ Matched Name: {matched_name}\n")
            consola.insert(tk.END, f"   ➤ Additional phones: {text_aditional}\n")
            consola.see(tk.END)
            consola.update()

            rezultate.append({
                "Company ID": id_link,
                "Company Name": companie,
                "Initial Phones": phone_col,
                "Matched Company Name": matched_name,
                "Unique Phones Found": text_aditional,
                "Google Phone(s)": telefoane_google,
                "Facebook Phone(s)": telefoane_fb,
                "Website Phone(s)": telefoane_site
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

    # === UI Buttons ===
    tk.Button(frame, text="Load Excel File", command=incarca_fisier).pack(side=tk.LEFT, padx=5)
    tk.Button(frame, text="Start", command=lambda: (
        stop_flag.set(False),
        threading.Thread(target=proceseaza_fisier, args=(filepath_var.get(), consola, stop_flag), daemon=True).start()
    )).pack(side=tk.LEFT, padx=5)
    tk.Button(frame, text="Stop", command=oprire).pack(side=tk.LEFT, padx=5)

    # === Console ===
    global consola
    consola = scrolledtext.ScrolledText(root, width=120, height=30)
    consola.pack(padx=10, pady=10)

    root.mainloop()
    driver.quit()

if __name__ == "__main__":
    interfata()
