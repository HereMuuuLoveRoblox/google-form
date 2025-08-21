# fill_gform_edge.py
# -*- coding: utf-8 -*-
"""
กรอก Google Form อัตโนมัติด้วย Microsoft Edge (Selenium 4) โดยแยกค่าฟอร์มเป็นไฟล์ config
วิธีใช้:
    1) pip install selenium requests
    2) แก้ไฟล์ config (FORM_URL, Q, GRID_TITLES, POOLS, ฯลฯ)
    3) python fill_gform_edge.py
"""

import random
import time
from typing import List, Optional

import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ======= เลือก config ที่ต้องการ =======
import gform_config_ssru as config
# ถ้าทำฟอร์มใหม่: ก็อปไฟล์ config นี้เป็น gform_config_myform.py แล้วเปลี่ยนเป็น:
# import gform_config_myform as config

# ---------------- Utilities ----------------
def resolve_viewform_url(url: str) -> str:
    try:
        r = requests.get(url, timeout=20, allow_redirects=True)
        r.raise_for_status()
        final = r.url
        if "/viewform" not in final:
            final = final.rstrip("/") + "/viewform"
        return final
    except Exception:
        if "docs.google.com/forms" in url and "/viewform" not in url:
            return url.rstrip("/") + "/viewform"
        return url

def make_driver():
    opts = EdgeOptions()
    # opts.add_argument("--headless=new")  # ถ้าต้องการรันแบบไม่โชว์หน้าต่าง
    opts.add_argument("--start-maximized")
    if config.USE_EDGE_PROFILE:
        opts.add_argument(f"--user-data-dir={config.EDGE_USER_DATA_DIR}")
        opts.add_argument(f"--profile-directory={config.EDGE_PROFILE_DIR}")
    if config.EDGE_DRIVER_PATH:
        service = EdgeService(executable_path=config.EDGE_DRIVER_PATH)
        return webdriver.Edge(service=service, options=opts)
    return webdriver.Edge(options=opts)

def wait_form_loaded(driver):
    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.XPATH, "//div[@role='listitem']")))

def find_question_card(driver, text):
    xp = (
        f"//div[@role='listitem']"
        f"[.//div[contains(normalize-space(), '{text}')]"
        f" or .//span[contains(normalize-space(), '{text}')]"
        f" or .//div[contains(., '{text}')]]"
    )
    return WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, xp)))

def safe_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)

def is_other_label(text: str) -> bool:
    if not config.AVOID_OTHER:
        return False
    t = (text or "").strip().lower()
    return any(k.lower() in t for k in config.OTHER_KEYWORDS)

# ---------------- Interactions ----------------
def click_radio_option(driver, question_text, option_label):
    card = find_question_card(driver, question_text)
    # aria-label
    try:
        el = WebDriverWait(card, 8).until(
            EC.element_to_be_clickable((By.XPATH, f".//div[@role='radio' and @aria-label='{option_label}']"))
        )
        safe_click(driver, el)
        return
    except TimeoutException:
        pass
    # label จาก span
    el = WebDriverWait(card, 8).until(
        EC.element_to_be_clickable((By.XPATH, f".//div[@role='radio'][.//span[normalize-space()='{option_label}']]"))
    )
    safe_click(driver, el)

def debug_dump_options(driver, question_text):
    card = find_question_card(driver, question_text)
    opts = card.find_elements(By.XPATH, ".//div[@role='checkbox']|.//div[@role='radio']")
    out = []
    for el in opts:
        label = el.get_attribute("aria-label") or ""
        if not label.strip():
            try:
                span = el.find_element(By.XPATH, ".//span[normalize-space()]")
                label = span.text.strip()
            except Exception:
                label = "(no-label)"
        out.append(label)
    print(f"[DEBUG] Options in '{question_text}': {out}")

def click_checkbox_options(driver, question_text, want_labels: Optional[List[str]] = None, k_auto: int = 2):
    card = find_question_card(driver, question_text)
    opt_elems = card.find_elements(By.XPATH, ".//div[@role='checkbox']")
    options = []
    for el in opt_elems:
        label = el.get_attribute("aria-label") or ""
        if not label.strip():
            try:
                span = el.find_element(By.XPATH, ".//span[normalize-space()]")
                label = span.text.strip()
            except Exception:
                label = ""
        options.append((el, label))

    if config.DEBUG_PRINT_CHECKBOX:
        print(f"[DEBUG] Checkbox count for '{question_text}':", len(options))
        print([lbl for (_, lbl) in options])

    # ไม่มี checkbox → ลอง radio (เลือกแบบไม่เอา 'อื่นๆ')
    if not options:
        radios = card.find_elements(By.XPATH, ".//div[@role='radio']")
        radios_no_other = []
        for el in radios:
            lab = el.get_attribute("aria-label") or ""
            if not lab.strip():
                try:
                    lab = el.find_element(By.XPATH, ".//span[normalize-space()]").text.strip()
                except Exception:
                    lab = ""
            if not is_other_label(lab):
                radios_no_other.append(el)
        if radios_no_other:
            safe_click(driver, random.choice(radios_no_other))
        else:
            print(f"[WARN] No selectable options (non-other) in '{question_text}'. Skipped.")
        return

    def norm(s): return (s or "").strip().lower()
    # กรอง 'อื่นๆ'
    options = [(el, lbl) for (el, lbl) in options if lbl and not is_other_label(lbl)]
    label_to_el = {norm(lbl): el for (el, lbl) in options}

    clicked = set()
    # พยายามคลิกตามที่อยากได้ก่อน
    if want_labels:
        for w in want_labels:
            if is_other_label(w):
                continue
            el = label_to_el.get(norm(w))
            if el and el.get_attribute("aria-checked") != "true":
                safe_click(driver, el)
                clicked.add(norm(w))

    # เติมแบบสุ่มจากตัวเลือกที่เหลือ
    remaining = [el for (el, lbl) in options if norm(lbl) not in clicked]
    need = max(0, k_auto - len(clicked))
    if need > 0 and remaining:
        for el in random.sample(remaining, k=min(need, len(remaining))):
            if el.get_attribute("aria-checked") != "true":
                safe_click(driver, el)

def select_dropdown(driver, question_text, option_label):
    card = find_question_card(driver, question_text)
    open_candidates = [
        ".//*[@role='listbox']",
        ".//*[@role='combobox']",
        ".//div[contains(@class,'quantumWizMenuPaperselectEl')]",
        ".//div[contains(@class,'MocG8c')]",
    ]
    opener = None
    for ox in open_candidates:
        try:
            opener = WebDriverWait(card, 5).until(EC.element_to_be_clickable((By.XPATH, ox)))
            break
        except TimeoutException:
            continue
    if not opener:
        raise TimeoutException("ไม่พบตัวเปิด dropdown")
    safe_click(driver, opener)

    # เลือกตามชื่อก่อน
    opt_candidates = [
        f"//div[@role='option']//span[normalize-space()='{option_label}']",
        f"//div[@role='option' and .//span[normalize-space()='{option_label}']]",
        f"//div[@role='option' and normalize-space()='{option_label}']",
        f"//span[@class='vRMGwf oJeWuf' and normalize-space()='{option_label}']",
    ]
    for ox in opt_candidates:
        try:
            opt = WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.XPATH, ox)))
            safe_click(driver, opt)
            return
        except TimeoutException:
            continue

    # Fallback: เลือกตัวเลือกแรกที่ไม่ใช่ 'อื่นๆ'
    all_opts = WebDriverWait(driver, 6).until(
        EC.presence_of_all_elements_located((By.XPATH, "//div[@role='option']//span[normalize-space()]"))
    )
    for o in all_opts:
        label = o.text.strip()
        if not is_other_label(label):
            safe_click(driver, o)
            return
    print(f"[WARN] Dropdown '{question_text}' has only 'อื่นๆ' or no valid options. Skipped.")

def select_major_smart(driver, option_label):
    # ลอง radio ก่อน → ไม่ได้ค่อย dropdown
    try:
        click_radio_option(driver, config.Q["major"], option_label)
        return
    except Exception:
        pass
    select_dropdown(driver, config.Q["major"], option_label)

# -------------- กริดตอนที่ 2 --------------
def pick_col_index(n: int) -> int:
    """เลือกคอลัมน์ตามโหมดใน config."""
    if n <= 1:
        return 0
    mode = getattr(config, "GRID_SELECT_MODE", "uniform")
    if mode == "uniform":
        return random.randrange(n)
    if mode == "indices" and config.GRID_ALLOWED_INDICES:
        allowed = [i for i in config.GRID_ALLOWED_INDICES if 0 <= i < n]
        return random.choice(allowed) if allowed else random.randrange(n)
    if mode == "topk":
        k = max(1, min(getattr(config, "GRID_TOP_K", 2), n))
        side = getattr(config, "GRID_HIGH_END_SIDE", "right").lower()
        if side == "left":
            indices = list(range(0, k))
        else:
            indices = list(range(n - k, n))
        return random.choice(indices)
    # default
    return random.randrange(n)

def answer_grid(driver, grid_title):
    card = find_question_card(driver, grid_title)
    groups = card.find_elements(By.XPATH, ".//div[@role='radiogroup']")
    if not groups:
        groups = card.find_elements(By.XPATH, ".//div[@role='radio']/ancestor::div[@role='radiogroup']")
    if config.DEBUG_PRINT_GRID:
        print(f"[GRID] '{grid_title}' rows={len(groups)}")
    for idx, rg in enumerate(groups, start=1):
        radios = rg.find_elements(By.XPATH, ".//div[@role='radio']")
        if not radios:
            radios = rg.find_elements(By.XPATH, ".//span[normalize-space()]/ancestor::div[@role='radio']")
        if config.DEBUG_PRINT_GRID:
            print(f"   - row#{idx}: cols={len(radios)}")
        if not radios:
            continue
        col_idx = pick_col_index(len(radios))
        safe_click(driver, radios[col_idx])

# -------------- ฟิลด์ทั่วไป --------------
def fill_short_answer(driver, question_text, value):
    card = find_question_card(driver, question_text)
    inp = WebDriverWait(card, 12).until(EC.element_to_be_clickable((By.XPATH, ".//input[@type='text' or @type='number']")))
    inp.clear()
    inp.send_keys(str(value))

def submit_form(driver):
    xp = "//span[text()='ส่ง']/ancestor::div[@role='button']"
    btn = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, xp)))
    safe_click(driver, btn)

# -------------- เติมคำตอบ 1 ชุด --------------
def fill_one_response(driver):
    Q = config.Q
    P = config.POOLS

    # ตอนที่ 1
    click_radio_option(driver, Q["gender"], random.choice(P["genders"]))
    age_lo, age_hi = config.AGE_RANGE
    fill_short_answer(driver, Q["age"], random.randint(age_lo, age_hi))
    click_radio_option(driver, Q["year"], random.choice(P["years"]))
    select_major_smart(driver, random.choice(P["majors"]))
    freq = random.choices(config.FREQ_VALUES, weights=config.FREQ_WEIGHTS)[0]
    fill_short_answer(driver, Q["freq"], freq)
    click_radio_option(driver, Q["device"],  random.choice(P["devices"]))
    click_radio_option(driver, Q["browser"], random.choice(P["browsers"]))

    # วัตถุประสงค์ (checkbox)
    if config.DEBUG_PRINT_CHECKBOX:
        debug_dump_options(driver, Q["purpose"])
    want = random.sample(P["purposes_preferred"], k=min(3, len(P["purposes_preferred"])))
    click_checkbox_options(driver, Q["purpose"], want_labels=want, k_auto=2)

    # ตอนที่ 2 — กริด
    for title in config.GRID_TITLES:
        answer_grid(driver, title)

    # ตอนที่ 3 — ข้อเสนอแนะ (ถ้ามี)
    try:
        card = find_question_card(driver, Q["comment"])
        area = WebDriverWait(card, 6).until(EC.element_to_be_clickable((By.XPATH, ".//textarea")))
        area.clear()
        area.send_keys(random.choice(P["comments"]))
    except Exception:
        pass

    submit_form(driver)

# ---------------- MAIN ----------------
if __name__ == "__main__":
    view_url = resolve_viewform_url(config.FORM_URL)
    driver = make_driver()
    try:
        driver.get(view_url)
        wait_form_loaded(driver)

        for i in range(1, config.N_SUBMISSIONS + 1):
            fill_one_response(driver)
            print(f"[{i}/{config.N_SUBMISSIONS}] submitted")

            # หลังส่ง: รีเฟรชกลับหน้า viewform ใหม่ (หรือจะทำเป็นกด "ส่งคำตอบอีกครั้ง" ก็ได้)
            time.sleep(random.uniform(0.5, 1.1))
            if config.REFRESH_AFTER_SUBMIT:
                driver.get(view_url)
            else:
                # วิธีทางเลือก (ไม่ใช้ตามดีฟอลต์)
                for xp in ["//a[normalize-space()='ส่งคำตอบอีกครั้ง']",
                           "//a[normalize-space()='Submit another response']"]:
                    try:
                        el = WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.XPATH, xp)))
                        safe_click(driver, el)
                        break
                    except Exception:
                        pass
            wait_form_loaded(driver)
            time.sleep(random.uniform(config.MIN_DELAY, config.MAX_DELAY))
    finally:
        driver.quit()
