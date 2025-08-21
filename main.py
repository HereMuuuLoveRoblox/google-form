# fill_ssru_form_edge_full_v4.py
# -*- coding: utf-8 -*-
"""
Selenium + Microsoft Edge — กรอก Google Form (ภาษาไทย) อัตโนมัติ 100 ชุด
- FIX: ตอบกริดโดย "เลือกตามตำแหน่งคอลัมน์" (ไม่พึ่ง aria-label '1'..'5')
- เช็กบ็อกซ์: แมตช์ด้วย aria-label/ข้อความ และ fallback เป็นสุ่ม/วิทยุ
- รองรับ forms.gle (resolve เป็น viewform)
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

# =======================
# CONFIG
# =======================
FORM_URL = "https://forms.gle/jCAgKEsXakH5U3gR9"   # ใส่ forms.gle ได้ เดี๋ยวขยายเอง
N_SUBMISSIONS = 100
MIN_DELAY, MAX_DELAY = 0.6, 2.0

EDGE_DRIVER_PATH: Optional[str] = None   # ใช้ Selenium Manager ถ้า None
USE_EDGE_PROFILE = False
EDGE_USER_DATA_DIR = r"C:\Users\<USERNAME>\AppData\Local\Microsoft\Edge\User Data"
EDGE_PROFILE_DIR = "Default"

# Debug พิมพ์โครงสร้างกริด/ตัวเลือก (ช่วยไล่บั๊ก)
DEBUG_PRINT_GRID = True
DEBUG_PRINT_CHECKBOX = False

# =======================
# DATA POOLS
# =======================
GENDERS = ["ชาย", "หญิง"]
YEARS = ["ปี1", "ปี2", "ปี3", "ปี4"]
DEVICES = ["สมาร์ทโฟน android", "สมาร์ทโฟน ios", "แท็บเล็ต", "แล็ปท็อป", "เดสก์ท็อป"]
BROWSERS = ["Chrome", "Safari", "Edge", "Firefox"]
MAJORS = [
    "สาขาวิชาชีววิทยาสิ่งแวดล้อม",
    "สาขาวิชาจุลชีววิทยาอุตสาหกรรมอาหารและนวัตกรรมชีวภาพ",
    "สาขาวิชาวิทยาศาสตร์และนวัตกรรม",
    "สาขาวิชาคหกรรมศาสตร์",
    "สาขาวิชาวิทยาศาสตร์และเทคโนโลยีการอาหาร",
    "สาขาวิชานิติวิทยาศาสตร์",
    "สาขาวิชาวิทยาศาสตร์การกีฬาและสุขภาพ",
    "สาขาวิชาวิทยาการคอมพิวเตอร์และนวัตกรรมข้อมูล",
    "สาขาวิชาวิทยาศาสตร์และเทคโนโลยีสิ่งแวดล้อม",
    "สาขาวิชานวัตกรรมอาหารและเชฟมืออาชีพ",
    "สาขาวิชาการจัดการนวัตกรรมดิจิทัลและคอนเทนท์",
]
PURPOSES_WANT = [
    "ข่าว/ประกาศ",
    "ปฏิทิน/ตารางเรียน-สอบ",
    "ดาวน์โหลดเอกสาร",
    "ข้อมูลหลักสูตร/ภาควิชา",
    "ติดต่อหน่วยงาน/อาจารย์",
]
COMMENTS_POOL = [
    "เว็บไซต์ใช้งานง่ายขึ้นมาก ขอให้มีหน้ารวมข่าวทุนวิจัย",
    "อยากให้เพิ่มขนาดฟอนต์อีกเล็กน้อย",
    "ขอให้ประกาศสำคัญปักหมุดไว้ด้านบน",
    "มือถือใช้งานดีขึ้น แต่บางหน้ายังช้าเล็กน้อย",
    "อยากมีระบบแจ้งเตือนข่าวใหม่ทางอีเมล",
]

# Likert bias → เอียงไปกลางๆ
def pick_index_uniform(n: int) -> int:
    # สุ่ม index 0..n-1 แบบโอกาสเท่ากันทุกคอลัมน์
    return 0 if n <= 1 else random.randrange(n)

# =======================
# FORM TEXT KEYS (ต้องตรงกับของจริง)
# =======================
Q_GENDER   = "เพศ"
Q_AGE      = "อายุ"
Q_YEAR     = "ระดับชั้น"
Q_MAJOR    = "สาขาวิชาหลักสูตร"
Q_FREQ     = "ความถี่ในการเข้าเว็บไซต์คณะ"
Q_DEVICE   = "อุปกรณ์ที่ใช้เข้าเว็บไซต์บ่อยมากที่สุด"
Q_BROWSER  = "เบราว์เซอร์ที่ใช้มากที่สุด"
Q_PURPOSE  = "วัตถุประสงค์หลักในการเข้าเว็บไซต์"
Q_COMMENT  = "ข้อเสนอแนะ"

GRID_TITLES = [
    "ด้านเนื้อหา",
    "ด้านการออกแบบและความสวยงาม",
    "ด้านการใช้งานและการเข้าถึง",
    "ด้านประสิทธิภาพและความรวดเร็ว",
    "ด้านการสื่อสารและบริการสนับสนุน",
]

# =======================
# UTILITIES
# =======================
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
    if USE_EDGE_PROFILE:
        opts.add_argument(f"--user-data-dir={EDGE_USER_DATA_DIR}")
        opts.add_argument(f"--profile-directory={EDGE_PROFILE_DIR}")
    if EDGE_DRIVER_PATH:
        service = EdgeService(executable_path=EDGE_DRIVER_PATH)
        return webdriver.Edge(service=service, options=opts)
    return webdriver.Edge(options=opts)

def open_form(driver):
    url = resolve_viewform_url(FORM_URL)
    driver.get(url)
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

def click_radio_option(driver, question_text, option_label):
    card = find_question_card(driver, question_text)
    # radio มีทั้งแบบ aria-label = label และแบบ span ข้างใน
    try:
        el = WebDriverWait(card, 8).until(
            EC.element_to_be_clickable((By.XPATH, f".//div[@role='radio' and @aria-label='{option_label}']"))
        )
        safe_click(driver, el)
        return
    except TimeoutException:
        pass
    # หาโดยข้อความใน span
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

def is_other_label(text: str) -> bool:
    t = (text or "").strip().lower()
    # ครอบคลุม: อื่น, อื่นๆ, อื่น ๆ, other, others
    return ("อื่น" in t) or ("other" in t)

def click_checkbox_options(driver, question_text, want_labels: Optional[List[str]] = None, k_auto: int = 2):
    card = find_question_card(driver, question_text)
    opt_elems = card.find_elements(By.XPATH, ".//div[@role='checkbox']")

    # รวบรวม (element, label)
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

    # ถ้าไม่มี checkbox → ลอง radio และ “ไม่เลือกอื่นๆ” เช่นกัน
    if not options:
        radios = card.find_elements(By.XPATH, ".//div[@role='radio']")
        # กรองตัวเลือกอื่นๆ ออก
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

    # map label -> element (กรอง 'อื่นๆ' ออกตั้งแต่ต้น)
    def norm(s): return (s or "").strip().lower()
    options = [(el, lbl) for (el, lbl) in options if lbl and not is_other_label(lbl)]
    label_to_el = {norm(lbl): el for (el, lbl) in options}

    clicked = set()

    # 1) พยายามคลิกตามรายการที่อยากได้ (อย่าคลิกถ้าเป็น 'อื่นๆ')
    if want_labels:
        for w in want_labels:
            if is_other_label(w):
                continue
            el = label_to_el.get(norm(w))
            if el and el.get_attribute("aria-checked") != "true":
                safe_click(driver, el)
                clicked.add(norm(w))
    
    # 2) เติมแบบสุ่มจากตัวเลือกที่เหลือ (ยังคง “ไม่เอาอื่นๆ”)
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

    # พยายามเลือกตามชื่อก่อน
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

    # Fallback: เลือกตัวเลือกแรกที่ "ไม่ใช่ อื่นๆ"
    all_opts = WebDriverWait(driver, 6).until(
        EC.presence_of_all_elements_located((By.XPATH, "//div[@role='option']//span[normalize-space()]"))
    )
    for o in all_opts:
        label = o.text.strip()
        if not is_other_label(label):
            safe_click(driver, o)
            return
    # ถ้าจำใจไม่มีจริง ๆ ก็ข้าม
    print(f"[WARN] Dropdown '{question_text}' has only 'อื่นๆ' or no valid options. Skipped.")


def select_major_smart(driver, option_label):
    # ลองเป็น radio ก่อน → ไม่ได้ค่อย dropdown
    try:
        click_radio_option(driver, Q_MAJOR, option_label)
        return
    except Exception:
        pass
    select_dropdown(driver, Q_MAJOR, option_label)

def fill_short_answer(driver, question_text, value):
    card = find_question_card(driver, question_text)
    inp = WebDriverWait(card, 12).until(EC.element_to_be_clickable((By.XPATH, ".//input[@type='text' or @type='number']")))
    inp.clear()
    inp.send_keys(str(value))

# ========== KEY FIX ==========
def answer_grid(driver, grid_title):
    """
    เลือกคำตอบในกริดแบบ "ตามตำแหน่งคอลัมน์" ไม่สน aria-label
    - หาทุก radiogroup (แถว)
    - ในแต่ละแถว ดึงปุ่ม radio ทั้งหมด แล้วเลือก index แบบเอียงไปกลางๆ
    """
    card = find_question_card(driver, grid_title)
    groups = card.find_elements(By.XPATH, ".//div[@role='radiogroup']")
    if not groups:
        # fallback บางธีม
        groups = card.find_elements(By.XPATH, ".//div[@role='radio']/ancestor::div[@role='radiogroup']")

    if DEBUG_PRINT_GRID:
        print(f"[GRID] '{grid_title}' rows={len(groups)}")

    for idx, rg in enumerate(groups, start=1):
        radios = rg.find_elements(By.XPATH, ".//div[@role='radio']")
        if not radios:
            # บางเคส radio อยู่ลึก: ลองหยิบ span แล้วขึ้น ancestor
            radios = rg.find_elements(By.XPATH, ".//span[normalize-space()]/ancestor::div[@role='radio']")
        if DEBUG_PRINT_GRID:
            print(f"   - row#{idx}: cols={len(radios)}")
        if not radios:
            continue
        col_idx = pick_index_uniform(len(radios))
        safe_click(driver, radios[col_idx])

def submit_form(driver):
    xp = "//span[text()='ส่ง']/ancestor::div[@role='button']"
    btn = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, xp)))
    safe_click(driver, btn)

def click_submit_another(driver):
    for xp in [
        "//a[normalize-space()='ส่งคำตอบอีกครั้ง']",
        "//a[normalize-space()='Submit another response']",
    ]:
        try:
            el = WebDriverWait(driver, 6).until(EC.element_to_be_clickable((By.XPATH, xp)))
            safe_click(driver, el)
            return
        except Exception:
            pass
    driver.get(resolve_viewform_url(FORM_URL))

# =======================
# ONE RESPONSE
# =======================
def fill_one_response(driver):
    # ตอนที่ 1
    click_radio_option(driver, Q_GENDER, random.choice(GENDERS))
    fill_short_answer(driver, Q_AGE, random.randint(19, 24))
    click_radio_option(driver, Q_YEAR, random.choice(YEARS))
    select_major_smart(driver, random.choice(MAJORS))
    fill_short_answer(driver, Q_FREQ, random.choices([0,1,2,3,4,5,6,7], weights=[2,5,15,25,25,15,8,5])[0])
    click_radio_option(driver, Q_DEVICE, random.choice(DEVICES))
    click_radio_option(driver, Q_BROWSER, random.choice(BROWSERS))

    # วัตถุประสงค์ (พยายามตามชื่อ → ไม่ครบสุ่มเติม)
    if DEBUG_PRINT_CHECKBOX:
        debug_dump_options(driver, Q_PURPOSE)
    click_checkbox_options(driver, Q_PURPOSE, want_labels=random.sample(PURPOSES_WANT, k=3), k_auto=2)

    # ตอนที่ 2: กริด 5 ด้าน
    for title in GRID_TITLES:
        answer_grid(driver, title)

    # ตอนที่ 3: ข้อเสนอแนะ
    try:
        card = find_question_card(driver, Q_COMMENT)
        area = WebDriverWait(card, 6).until(EC.element_to_be_clickable((By.XPATH, ".//textarea")))
        area.clear()
        area.send_keys(random.choice(COMMENTS_POOL))
    except Exception:
        pass

    submit_form(driver)

def wait_form_loaded(driver):
    WebDriverWait(driver, 25).until(EC.presence_of_element_located((By.XPATH, "//div[@role='listitem']")))


# =======================
# MAIN
# =======================
if __name__ == "__main__":
    view_url = resolve_viewform_url(FORM_URL)
    driver = make_driver()
    try:
        driver.get(view_url)
        wait_form_loaded(driver)

        for i in range(1, N_SUBMISSIONS + 1):
            fill_one_response(driver)
            print(f"[{i}/{N_SUBMISSIONS}] submitted")

            # พักสั้น ๆ กัน rate-limit แล้วเปิดหน้า viewform ใหม่ 'ทันที'
            time.sleep(random.uniform(0.5, 1.1))
            driver.get(view_url)
            wait_form_loaded(driver)
            time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))
    finally:
        driver.quit()