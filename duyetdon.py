import time, traceback, threading, sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, ElementClickInterceptedException, JavascriptException,
    StaleElementReferenceException, NoSuchElementException, ElementNotInteractableException
)

# ---- Tkinter GUI ----
import tkinter as tk
from tkinter import ttk, messagebox

# ============== LOG UI HELPERS ==============
class UILogger:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def log(self, msg):
        try:
            print(msg)
        except UnicodeEncodeError:
            print(msg.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding))
        if self.text_widget:
            self.text_widget.after(0, lambda: self._append(msg))

    def _append(self, msg):
        self.text_widget.configure(state="normal")
        self.text_widget.insert("end", msg + "\n")
        self.text_widget.see("end")
        self.text_widget.configure(state="disabled")

# ============== WAITERS / HELPERS ==============
def wait_page_idle(driver, wait, extra_ms=300):
    wait.until(lambda x: x.execute_script("return document.readyState") == "complete")
    time.sleep(extra_ms/1000.0)

def switch_to_iframe_containing_table(driver, table_id="tblTTThuaDat", timeout=10):
    # quay về top trước
    driver.switch_to.default_content()
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    deadline = time.time() + timeout
    for idx in range(len(iframes)):
        if time.time() > deadline:
            break
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")  # refresh
        try:
            driver.switch_to.frame(iframes[idx])
            # kiểm tra có bảng không
            if driver.find_elements(By.CSS_SELECTOR, f"#{table_id}"):
                return True
            # nếu còn iframe lồng nhau
            inner_iframes = driver.find_elements(By.TAG_NAME, "iframe")
            for j in range(len(inner_iframes)):
                driver.switch_to.frame(inner_iframes[j])
                if driver.find_elements(By.CSS_SELECTOR, f"#{table_id}"):
                    return True
                driver.switch_to.parent_frame()
        except Exception:
            continue
    driver.switch_to.default_content()
    return False


def wait_for_table_loaded(driver, table_id="tblTTThuaDat", timeout=15):
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CSS_SELECTOR, f"#{table_id}_processing"))
        )
    except TimeoutException:
        pass

def safe_click_row_css(driver, wait, row_css="#tblTraCuuDotBanGiao tbody tr", logger=None):
    wait_page_idle(driver, wait, 300)
    row = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, row_css)))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
    cell = row.find_element(By.CSS_SELECTOR, "td:nth-child(2)")
    try:
        wait.until(EC.element_to_be_clickable((By.XPATH, "//table[@id='tblTraCuuDotBanGiao']//tbody//tr[1]//td[2]")))
        cell.click()
        return
    except ElementClickInterceptedException:
        wait_page_idle(driver, wait, 300)
        try:
            cell.click()
            return
        except ElementClickInterceptedException:
            pass
    try:
        driver.execute_script("""
            document.querySelectorAll('.jquery-loading-modal__bg')
                  .forEach(el => { el.style.pointerEvents='none'; el.style.display='none'; });
        """)
    except JavascriptException:
        pass
    try:
        driver.execute_script("arguments[0].click();", cell)
        return
    except Exception:
        pass
    first_cell = row.find_element(By.CSS_SELECTOR, "td:nth-child(1)")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", first_cell)
    driver.execute_script("arguments[0].click();", first_cell)

def goto_page(driver, page_number, table_id="tblTTThuaDat", verify_timeout=5):
    driver.execute_script(f"""
        if (window.jQuery && jQuery.fn.DataTable) {{
            var table = jQuery('#{table_id}').DataTable();
            var info  = table.page.info();
            var maxp  = info.pages || 1;
            var target = Math.max(0, Math.min({page_number}-1, maxp-1));
            table.page(target).draw('page');
        }}
    """)
    # verify page changed
    end = time.time() + verify_timeout
    target0 = max(0, page_number-1)
    while time.time() < end:
        ok = driver.execute_script(f"""
            try {{
                var t = jQuery('#{table_id}').DataTable();
                return t.page.info().page;
            }} catch(e){{ return -1; }}
        """)
        if ok == target0:
            return True
        time.sleep(0.2)
    return False

def go_next_datatables(driver, table_id="tblTTThuaDat", timeout=15):
    wait = WebDriverWait(driver, timeout)
    try:
        wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, f"#{table_id}_processing")))
    except TimeoutException:
        pass
    li_next = wait.until(EC.presence_of_element_located((By.ID, f"{table_id}_next")))
    if "disabled" in (li_next.get_attribute("class") or ""):
        return False
    a_next = li_next.find_element(By.TAG_NAME, "a")
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"#{table_id}_next a")))
    w, h, vis = driver.execute_script("""
        const a = arguments[0];
        const r = a.getBoundingClientRect();
        const style = window.getComputedStyle(a);
        return [r.width, r.height, style.visibility !== 'hidden' && style.display !== 'none'];
    """, a_next)
    if not (w > 0 and h > 0 and vis):
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", a_next)
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, f"#{table_id}_next a")))
    first_row = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, f"#{table_id} tbody tr")))
    try:
        a_next.click()
    except Exception:
        driver.execute_script("arguments[0].click();", a_next)
    try:
        wait.until(EC.staleness_of(first_row))
    except (TimeoutException, StaleElementReferenceException):
        try:
            wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, f"#{table_id}_processing")))
        except TimeoutException:
            pass
    return True

def handle_whole_page_action(driver, logger: UILogger, table_id="tblTTThuaDat", timeout=15):
    """
    Chọn tất cả các hàng trên trang hiện tại (Shift+Click), sau đó lặp qua
    và bỏ chọn (Ctrl+Click) những hàng đã có trạng thái "Đã duyệt ghi đè"
    để chỉ giữ lại các hàng "Chưa xử lý".
    """
    wait = WebDriverWait(driver, timeout)
    wait.until(EC.presence_of_element_located((By.ID, table_id)))
    rows = driver.find_elements(By.CSS_SELECTOR, f"#{table_id} tbody > tr:not(.child)")

    # Lọc các hàng đang hiển thị và có thể tương tác
    visible_rows = []
    for r in rows:
        try:
            tds = r.find_elements(By.CSS_SELECTOR, "td")
            if tds and r.is_displayed():
                visible_rows.append((r, tds))
        except StaleElementReferenceException:
            continue

    if len(visible_rows) < 1:
        logger.log("   (Không có hàng nào hiển thị để chọn)")
        return 0

    first_row, first_tds = visible_rows[0]
    last_row, last_tds = visible_rows[-1]

    def pick_click_target(row, tds):
        # Ưu tiên click vào checkbox hoặc button nếu có, fallback về ô đầu tiên
        for css in ["input[type='checkbox']:not([disabled])", "button", "a"]:
            try:
                el = row.find_element(By.CSS_SELECTOR, css)
                if el.is_displayed(): return el
            except NoSuchElementException: pass
        return tds[0]

    first_target = pick_click_target(first_row, first_tds)
    last_target = pick_click_target(last_row, last_tds)

    def ensure_visible_and_sized(el):
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        WebDriverWait(driver, timeout).until(lambda d: d.execute_script("""
            const r = arguments[0].getBoundingClientRect();
            const s = getComputedStyle(arguments[0]);
            return r.width > 0 && r.height > 0 && s.display!=='none' && s.visibility!=='hidden';
        """, el))

    try:
        ensure_visible_and_sized(first_target)
        first_target.click() # Click hàng đầu
        if len(visible_rows) > 1:
            ensure_visible_and_sized(last_target)
            # Giữ SHIFT và click hàng cuối để chọn tất cả
            ActionChains(driver).key_down(Keys.SHIFT).click(last_target).key_up(Keys.SHIFT).perform()
    except Exception as e:
        logger.log(f"   (Lỗi Shift-Click, thử fallback... Lỗi: {e})")
        # Fallback nếu Shift-Click lỗi: chọn từng cái một
        for row, tds in visible_rows:
            try:
                target = pick_click_target(row, tds)
                ensure_visible_and_sized(target)
                target.click()
            except Exception:
                continue

    logger.log("   → Đã chọn tất cả, bắt đầu lọc bỏ những bản ghi đã duyệt...")
    time.sleep(0.2) # Chờ một chút để UI cập nhật trạng thái "selected"

    # Bỏ chọn những hàng đã được duyệt
    actions = ActionChains(driver).key_down(Keys.CONTROL)
    deselected_count = 0
    # Lấy lại danh sách hàng đã chọn (có class 'selected')
    selected_rows = driver.find_elements(By.CSS_SELECTOR, f"#{table_id} tbody tr.selected")
    for row in selected_rows:
        try:
            txt = (row.get_attribute("innerText") or row.text).strip().lower()
            if "đã duyệt ghi đè" in txt:
                actions.click(row.find_element(By.CSS_SELECTOR, "td:first-child"))
                deselected_count += 1
        except (StaleElementReferenceException, NoSuchElementException):
            continue
    actions.key_up(Keys.CONTROL).perform()

    # Kiểm tra lại số lượng đã chọn bằng API của DataTable
    selected_count = driver.execute_script(f"""
        try {{
            if (window.jQuery && jQuery.fn.DataTable) {{
                const dt = jQuery("#{table_id}").DataTable();
                return dt.rows({{selected:true, page:'current'}}).count();
            }}
        }} catch(e) {{}}
        const table = document.querySelector("#{table_id}");
        return table ? table.querySelectorAll("tbody tr.selected").length : 0;
    """)

    if deselected_count > 0:
        logger.log(f"   → Đã bỏ chọn {deselected_count} bản ghi đã duyệt. Còn lại {selected_count} bản ghi.")

    return selected_count

def quick_confirm_if_present(driver, soft_timeout=1.0):
    sw = WebDriverWait(driver, soft_timeout)
    els = driver.find_elements(By.CSS_SELECTOR, ".swal2-confirm")
    if els:
        sw.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".swal2-confirm"))).click()
        return True
    sel = ".modal.show .btn-primary, .modal.in .btn-primary, .bootbox .btn-primary"
    els = driver.find_elements(By.CSS_SELECTOR, sel)
    if els:
        sw.until(EC.element_to_be_clickable((By.CSS_SELECTOR, sel))).click()
        return True
    xp = "//button[contains(., 'Đồng ý') or contains(., 'Xác nhận') or contains(., 'OK')]"
    els = driver.find_elements(By.XPATH, xp)
    if els:
        sw.until(EC.element_to_be_clickable((By.XPATH, xp))).click()
        return True
    return False

def wait_processing_quick(driver, table_id="tblTTThuaDat", max_wait=6):
    def cond(d):
        try:
            ajax_zero = d.execute_script("return (window.jQuery ? jQuery.active : 0)") == 0
            proc = d.execute_script(f"""
                var e = document.querySelector('#{table_id}_processing');
                if(!e) return true;
                var s = getComputedStyle(e);
                return (s.display==='none' || s.visibility==='hidden' || e.offsetParent===null);
            """)
            return ajax_zero and proc
        except Exception:
            return True
    try:
        WebDriverWait(driver, max_wait, poll_frequency=0.1).until(cond)
        return True
    except Exception:
        return False
    
def hard_jump_pagination(driver, page_number, table_id="tblTTThuaDat", timeout=10):
    wait = WebDriverWait(driver, timeout)
    # xác định trang hiện tại
    cur = driver.execute_script(f"""
        try {{
            return jQuery('#{table_id}').DataTable().page.info().page + 1;
        }} catch(e) {{ return 1; }}
    """) or 1

    if page_number == cur:
        return True

    # nếu có nút số trang, thử click trực tiếp
    try:
        btn = wait.until(EC.presence_of_element_located((
            By.XPATH, f"//div[@id='{table_id}_paginate']//a[normalize-space(text())='{page_number}']"
        )))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)
    except TimeoutException:
        # nếu không có nút số trang (hiển thị dạng next/prev) → lặp next/prev
        step = 1 if page_number > cur else -1
        next_sel = f"#{table_id}_next a"
        prev_sel = f"#{table_id}_previous a"
        while cur != page_number:
            sel = next_sel if step == 1 else prev_sel
            try:
                a = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
                a.click()
            except Exception:
                driver.execute_script("document.querySelector(arguments[0])?.click()", sel)
            wait_for_table_loaded(driver, table_id, timeout=10)
            cur = driver.execute_script(f"return jQuery('#{table_id}').DataTable().page.info().page + 1;") or cur
            # tránh lặp vô hạn
            if (step == 1 and cur < page_number) or (step == -1 and cur > page_number):
                continue
            if cur == page_number:
                break

    wait_for_table_loaded(driver, table_id, timeout=10)
    cur2 = driver.execute_script(f"return jQuery('#{table_id}').DataTable().page.info().page + 1;")
    return cur2 == page_number

def click_duyet_ghi_de(driver, timeout=15, table_id="tblTTThuaDat"):
    wait = WebDriverWait(driver, timeout)
    btn = wait.until(EC.element_to_be_clickable((By.ID, "btnDropTTTD")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    try:
        btn.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", btn)
    item = wait.until(EC.presence_of_element_located((By.ID, "btnDuyetGhiDeTTTD")))
    for _ in range(3):
        vis = driver.execute_script("""
            const el = arguments[0];
            const r = el.getBoundingClientRect();
            const s = getComputedStyle(el);
            return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
        """, item)
        if vis: break
        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)
    wait.until(EC.element_to_be_clickable((By.ID, "btnDuyetGhiDeTTTD")))
    try:
        item.click()
    except ElementClickInterceptedException:
        driver.execute_script("arguments[0].click();", item)

    # 👉 xác nhận nhanh nếu có + chờ xử lý ngắn
    _ = quick_confirm_if_present(driver, soft_timeout=1.0)
    wait_processing_quick(driver, table_id=table_id, max_wait=6)

def switch_to_frame_having(driver, by, value, timeout=8):
    driver.switch_to.default_content()
    # thử ở top trước
    try:
        if driver.find_elements(by, value):
            return True
    except Exception:
        pass
    # duyệt qua tất cả iframes (kể cả lồng nhau)
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    deadline = time.time() + timeout
    for i in range(len(frames)):
        if time.time() > deadline: break
        driver.switch_to.default_content()
        frames = driver.find_elements(By.TAG_NAME, "iframe")  # refresh
        try:
            driver.switch_to.frame(frames[i])
            if driver.find_elements(by, value):
                return True
            # thử thêm 1 tầng lồng
            inner = driver.find_elements(By.TAG_NAME, "iframe")
            for j in range(len(inner)):
                driver.switch_to.frame(inner[j])
                if driver.find_elements(by, value):
                    return True
                driver.switch_to.parent_frame()
        except Exception:
            continue
    driver.switch_to.default_content()
    return False

def context_click_jstree_pick(driver, wait, anchor_id="j34_1_anchor",
                              menu_text="Thêm vào dữ liệu vận hành", logger=None):
    # 1) Đảm bảo đang ở frame có anchor
    def switch_to_frame_having(by, value, timeout=8):
        driver.switch_to.default_content()
        try:
            if driver.find_elements(by, value):
                return True
        except: pass
        import time
        deadline = time.time() + timeout
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        for i in range(len(frames)):
            if time.time() > deadline: break
            driver.switch_to.default_content()
            frames = driver.find_elements(By.TAG_NAME, "iframe")
            try:
                driver.switch_to.frame(frames[i])
                if driver.find_elements(by, value):
                    return True
                inner = driver.find_elements(By.TAG_NAME, "iframe")
                for j in range(len(inner)):
                    driver.switch_to.frame(inner[j])
                    if driver.find_elements(by, value):
                        return True
                    driver.switch_to.parent_frame()
            except:
                continue
        driver.switch_to.default_content()
        return False

    switch_to_frame_having(By.ID, anchor_id, timeout=8)

    anchor = wait.until(EC.presence_of_element_located((By.ID, anchor_id)))
    # 2) Lấy node <li> và container tree có size > 0
    li = anchor.find_element(By.XPATH, "./ancestor::li[1]")
    # Container ứng viên theo thứ tự dễ gặp
    candidates = [
        "./ancestor::*[contains(@class,'jstree-default')][1]",
        "./ancestor::*[contains(@class,'jstree')][1]",
        "./ancestor::*[contains(@class,'jstree-container-ul')][1]",
        "./ancestor::ul[1]"
    ]
    tree = None
    for xp in candidates:
        try:
            e = anchor.find_element(By.XPATH, xp)
            vis = driver.execute_script("""
                const el=arguments[0], r=el.getBoundingClientRect(), s=getComputedStyle(el);
                return {w:r.width, h:r.height, ok:(r.width>0 && r.height>0 && s.display!=='none' && s.visibility!=='hidden')};
            """, e)
            if vis["ok"]:
                tree = e
                break
        except: pass
    if tree is None:
        # fallback: dùng chính body
        tree = driver.find_element(By.TAG_NAME, "body")

    # 3) Left-click để select node (tăng xác suất menu hiện đúng)
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor)
        anchor.click()
    except Exception:
        # click JS nếu anchor không tương tác được
        try: driver.execute_script("arguments[0].click();", anchor)
        except: pass

    # 4) Right-click: ưu tiên vào <li>, nếu size 0 thì context trên container với offset
    import math, time as _t
    size_li = driver.execute_script("""
        const el=arguments[0], r=el.getBoundingClientRect(), s=getComputedStyle(el);
        return {w:r.width, h:r.height, ok:(r.width>0 && r.height>0 && s.display!=='none' && s.visibility!=='hidden'),
                cx:r.left + r.width/2, cy:r.top + Math.min(18, Math.max(10, r.height/2))};
    """, li)
    actions = ActionChains(driver)

    if size_li["ok"]:
        actions.context_click(li).perform()
    else:
        # dùng container tree
        rect = driver.execute_script("""
            const el=arguments[0], r=el.getBoundingClientRect();
            return {cx:r.left + r.width/2, cy:r.top + 40};
        """, tree)
        # reset pointer rồi move theo tọa độ tuyệt đối
        actions.move_by_offset(1,1).perform()
        # Selenium move_by_offset là tương đối so với vị trí hiện tại → dùng JS để bắn event native nếu cần
        try:
            actions.move_by_offset(int(rect["cx"]), int(rect["cy"])).context_click().perform()
        except Exception:
            driver.execute_script("""
                const el=arguments[0]; const p=arguments[1];
                const evt = new MouseEvent('contextmenu',{bubbles:true,cancelable:true,view:window,
                    clientX:p.cx, clientY:p.cy, button:2});
                el.dispatchEvent(evt);
            """, tree, rect)
    _t.sleep(0.3)  # cho menu render

    # 5) Nếu vẫn chưa thấy menu, gọi jsTree API show_contextmenu
    try:
        driver.execute_script("""
            try{
              var a=document.getElementById(arguments[0]);
              if(!a) return;
              var inst=null;
              if(window.jQuery){
                inst = jQuery(a).closest('.jstree-default,.jstree,.jstree-container-ul').jstree(true);
              }
              if(inst){
                var li = a.closest('li');
                inst.show_contextmenu(li||a);
              }
            }catch(e){}
        """, anchor_id)
    except Exception:
        pass

    # 6) Chờ menu vakata hiện + click item theo text (dùng contains thay vì so khớp tuyệt đối)
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "ul.vakata-context")))
    item = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//ul[contains(@class,'vakata-context')]//a[contains(., 'Thêm vào dữ liệu vận hành')]")
    ))
    try:
        item.click()
    except Exception:
        driver.execute_script("arguments[0].click();", item)

# ============== BOT CORE ==============
def run_bot(username, password, code, start_page, logger: UILogger):
    driver = None
    try:
        logger.log("🚀 Khởi động Chrome…")
        options = Options()
        options.add_argument("--start-maximized")
        # Tự động accept bất kỳ alert/prompt nào không xử lý:
        service = Service(r"D:\Python\chromedriver\chromedriver-win64\chromedriver-win64\chromedriver.exe")
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 20)

        driver.get("https://dla.mplis.gov.vn/dc/TichHopDongBoDuLieu/QuanLyDotBanGiao")

        # Đăng nhập
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(username)
        driver.find_element(By.NAME, "password").send_keys(password)
        driver.find_element(By.NAME, "password").send_keys(Keys.ENTER)
        logger.log("🔐 Đang đăng nhập…")
        # Nếu có captcha/xác minh thủ công, dừng lại cho người dùng thao tác  
        messagebox.showinfo("Xác minh", "Nếu có xác minh thủ công (captcha/SSO), hãy hoàn tất trên trình duyệt rồi bấm OK để tiếp tục.")

        # Tìm code
        try:
            old_first_row = driver.find_element(By.CSS_SELECTOR, "#tblTraCuuDotBanGiao tbody tr")
        except Exception:
            old_first_row = None
        driver.find_element(By.ID, "txtTraCuuDotBanGiao").send_keys(code)
        driver.find_element(By.ID, "txtTraCuuDotBanGiao").send_keys(Keys.ENTER)
        logger.log(f"🔎 Đang tìm kiếm CODE: {code}")
        if old_first_row:
            WebDriverWait(driver, 20).until(EC.staleness_of(old_first_row))

        safe_click_row_css(driver, wait)
        # Mở menu Chức năng → Xem danh sách
        logger.log("📂 Mở 'Chức năng' → 'Xem danh sách'…")
        wait.until(EC.element_to_be_clickable((By.ID, "drop1"))).click()
        wait.until(EC.element_to_be_clickable((By.ID, "btnXemDanhSach"))).click()

        # Sử dụng JavaScript để click, tránh lỗi bị che khuất
        btn_xu_ly = wait.until(EC.presence_of_element_located((By.ID, "btnXuLyDon")))
        driver.execute_script("arguments[0].scrollIntoView(true);", btn_xu_ly)
        driver.execute_script("arguments[0].click();", btn_xu_ly)

        # 2️⃣ Nhấp chuột phải và chọn menu
        context_click_jstree_pick(driver, wait,
            anchor_id="j34_1_anchor",
            menu_text="Thêm vào dữ liệu vận hành",
            logger=logger
        )

    except Exception as ex:
        logger.log(f"❌ Có lỗi xảy ra: {ex}")
        logger.log(traceback.format_exc())
        try:
            if driver:
                driver.quit()
        except Exception:
            pass

# ============== TKINTER UI ==============
def main():
    root = tk.Tk()
    root.title("Tự động duyệt - DakLak LIS")
    root.geometry("760x520")

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="x")

    # Inputs
    ttk.Label(frm, text="Username:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
    ent_user = ttk.Entry(frm, width=36)
    ent_user.grid(row=0, column=1, sticky="w", padx=4, pady=4)
    ent_user.insert(0, "dla.tuannhqt")  # default (có thể xoá)

    ttk.Label(frm, text="Password:").grid(row=1, column=0, sticky="e", padx=4, pady=4)
    ent_pass = ttk.Entry(frm, width=36, show="•")
    ent_pass.grid(row=1, column=1, sticky="w", padx=4, pady=4)
    ent_pass.insert(0, "NNMT@2025")  # default (có thể xoá)

    ttk.Label(frm, text="CODE (Mã đợt bàn giao):").grid(row=2, column=0, sticky="e", padx=4, pady=4)
    ent_code = ttk.Entry(frm, width=36)
    ent_code.grid(row=2, column=1, sticky="w", padx=4, pady=4)
    ent_code.insert(0, "8919bd18-177b-48fa-9ae9-9846d359a048")  # default (có thể xoá)

    ttk.Label(frm, text="Trang bắt đầu:").grid(row=3, column=0, sticky="e", padx=4, pady=4)
    ent_start = ttk.Entry(frm, width=12)
    ent_start.grid(row=3, column=1, sticky="w", padx=4, pady=4)
    ent_start.insert(0, "1")

    # Run button
    btn_run = ttk.Button(frm, text="Chạy")
    btn_run.grid(row=4, column=1, sticky="w", padx=4, pady=8)

    # Log box
    txt = tk.Text(root, height=18, state="disabled", bg="#0f1115", fg="#e5e7eb")
    txt.pack(fill="both", expand=True, padx=12, pady=(0,12))
    logger = UILogger(txt)

    def on_run():
        username = ent_user.get().strip()
        password = ent_pass.get()
        code     = ent_code.get().strip()
        try:
            start_page = int(ent_start.get().strip())
        except:
            messagebox.showerror("Lỗi", "Trang bắt đầu phải là số nguyên ≥ 1.")
            return

        if not username or not password or not code:
            messagebox.showerror("Thiếu thông tin", "Vui lòng nhập đủ Username, Password và CODE.")
            return

        btn_run.config(state="disabled")
        logger.log("=== BẮT ĐẦU CHẠY ===")
        th = threading.Thread(target=lambda: [run_bot(username, password, code, start_page, logger), btn_run.after(0, lambda: btn_run.config(state="normal"))], daemon=True)
        th.start()

    btn_run.configure(command=on_run)

    root.mainloop()

if __name__ == "__main__":
    main()