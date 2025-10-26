import time, traceback, threading, sys, re
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

# ================== CẤU HÌNH TRẠNG THÁI ==================
# Sửa value cho khớp dropdown trạng thái của hệ thống bạn nếu khác
STATUS_MAP = {
    "CHUA_XU_LY": ("Chưa xử lý", "0"),
    "DA_DUYET_KHONG_XL": ("Đã duyệt, không xử lý", "3"),
}

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
def _is_enabled_vakata_item(driver, a_el):
    try:
        cls = (a_el.get_attribute("class") or "").lower()
        aria = (a_el.get_attribute("aria-disabled") or "").lower()
        li = a_el.find_element(By.XPATH, "./ancestor::li[1]")
        li_cls = (li.get_attribute("class") or "").lower()
        return ("disabled" not in cls) and ("disabled" not in li_cls) and (aria not in ["true", "1"])
    except Exception:
        return False

def _ensure_node_selected(driver, anchor):
    try:
        li = anchor.find_element(By.XPATH, "./ancestor::li[1]")
        selected = "jstree-clicked" in (anchor.get_attribute("class") or "")
        if not selected:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor)
            try:
                anchor.click()
            except Exception:
                driver.execute_script("arguments[0].click();", anchor)
            WebDriverWait(driver, 3).until(
                lambda d: "jstree-clicked" in (anchor.get_attribute("class") or "")
                         or "jstree-selected" in (li.get_attribute("class") or "")
            )
    except Exception:
        pass

def _hard_close_all_context_menus(driver):
    try:
        driver.execute_script("document.body.click();")
    except Exception:
        pass
    try:
        try:
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except Exception:
            pass
        driver.execute_script("""
            try {
                document.querySelectorAll('ul.vakata-context').forEach(ul=>{
                    ul.style.display='none';
                    ul.classList.add('vakata-context-hidden');
                });
                document.querySelectorAll('.jquery-loading-modal__bg,.modal-backdrop,.swal2-container')
                    .forEach(el=>{ el.style.pointerEvents='none'; });
            } catch(e){}
        """)
    except Exception:
        pass

def _open_context_menu(driver, anchor):
    _hard_close_all_context_menus(driver)
    try:
        if "jstree-clicked" not in (anchor.get_attribute("class") or ""):
            try:
                anchor.click()
            except Exception:
                driver.execute_script("arguments[0].click();", anchor)
            time.sleep(0.05)
    except Exception:
        pass
    ActionChains(driver).move_to_element(anchor).pause(0.05).context_click(anchor).perform()
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((
        By.XPATH,
        "//ul[contains(@class,'vakata-context')][contains(@style,'display') and not(contains(@style,'display: none'))]"
    )))

def context_click_when_enabled(driver, anchor, rel=None, label=None,
                               logger=None, modal=None):
    _ensure_node_selected(driver, anchor)
    try:
        _open_context_menu(driver, anchor)
        menu = driver.find_element(By.XPATH,
            "//ul[contains(@class,'vakata-context')][contains(@style,'display') and not(contains(@style,'display: none'))]"
        )
        item = None
        if rel is not None:
            els = menu.find_elements(By.CSS_SELECTOR, f"a[rel='{rel}']")
            if els: item = els[0]
        if item is None and label:
            els = menu.find_elements(By.XPATH, f".//a[normalize-space()='{label}']") or \
                  menu.find_elements(By.XPATH, f".//a[contains(normalize-space(.), '{label}')]")
            if els: item = els[0]

        if item and _is_enabled_vakata_item(driver, item):
            if logger: logger(f"   ✓ Menu '{label or rel}' đã bật, click.")
            try:
                item.click()
            except (ElementClickInterceptedException, ElementNotInteractableException):
                driver.execute_script("arguments[0].click();", item)
            return True
        else:
            if logger: logger(f"   ⚠️ Menu '{label or rel}' bị disabled/không thấy. Thử nudge Next→Back.")
            _hard_close_all_context_menus(driver)
            if modal is not None:
                if nudge_by_next_back(driver, modal, logger=logger):
                    if logger: logger("   (Đã nudge xong, thử lại thao tác một lần cuối.)")
                    try:
                        module = modal.find_element(By.CSS_SELECTOR, "#vModuleThiCong[vmodule-name='xulydondangky']")
                        tree = wait_jstree_ready_in(module, timeout=10)
                        refreshed_anchor = find_tt_dangky_anchor(tree)
                        _ensure_node_selected(driver, refreshed_anchor)
                        _open_context_menu(driver, refreshed_anchor)
                        menu = driver.find_element(By.XPATH,
                            "//ul[contains(@class,'vakata-context')][contains(@style,'display') and not(contains(@style,'display: none'))]"
                        )
                        item = None
                        if rel is not None:
                            els = menu.find_elements(By.CSS_SELECTOR, f"a[rel='{rel}']")
                            if els: item = els[0]
                        if item is None and label:
                            els = menu.find_elements(By.XPATH, f".//a[normalize-space()='{label}']")
                            if els: item = els[0]
                        if item and _is_enabled_vakata_item(driver, item):
                            if logger: logger("   ✓ Menu đã bật sau nudge! Click.")
                            item.click()
                            return True
                        else:
                            if logger: logger("   (Menu vẫn không bật sau nudge.)")
                    except Exception as e:
                        if logger: logger(f"   (Lỗi thử lại sau nudge: {e.__class__.__name__})")
            return False

    except Exception as e:
        if logger: logger(f"   ❌ Lỗi context menu: {e.__class__.__name__}")
        _hard_close_all_context_menus(driver)
        return False

def dismiss_jconfirm(driver, timeout=2.0):
    """
    Tìm và bấm 'Đồng ý' trên jQuery-Confirm (jconfirm), trả về True nếu đã bấm.
    """
    driver.switch_to.default_content()
    end = time.time() + timeout
    clicked = False
    while time.time() < end:
        try:
            boxes = driver.find_elements(By.CSS_SELECTOR, ".jconfirm-box-container")
            visible = []
            for b in boxes:
                try:
                    vis = driver.execute_script("""
                        const el = arguments[0];
                        if(!el) return false;
                        const s = getComputedStyle(el);
                        return s.display!=='none' && s.visibility!=='hidden' && el.offsetParent!==null;
                    """, b)
                    if vis:
                        visible.append(b)
                except Exception:
                    continue

            if not visible:
                break

            for box in visible:
                try:
                    try:
                        btn = box.find_element(By.XPATH, ".//button[normalize-space()='Đồng ý']")
                    except NoSuchElementException:
                        btns = box.find_elements(By.CSS_SELECTOR, ".jconfirm-buttons button, button.btn")
                        btn = btns[0] if btns else None

                    if not btn:
                        continue

                    driver.execute_script("""
                        document.querySelectorAll('.modal-backdrop,.jquery-loading-modal__bg')
                          .forEach(el=>{el.style.pointerEvents='none';el.style.opacity='0';el.style.display='none';});
                    """)
                    try: btn.click()
                    except Exception: driver.execute_script("arguments[0].click();", btn)

                    try: WebDriverWait(driver, 2).until(EC.staleness_of(box))
                    except Exception: pass
                    clicked = True
                except Exception:
                    continue
            if clicked: return True
        except Exception: pass
        time.sleep(0.15)
    return clicked

def wait_xuly_modal(driver, timeout=20):
    wait = WebDriverWait(driver, timeout)
    driver.switch_to.default_content()
    modal = wait.until(EC.visibility_of_element_located((
        By.CSS_SELECTOR, "div.modal.modal-fullscreen.in[id^='mdlXuLyDonDangKy-'][style*='display: block'], div.modal.modal-fullscreen.show[id^='mdlXuLyDonDangKy-']"
    )))
    try:
        WebDriverWait(driver, 5).until(lambda d: d.execute_script("return (window.jQuery? jQuery.active:0)") == 0)
    except Exception:
        pass
    return modal

def wait_jstree_ready_in(container_el, timeout=20):
    end = time.time() + timeout
    while time.time() < end:
        trees = container_el.find_elements(By.CSS_SELECTOR, "#treeDonDangKy")
        if trees:
            anchors = trees[0].find_elements(By.CSS_SELECTOR, "a.jstree-anchor")
            if anchors:
                if not (len(anchors) == 1 and "Không có dữ liệu" in (anchors[0].text or "")):
                    return trees[0]
        time.sleep(0.2)
    raise TimeoutException("jsTree chưa có dữ liệu trong thời gian cho phép.")

def find_tt_dangky_anchor(tree_el):
    xpaths = [
        ".//a[.//b[normalize-space()='Thông tin đăng ký']]",
        ".//a[normalize-space()='Thông tin đăng ký']",
        ".//a[contains(normalize-space(.), 'Thông tin đăng ký')]",
    ]
    for xp in xpaths:
        els = tree_el.find_elements(By.XPATH, xp)
        if els:
            return els[0]
    raise NoSuchElementException("Không tìm thấy anchor 'Thông tin đăng ký'.")

def wait_page_idle(driver, wait, extra_ms=300):
    wait.until(lambda x: x.execute_script("return document.readyState") == "complete")
    time.sleep(extra_ms/1000.0)

def switch_to_iframe_containing_table(driver, table_id="tblTTThuaDat", timeout=10):
    driver.switch_to.default_content()
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    deadline = time.time() + timeout
    for idx in range(len(iframes)):
        if time.time() > deadline:
            break
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        try:
            driver.switch_to.frame(iframes[idx])
            if driver.find_elements(By.CSS_SELECTOR, f"#{table_id}"):
                return True
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

def wait_datatable_reload(driver, table_id="tblDanhSachGoiTinDongBo_info", timeout=10):
    """Đợi DataTable load xong sau khi bấm 'Tìm kiếm'."""
    driver.switch_to.default_content()
    try:
        # 1️⃣ Đợi phần tử _processing xuất hiện
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, f"{table_id}_processing"))
        )
    except TimeoutException:
        pass

    # 2️⃣ Rồi đợi nó biến mất
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.ID, f"{table_id}_processing"))
        )
    except TimeoutException:
        pass

    # 3️⃣ Đợi vài dòng đầu tiên xuất hiện
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, f"#{table_id} tbody tr"))
        )
    except TimeoutException:
        pass

    # 4️⃣ Delay nhỏ để đảm bảo text trong _info được cập nhật
    time.sleep(0.5)


def handle_whole_page_action(driver, logger: UILogger, table_id="tblTTThuaDat", timeout=15):
    wait = WebDriverWait(driver, timeout)
    wait.until(EC.presence_of_element_located((By.ID, table_id)))
    rows = driver.find_elements(By.CSS_SELECTOR, f"#{table_id} tbody > tr:not(.child)")
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
        first_target.click()
        if len(visible_rows) > 1:
            ensure_visible_and_sized(last_target)
            ActionChains(driver).key_down(Keys.SHIFT).click(last_target).key_up(Keys.SHIFT).perform()
    except Exception as e:
        logger.log(f"   (Lỗi Shift-Click, thử fallback… Lỗi: {e})")
        for row, tds in visible_rows:
            try:
                target = pick_click_target(row, tds)
                ensure_visible_and_sized(target)
                target.click()
            except Exception:
                continue

    logger.log("   → Đã chọn tất cả, bắt đầu lọc bỏ những bản ghi đã duyệt…")
    time.sleep(0.2)
    actions = ActionChains(driver).key_down(Keys.CONTROL)
    deselected_count = 0
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

def quick_confirm_if_present(driver, root_el=None, soft_timeout=1.2):
    try:
        scope = root_el if root_el is not None else driver
        btns = scope.find_elements(By.CSS_SELECTOR, ".swal2-container .swal2-confirm")
        if not btns:
            btns = scope.find_elements(By.CSS_SELECTOR, ".modal.in .btn-primary, .modal.show .btn-primary")
        if not btns:
            xp = ".//button[normalize-space()='Đồng ý' or normalize-space()='Xác nhận' or normalize-space()='OK' or normalize-space()='Có' or normalize-space()='Yes']"
            try:
                btns = scope.find_elements(By.XPATH, xp)
            except Exception:
                btns = []
        if not btns:
            return False
        cand = None
        for b in btns:
            try:
                vis = driver.execute_script("""
                    const el = arguments[0];
                    const r = el.getBoundingClientRect();
                    const s = getComputedStyle(el);
                    return r.width>0 && r.height>0 && s.visibility!=='hidden' && s.display!=='none';
                """, b)
                if vis:
                    cand = b
                    break
            except Exception:
                continue
        if cand is None:
            return False
        try:
            driver.execute_script("""
                document.querySelectorAll('.modal-backdrop, .swal2-container, .jquery-loading-modal__bg')
                    .forEach(el=>{ el.style.pointerEvents='auto'; });
            """)
        except Exception:
            pass
        try:
            cand.click()
            return True
        except Exception:
            pass
        try:
            driver.execute_script("arguments[0].click();", cand)
            return True
        except Exception:
            pass
        try:
            driver.switch_to.active_element.send_keys(Keys.ENTER)
            return True
        except Exception:
            pass
        return False
    except Exception:
        return False

def quick_ack_any_popup(driver, root_el=None):
    """
    Hợp nhất xác nhận: swal2, bootstrap modal, và jconfirm.
    """
    ok1 = quick_confirm_if_present(driver, root_el=root_el, soft_timeout=0.8)
    ok2 = dismiss_jconfirm(driver, timeout=0.8)
    return ok1 or ok2

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
    cur = driver.execute_script(f"""
        try {{
            return jQuery('#{table_id}').DataTable().page.info().page + 1;
        }} catch(e) {{ return 1; }}
    """) or 1
    if page_number == cur:
        return True
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
    _ = quick_confirm_if_present(driver, soft_timeout=1.0)
    wait_processing_quick(driver, table_id=table_id, max_wait=6)

def switch_to_frame_having(driver, by, value, timeout=8):
    driver.switch_to.default_content()
    try:
        if driver.find_elements(by, value):
            return True
    except Exception:
        pass
    frames = driver.find_elements(By.TAG_NAME, "iframe")
    deadline = time.time() + timeout
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
        except Exception:
            continue
    driver.switch_to.default_content()
    return False

def context_click_jstree_pick(driver, wait, node_text: str,
                              menu_text: str, logger: UILogger = None):
    anchor_xpath = f"//a[contains(@class, 'jstree-anchor') and normalize-space(.)='{node_text}']"
    if not switch_to_frame_having(driver, By.XPATH, anchor_xpath, timeout=8):
        switched = switch_to_frame_having(driver, By.CLASS_NAME, "jstree-anchor", timeout=5)
        if not switched and logger:
            logger.log(f"⚠️ Không tìm thấy iframe chứa jstree hoặc node '{node_text}'.")
    try:
        anchor_for_script = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, anchor_xpath)))
        driver.execute_script(""" 
            try {
                var a = arguments[0];
                if (!a) return;
                var li = a.closest('li');
                var tree = li.closest('.jstree, .jstree-default, .jstree-container-ul');
                var inst = (window.jQuery && tree) ? jQuery(tree).jstree(true) : null;
                if (inst) { inst.open_node(li, null, true); }
            } catch(e) {}
        """, anchor_for_script)
        time.sleep(0.5)
    except TimeoutException:
        if logger:
            logger.log(f"   (Không tìm thấy node '{node_text}' để mở rộng, có thể nó đã hiển thị)")
    try:
        anchor = wait.until(EC.visibility_of_element_located((By.XPATH, anchor_xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", anchor)
        wait.until(EC.element_to_be_clickable((By.XPATH, anchor_xpath)))
        ActionChains(driver).context_click(anchor).perform()
    except TimeoutException as e:
        if logger:
            logger.log(f"❌ Không thể tìm thấy/tương tác node '{node_text}'.")
        raise e
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "ul.vakata-context")))
    item_xpath = f"//ul[contains(@class,'vakata-context')]//a[normalize-space(.)='{menu_text}']"
    try:
        item = wait.until(EC.element_to_be_clickable((By.XPATH, item_xpath)))
        item.click()
    except (TimeoutException, ElementClickInterceptedException) as e:
        if logger:
            logger.log(f"   (Không thể click trực tiếp menu '{menu_text}', thử Javascript. Lỗi: {e.__class__.__name__})")
        item = wait.until(EC.presence_of_element_located((By.XPATH, item_xpath)))
        driver.execute_script("arguments[0].click();", item)

def extract_ma_don_from_tree(tree_el):
    try:
        el = tree_el.find_element(By.XPATH, ".//a[starts-with(normalize-space(.), 'Mã đơn:')]")
        return (el.text or "").strip()
    except Exception:
        try:
            return (tree_el.text or "").strip()
        except Exception:
            return ""

def click_step_forward(modal):
    try:
        btn = modal.find_element(By.ID, "btnStepForward")
    except NoSuchElementException:
        return False
    try:
        dis_attr = btn.get_attribute("disabled")
        cls = btn.get_attribute("class") or ""
        if (dis_attr is not None) or ("disabled" in cls.lower()):
            return False
    except Exception:
        pass
    try:
        modal.parent.execute_script("arguments[0].click();", btn)
    except Exception:
        try:
            btn.click()
        except Exception:
            return False
    return True

def click_step_backward(modal):
    try:
        btn = modal.find_element(By.ID, "btnStepBackward")
    except NoSuchElementException:
        return False
    try:
        dis_attr = btn.get_attribute("disabled")
        cls = (btn.get_attribute("class") or "").lower()
        if (dis_attr is not None) or ("disabled" in cls):
            return False
    except Exception:
        pass
    try:
        modal.parent.execute_script("arguments[0].click();", btn)
    except Exception:
        try: btn.click()
        except Exception: return False
    return True

def nudge_by_next_back(driver, modal, logger=None, change_timeout=12):
    def log(m): (logger and logger(m))
    try:
        ma0 = current_ma_don_in_thicong(modal, timeout=8)
    except Exception:
        ma0 = ""
    if click_step_forward(modal):
        try:
            WebDriverWait(driver, change_timeout).until(
                lambda d: (lambda x: x and x != ma0)(current_ma_don_in_thicong(modal, timeout=8))
            )
        except TimeoutException:
            try: click_step_backward(modal)
            except Exception: pass
            return False
        if not click_step_backward(modal):
            return False
        try:
            WebDriverWait(driver, change_timeout).until(
                lambda d: current_ma_don_in_thicong(modal, timeout=8) == ma0
            )
        except TimeoutException:
            return False
        log and log("   ↩️ Nudge Next→Back xong.")
        return True
    if click_step_backward(modal):
        try:
            WebDriverWait(driver, change_timeout).until(
                lambda d: (lambda x: x and x != ma0)(current_ma_don_in_thicong(modal, timeout=8))
            )
        except TimeoutException:
            try: click_step_forward(modal)
            except Exception: pass
            return False
        if not click_step_forward(modal):
            return False
        try:
            WebDriverWait(driver, change_timeout).until(
                lambda d: current_ma_don_in_thicong(modal, timeout=8) == ma0
            )
        except TimeoutException:
            return False
        log and log("   ↩️ Nudge Back→Next xong.")
        return True
    return False

def current_ma_don_in_thicong(modal, timeout=15):
    module_thicong = modal.find_element(By.CSS_SELECTOR, "#vModuleThiCong[vmodule-name='xulydondangky']")
    tree = wait_jstree_ready_in(module_thicong, timeout=timeout)
    return extract_ma_don_from_tree(tree)

def process_current_record(driver, wait, logger, modal, phase_key=None):
    """
    Xử lý 1 hồ sơ trong modal.
    - Nếu phase_key == "DA_DUYET_KHONG_XL": làm thêm bước đầu ở module Vận hành: 'Bỏ duyệt'.
    - Sau đó làm như bình thường:
        B1 (Thi công): 'Thêm vào dữ liệu vận hành'
        B2 (Vận hành): 'Cập nhật lịch sử tất cả'
    Trả về 'RECOVER' nếu gặp popup và cần phục hồi, True nếu thành công, False nếu lỗi.
    """
    try:
        module_thicong  = modal.find_element(By.CSS_SELECTOR, "#vModuleThiCong[vmodule-name='xulydondangky']")
        module_vanhanh  = modal.find_element(By.CSS_SELECTOR, "#vModuleVanHanh[vmodule-name='xulydondangky']")

        # Đợi jsTree sẵn
        tree_thicong = wait_jstree_ready_in(module_thicong, timeout=30)
        tree_vanhanh = wait_jstree_ready_in(module_vanhanh, timeout=30)

        # ====== NHÁNH ĐẶC BIỆT CHO 'ĐÃ DUYỆT, KHÔNG XỬ LÝ' ======
        if phase_key == "DA_DUYET_KHONG_XL":
            logger.log("   - B0 (Vận hành): Bỏ duyệt…")
            try:
                anchor_vh0 = find_tt_dangky_anchor(tree_vanhanh)
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor_vh0)
                ok_boduyet = context_click_when_enabled(
                    driver, anchor_vh0,
                    rel=None, label="Bỏ duyệt",  # menu có id=btnBoDuyetTTTD, text = 'Bỏ duyệt'
                    logger=logger.log, modal=modal
                )
                if ok_boduyet:
                    quick_ack_any_popup(driver, root_el=modal)
                    logger.log("     ✓ Hoàn tất B0: Bỏ duyệt.")
                else:
                    if quick_ack_any_popup(driver, root_el=modal):
                        logger.log("     ↻ Bắt gặp popup ở B0 → đã bấm 'Đồng ý', yêu cầu phục hồi.")
                        return "RECOVER"
                    logger.log("     ❌ Không bấm được 'Bỏ duyệt' (menu disabled/không thấy).")
            except Exception as e:
                logger.log(f"     ❌ Lỗi B0: {e.__class__.__name__}")
            time.sleep(1)

        # ====== B1: THI CÔNG → Thêm vào dữ liệu vận hành ======
        logger.log("   - B1 (Thi công): Thêm vào dữ liệu vận hành…")
        anchor_tc = find_tt_dangky_anchor(tree_thicong)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor_tc)
        ok_tc = context_click_when_enabled(
            driver, anchor_tc,
            rel=2, label="Thêm vào dữ liệu vận hành",
            logger=logger.log, modal=modal
        )
        if ok_tc:
            quick_ack_any_popup(driver, root_el=modal)
            logger.log("     ✓ Hoàn tất B1.")
        else:
            if quick_ack_any_popup(driver, root_el=modal):
                logger.log("     ↻ Bắt gặp popup ở B1 → đã bấm 'Đồng ý', yêu cầu phục hồi.")
                return "RECOVER"
            logger.log("     ❌ Lỗi B1.")
        time.sleep(1)

        # ====== B2: VẬN HÀNH → Cập nhật lịch sử tất cả ======
        logger.log("   - B2 (Vận hành): Cập nhật lịch sử tất cả…")
        anchor_vh = find_tt_dangky_anchor(tree_vanhanh)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor_vh)
        ok_vh = context_click_when_enabled(
            driver, anchor_vh,
            rel=4, label="Cập nhật lịch sử tất cả",
            logger=logger.log, modal=modal
        )
        if ok_vh:
            quick_ack_any_popup(driver, root_el=modal)
            logger.log("     ✓ Hoàn tất B2.")
        else:
            if quick_ack_any_popup(driver, root_el=modal):
                logger.log("     ↻ Bắt gặp popup ở B2 → đã bấm 'Đồng ý', yêu cầu phục hồi.")
                return "RECOVER"
            logger.log("     ❌ Lỗi B2.")
        time.sleep(1)
        return True

    except Exception as e:
        logger.log(f"   🔥 Lỗi process_current_record: {e.__class__.__name__}")
        logger.log(traceback.format_exc())
        return False

def wait_ajax_idle(driver, max_wait=15, check_interval=0.3, log=None):
    start = time.time()
    last_busy = True
    def is_idle():
        try:
            return driver.execute_script("""
                const jqActive = window.jQuery ? jQuery.active : 0;
                const netIdle = (
                    !(window.pendingFetchCount > 0) &&
                    !(window.pendingXHRCount > 0)
                );
                const modals = document.querySelectorAll('.modal-backdrop.in, .swal2-container, .jquery-loading-modal, .loading-overlay, .blockUI');
                const hasVisibleModal = Array.from(modals).some(m => {
                    const s = getComputedStyle(m);
                    return s.display !== 'none' && s.visibility !== 'hidden' && s.opacity !== '0';
                });
                return (jqActive === 0 && netIdle && !hasVisibleModal);
            """)
        except Exception:
            return True
    try:
        driver.execute_script("""
            if (!window.__ajaxHookInstalled) {
                window.__ajaxHookInstalled = true;
                window.pendingFetchCount = 0;
                window.pendingXHRCount = 0;
                const origFetch = window.fetch;
                if (origFetch) {
                    window.fetch = function(...args) {
                        window.pendingFetchCount++;
                        return origFetch(...args).finally(() => { window.pendingFetchCount--; });
                    };
                }
                const origOpen = XMLHttpRequest.prototype.open;
                const origSend = XMLHttpRequest.prototype.send;
                XMLHttpRequest.prototype.open = function(...args) {
                    this.__ajax = true;
                    return origOpen.apply(this, args);
                };
                XMLHttpRequest.prototype.send = function(...args) {
                    if (this.__ajax) window.pendingXHRCount++;
                    this.addEventListener('loadend', () => window.pendingXHRCount--);
                    return origSend.apply(this, args);
                };
            }
        """)
    except Exception:
        pass
    while time.time() - start < max_wait:
        idle = is_idle()
        if idle:
            if last_busy:
                last_busy = False
                idle_since = time.time()
            elif time.time() - idle_since >= 0.8:
                return True
        else:
            last_busy = True
        time.sleep(check_interval)
    if log:
        log("⚠️ Hết thời gian chờ AJAX.")
    raise TimeoutException("AJAX requests không idle sau {:.1f}s".format(max_wait))

def _get_tt_anchor_and_key(modal, timeout=15):
    module_thicong = modal.find_element(By.CSS_SELECTOR, "#vModuleThiCong[vmodule-name='xulydondangky']")
    tree = wait_jstree_ready_in(module_thicong, timeout=timeout)
    anchor = find_tt_dangky_anchor(tree)
    li = anchor.find_element(By.XPATH, "./ancestor::li[1]")
    key = (li.get_attribute("id") or li.get_attribute("data-id") or (anchor.text or "")).strip()
    return anchor, key

def _wait_switched_to_new_record(driver, modal, old_anchor, old_key, timeout=12):
    end = time.time() + timeout
    try:
        WebDriverWait(driver, 3).until(EC.staleness_of(old_anchor))
        return True
    except Exception:
        pass
    while time.time() < end:
        try:
            _, new_key = _get_tt_anchor_and_key(modal, timeout=5)
            if new_key and new_key != old_key:
                return True
        except Exception:
            pass
        time.sleep(0.2)
    return False

# ============== FILTER HELPERS ==============
def open_filter_and_select_status(driver, wait, status_value: str, logger=None, table_id="formSearchDSDDK"):
    """Mở 'Tìm kiếm mở rộng' + set trạng thái + bấm 'Tìm kiếm'."""
    log = (logger.log if logger else print)
    log(f"🔍 Mở bộ lọc & chọn trạng thái value='{status_value}'…")
    btn_filter = wait.until(EC.element_to_be_clickable((By.ID, "btnShowFormSearchDSDDK")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_filter)
    try: btn_filter.click()
    except Exception: driver.execute_script("arguments[0].click();", btn_filter)

    time.sleep(1.0)

    # Set trạng thái
    try:
        driver.execute_script("""
            const sel = document.querySelector("select[name='trangThai']");
            if (sel) {
                sel.value = arguments[0];
                sel.dispatchEvent(new Event('change', {bubbles:true}));
                if (window.jQuery) jQuery(sel).trigger('change');
            }
        """, status_value)
    except Exception:
        pass

    except Exception:
        pass
    try:
        select_element = driver.find_element(By.CSS_SELECTOR, '#formSearchDSDDK > div > div > fieldset > div > div:nth-child(2) > div > div:nth-child(3) > div > select')
        dropdown = Select(select_element)
        dropdown.select_by_value(status_value)
    except Exception:
        pass

    # Bấm Tìm kiếm
    log("🔎 Bấm 'Tìm kiếm'…")
    btn_search = wait.until(EC.element_to_be_clickable((By.ID, "btnTimKiemMoRong")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_search)
    try: btn_search.click()
    except Exception: driver.execute_script("arguments[0].click();", btn_search)
    wait_datatable_reload(driver, table_id="tblDanhSachGoiTinDongBo_info", timeout=10)

    # Chờ bảng load lại
    try:
        old_row = driver.find_element(By.CSS_SELECTOR, f"#{table_id} tbody tr")
        WebDriverWait(driver, 10).until(EC.staleness_of(old_row))
    except Exception:
        pass
    try:
        wait_for_table_loaded(driver, table_id=table_id, timeout=20)
    except Exception:
        pass
    btn_filter = wait.until(EC.element_to_be_clickable((By.ID, "btnShowFormSearchDSDDK")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_filter)
    try: btn_filter.click()
    except Exception: driver.execute_script("arguments[0].click();", btn_filter)
    log("✅ Đã áp dụng bộ lọc & tải danh sách.")

def resubmit_current_search_if_present(driver, wait, logger=None):
    """Bấm nút 'Tìm kiếm' để refresh, và chờ cho đến khi bảng được tải lại hoàn toàn."""
    try:
        # 1. Tìm nút tìm kiếm
        btn = wait.until(EC.presence_of_element_located((By.ID, "btnTimKiemMoRong")))
        vis_ok = driver.execute_script("""
            const b = arguments[0]; const r = b.getBoundingClientRect();
            const s = getComputedStyle(b);
            return r.width>0 && r.height>0 && s.display!=='none' && s.visibility!=='hidden';
        """, btn)

        if vis_ok:
            # 2. Ghi lại hàng đầu tiên của bảng cũ
            try:
                old_first_row = driver.find_element(By.CSS_SELECTOR, "#tblDanhSachGoiTinDongBo tbody tr")
            except Exception:
                old_first_row = None

            # 3. Bấm tìm kiếm
            if logger: logger.log("🔁 Bấm lại 'Tìm kiếm' (refresh dữ liệu, không mở form).")
            try: btn.click()
            except Exception: driver.execute_script("arguments[0].click();", btn)

            # 4. Chờ bảng cũ biến mất và bảng mới load xong
            if old_first_row:
                WebDriverWait(driver, 15).until(EC.staleness_of(old_first_row))
            wait_for_table_loaded(driver, table_id="tblDanhSachGoiTinDongBo", timeout=20)
            return True
    except Exception:
        pass
    return False

def get_total_records_from_info(driver, info_id="tblDanhSachGoiTinDongBo_info", timeout=10):
    def _normalize_int(s: str) -> int:
        digits = re.sub(r'\D', '', s)
        return int(digits) if digits else 0
    def _locate_in_any_frame(by, value) -> bool:
        driver.switch_to.default_content()
        try:
            WebDriverWait(driver, 0.5).until(EC.presence_of_element_located((by, value)))
            return True
        except TimeoutException:
            pass
        for fr in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(fr)
                WebDriverWait(driver, 0.5).until(EC.presence_of_element_located((by, value)))
                return True
            except TimeoutException:
                continue
        driver.switch_to.default_content()
        return False
    try:
        if not _locate_in_any_frame(By.ID, info_id):
            return 0
        info_el = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.ID, info_id))
        )
        table_id = info_id[:-5] if info_id.endswith("_info") else None
        if table_id:
            try:
                driver.find_element(By.ID, f"{table_id}_processing")
                WebDriverWait(driver, timeout).until(
                    EC.invisibility_of_element_located((By.ID, f"{table_id}_processing"))
                )
            except Exception:
                pass
        text = info_el.text or ""
        print(f"DEBUG: get_total_records_from_info reading text: '{text}'") # Added debug log
        m = re.search(r"tổng\s*số\s+([\d\.,\s]+)", text, flags=re.IGNORECASE)
        if m:
            return _normalize_int(m.group(1))
        nums = re.findall(r"[\d\.,]+", text)
        return _normalize_int(nums[-1]) if nums else 0
    except Exception:
        return 0

# ============== PHỤC HỒI KHI NEXT LỖI ==============
def close_xuly_modal_if_open(driver, soft_timeout=6):
    try:
        driver.switch_to.default_content()
        modal = WebDriverWait(driver, soft_timeout).until(EC.presence_of_element_located((
            By.CSS_SELECTOR, "div.modal.in[id^='mdlXuLyDonDangKy-'], div.modal.show[id^='mdlXuLyDonDangKy-']"
        )))
    except TimeoutException:
        return False
    try:
        close_btn = modal.find_element(By.CSS_SELECTOR, ".panel-heading .close, .panel-heading > button")
        try: close_btn.click()
        except Exception: modal.parent.execute_script("arguments[0].click();", close_btn)
    except Exception:
        try: driver.execute_script("document.body.click();")
        except Exception: pass
    try:
        WebDriverWait(driver, soft_timeout).until(EC.invisibility_of_element_located((
            By.CSS_SELECTOR, "div.modal.in[id^='mdlXuLyDonDangKy-'], div.modal.show[id^='mdlXuLyDonDangKy-']"
        )))
        return True
    except TimeoutException:
        return False

def reopen_modal_from_list(driver, wait, logger=None, timeout=20):
    if logger: logger.log("🗂️ Mở lại modal 'Xử lý đơn đăng ký'…")
    btn = wait.until(EC.presence_of_element_located((By.ID, "btnXuLyDon")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    try: btn.click()
    except Exception: driver.execute_script("arguments[0].click();", btn)
    return wait_xuly_modal(driver, timeout=timeout)

def recover_close_resubmit_reopen(driver, wait, logger):
    """
    Khi Next bị kẹt hoặc không đổi hồ sơ:
      1) Đóng modal
      2) (Nếu có) bấm lại 'Tìm kiếm' để refresh danh sách
      3) Đọc lại tổng số hồ sơ 'Chưa xử lý' hoặc 'Đã duyệt, không xử lý' (tuỳ giai đoạn)
      4) Nếu còn hồ sơ → bấm 'Xử lý đơn' mở lại modal
    Trả về (modal_moi hoặc None, remaining_after).
    """
    ok = close_xuly_modal_if_open(driver, soft_timeout=6)
    logger and logger.log("   ↪️ Đóng modal " + ("✓" if ok else "✗ (không mở/đóng không được)"))

    _ = resubmit_current_search_if_present(driver, wait, logger)
    wait_datatable_reload(driver, table_id="tblDanhSachGoiTinDongBo_info", timeout=10)
    remaining_after = get_total_records_from_info(driver, info_id="tblDanhSachGoiTinDongBo_info")
    logger and logger.log(f"   📊 Sau refresh: còn {remaining_after if remaining_after else 0} hồ sơ trong danh sách hiện tại.")

    if remaining_after == 0:
        logger and logger.log("✅ Không còn hồ sơ nào trong danh sách hiện tại.")
        return None, 0

    try:
        modal = reopen_modal_from_list(driver, wait, logger, timeout=45)
        logger and logger.log("   🔄 Đã mở lại modal sau refresh.")
        return modal, remaining_after
    except Exception as e:
        logger and logger.log(f"   ⛔ Không thể mở lại modal: {e.__class__.__name__}")
        return None, remaining_after
    
def _is_xuly_modal_open(driver):
    try:
        driver.switch_to.default_content()
        el = driver.find_element(By.CSS_SELECTOR, "div.modal.modal-fullscreen[id^='mdlXuLyDonDangKy-']")
        style_ok = driver.execute_script("""
            const el = arguments[0], s = getComputedStyle(el);
            return s.display !== 'none' && s.visibility !== 'hidden' && el.offsetParent !== null;
        """, el)
        return bool(style_ok)
    except Exception:
        return False

def ensure_close_xuly_modal(driver, hard_timeout=5):
    """
    Đóng modal Xử lý đơn một cách *chắc chắn* (kể cả khi nút close bị che hoặc overlay).
    """
    driver.switch_to.default_content()
    end = time.time() + hard_timeout

    def try_close_once():
        try:
            modal = driver.find_element(By.CSS_SELECTOR, "div.modal.modal-fullscreen[id^='mdlXuLyDonDangKy-']")
        except NoSuchElementException:
            return True  # không còn modal trong DOM
        # 1) click nút close trên heading nếu có
        try:
            btn = modal.find_element(By.CSS_SELECTOR, ".panel-heading .close, .panel-heading > button")
            try: btn.click()
            except Exception: driver.execute_script("arguments[0].click();", btn)
        except Exception:
            pass
        # 2) gửi ESC
        try:
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except Exception:
            pass
        # 3) tắt backdrop/overlay cứng rồi click ra body
        try:
            driver.execute_script("""
                document.querySelectorAll('.modal-backdrop,.swal2-container,.jquery-loading-modal__bg')
                  .forEach(el=>{ el.style.pointerEvents='none'; el.style.opacity='0'; el.style.display='none'; });
                document.body.click();
            """)
        except Exception:
            pass
        # 4) modal bootstrap: gọi .modal('hide') nếu có jQuery
        try:
            driver.execute_script("""
                const m = document.querySelector("div.modal.modal-fullscreen[id^='mdlXuLyDonDangKy-']");
                if (m && window.jQuery) jQuery(m).modal('hide');
            """)
        except Exception:
            pass

        # kiểm tra đã ẩn chưa
        try:
            return WebDriverWait(driver, 1.5).until(EC.invisibility_of_element_located((
                By.CSS_SELECTOR, "div.modal.modal-fullscreen[id^='mdlXuLyDonDangKy-']"
            )))
        except Exception:
            return False

    # loop ngắn để đảm bảo tắt
    while time.time() < end:
        if try_close_once():
            return True
        time.sleep(0.2)
    return not _is_xuly_modal_open(driver)

# ============== BOT CORE (GỒM 2 GIAI ĐOẠN) ==============
def run_phase(driver, wait, logger, phase_name: str, status_value: str, phase_key: str, progress_cb=None):
    logger.log(f"=== BẮT ĐẦU GIAI ĐOẠN: {phase_name} ===")

    # 1) Áp dụng filter 1 lần cho phase này
    open_filter_and_select_status(driver, wait, status_value, logger=logger, table_id="formSearchDSDDK")
    time.sleep(0.6)

    # 2) Lấy tổng ban đầu
    phase_total = get_total_records_from_info(driver, info_id="tblDanhSachGoiTinDongBo_info")
    if phase_total > 0:
        logger.log(f"📊 Tổng '{phase_name}' ban đầu: {phase_total}.")
        if progress_cb: progress_cb(f"{phase_name}: 0/{phase_total}")
    else:
        logger.log(f"ℹ️ Không có bản ghi cho '{phase_name}'. Bỏ qua giai đoạn này.")
        if progress_cb: progress_cb(f"{phase_name}: 0/0")
        return 0

    # 3) Mở modal xử lý
    btn_xu_ly = wait.until(EC.presence_of_element_located((By.ID, "btnXuLyDon")))
    driver.execute_script("arguments[0].scrollIntoView(true);", btn_xu_ly)
    driver.execute_script("arguments[0].click();", btn_xu_ly)
    logger.log("⏳ Đợi modal 'Xử lý đơn đăng ký'…")
    modal = wait_xuly_modal(driver, timeout=25)

    phase_processed = 0

    while True:
        # Lấy thông tin hiện tại chỉ để log
        try:
            module_thicong = modal.find_element(By.CSS_SELECTOR, "#vModuleThiCong[vmodule-name='xulydondangky']")
            tree_thicong = wait_jstree_ready_in(module_thicong, timeout=20)
            current_ma_don = extract_ma_don_from_tree(tree_thicong)
        except Exception:
            current_ma_don = ""

        logger.log(f"— Hồ sơ {phase_processed + 1}/{phase_total}: {current_ma_don or '(không rõ)'}")

        # Giữ key để kiểm tra chuyển hồ sơ (nếu còn cần Next)
        try:
            old_anchor, old_key = _get_tt_anchor_and_key(modal, timeout=8)
        except Exception:
            old_anchor, old_key = None, ""

        # ===== XỬ LÝ HỒ SƠ =====
        process_status = None
        try:
            process_status = process_current_record(driver, wait, logger, modal, phase_key=phase_key)
        except Exception as e:
            logger.log(f"   ⚠️ Lỗi process_current_record: {e.__class__.__name__}: {e}. Vẫn tiếp tục…")

        # Nếu xử lý gặp popup và yêu cầu phục hồi
        if process_status == "RECOVER":
            logger.log("   ↪️ Xử lý yêu cầu phục hồi sau khi gặp popup...")
            # Giảm số đếm đã xử lý vì hồ sơ này sẽ được làm lại sau khi refresh
            phase_processed = max(0, phase_processed -1)
            # Nhảy đến logic phục hồi
            next_clicked = False # Giả lập như không next được

        time.sleep(0.3)

        # Cập nhật đếm
        phase_processed += 1
        logger.log(f"   ✅ Tiến độ {phase_name}: {phase_processed}/{phase_total}")
        if progress_cb: progress_cb(f"{phase_name}: {phase_processed}/{phase_total}")

        # === NEW: đủ số hồ sơ theo bộ đếm → kết thúc phase ngay ===
        if phase_processed >= phase_total:
            logger.log(f"🏁 Đã duyệt đủ {phase_total}/{phase_total} cho '{phase_name}'. Đóng modal & chuyển phase.")
            try:
                ensure_close_xuly_modal(driver, hard_timeout=5)
            except Exception:
                pass
            
            # Mở lại form tìm kiếm để chuẩn bị cho phase sau
            try:
                logger.log("🔍 Mở lại 'Tìm kiếm mở rộng' để chuẩn bị cho phase tiếp theo...")
                btn_filter = wait.until(EC.element_to_be_clickable((By.ID, "btnShowFormSearchDSDDK")))
                form_search = driver.find_element(By.ID, "formSearchDSDDK")
                if not form_search.is_displayed():
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_filter)
                    driver.execute_script("arguments[0].click();", btn_filter)
                    time.sleep(0.5)
            except Exception:
                logger.log("   (Không thể mở lại form tìm kiếm mở rộng.)")
            return phase_processed  # ← nhảy ra để run_bot chạy phase kế tiếp

        # Nếu chưa đủ số lượng theo tổng ban đầu, vẫn cần Next sang bản ghi sau
        next_clicked = False
        for attempt in range(1, 6):
            if click_step_forward(modal):
                next_clicked = True
                break
            logger.log(f"   (Click Next lần {attempt} chưa được, chờ 1s rồi thử lại…)"); time.sleep(1)

        if not next_clicked:
            # Không Next được → đóng modal, refresh danh sách HIỆN TẠI (không mở lại form), đo lại tổng
            logger.log("   ⛔ Không Next được. Đóng modal ⇒ refresh ⇒ đọc lại tổng ⇒ mở 'Xử lý đơn' (giữ nguyên trạng thái).")
            modal2, remaining_after = recover_close_resubmit_reopen(driver, wait, logger)

            if remaining_after == 0:
                # Hết thật: kết thúc phase
                logger.log(f"✅ Đã duyệt xong '{phase_name}'. Tổng đã xử lý (ước lượng): ~{phase_processed}.")
                try: ensure_close_xuly_modal(driver, hard_timeout=5)
                except Exception: pass
                if progress_cb: progress_cb(f"{phase_name}: {phase_processed}/{phase_processed}")
                return phase_processed

            # Còn hồ sơ: reset lại bộ đếm để bắt đầu lại với danh sách mới
            phase_total = remaining_after
            phase_processed = 0
            logger.log(f"   🔄 Reset bộ đếm. Bắt đầu lại từ 0/{phase_total}.")
            if progress_cb: progress_cb(f"{phase_name}: 0/{phase_total}")

            if modal2 is None:
                logger.log("   ⛔ Không mở lại modal được sau refresh. Kết thúc giai đoạn hiện tại.")
                try: ensure_close_xuly_modal(driver, hard_timeout=5)
                except Exception: pass
                return phase_processed

            modal = modal2
            continue

        # (Phần kiểm tra đã chuyển sang hồ sơ mới – giữ nguyên như cũ)
        try:
            wait_ajax_idle(driver, max_wait=1, log=logger.log)
        except Exception:
            pass

        switched = _wait_switched_to_new_record(driver, modal, old_anchor, old_key, timeout=10)
        if not switched:
            try:
                nudge_by_next_back(driver, modal, logger=logger.log)
                switched = _wait_switched_to_new_record(driver, modal, old_anchor, old_key, timeout=6)
            except Exception:
                pass

        if not switched:
            logger.log("   ⛔ Không chuyển được hồ sơ. Đóng modal ⇒ refresh ⇒ đọc lại tổng ⇒ mở 'Xử lý đơn'.")
            modal2, remaining_after = recover_close_resubmit_reopen(driver, wait, logger)

            if remaining_after == 0:
                logger.log(f"✅ Đã duyệt xong '{phase_name}'. Tổng đã xử lý (ước lượng): ~{phase_processed}.")
                try: ensure_close_xuly_modal(driver, hard_timeout=5)
                except Exception: pass
                if progress_cb: progress_cb(f"{phase_name}: {phase_processed}/{phase_processed}")
                return phase_processed

            # Còn hồ sơ: reset lại bộ đếm để bắt đầu lại với danh sách mới
            phase_total = remaining_after
            phase_processed = 0
            logger.log(f"   🔄 Reset bộ đếm. Bắt đầu lại từ 0/{phase_total}.")
            if progress_cb: progress_cb(f"{phase_name}: 0/{phase_total}")

            if modal2 is None:
                logger.log("   ⛔ Không mở lại modal được sau refresh. Kết thúc giai đoạn hiện tại.")
                try: ensure_close_xuly_modal(driver, hard_timeout=5)
                except Exception: pass
                return phase_processed

            modal = modal2
            continue

def run_bot(username, password, code, logger: UILogger, base_url, selected_phase_key):
    driver = None
    try:
        logger.log("🚀 Khởi động Chrome…")
        options = Options()
        options.add_argument("--start-maximized")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 20)

        # 👉 Link theo tỉnh
        driver.get(base_url)
        logger.log(f"🌐 Mở trang: {base_url}")

        # Đăng nhập
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(username)
        driver.find_element(By.NAME, "password").send_keys(password)
        driver.find_element(By.NAME, "password").send_keys(Keys.ENTER)
        logger.log("🔐 Đang đăng nhập…")
        messagebox.showinfo("Xác minh",
                            "Nếu có xác minh thủ công (captcha/SSO), hãy hoàn tất trên trình duyệt rồi bấm OK để tiếp tục.")

        # Tìm CODE đợt bàn giao
        logger.log(f"🔎 Đang tìm kiếm CODE: {code}")
        try:
            old_first_row = driver.find_element(By.CSS_SELECTOR, "#tblTraCuuDotBanGiao tbody tr")
        except Exception:
            old_first_row = None

        code_input = wait.until(EC.presence_of_element_located((By.ID, "txtTraCuuDotBanGiao")))
        code_input.clear()
        code_input.send_keys(code)
        code_input.send_keys(Keys.ENTER)
        
        # Chờ bảng kết quả tải lại một cách đáng tin cậy
        if old_first_row:
            logger.log("   (Chờ bảng kết quả tìm kiếm tải lại...)")
            WebDriverWait(driver, 20).until(EC.staleness_of(old_first_row))
        wait_for_table_loaded(driver, table_id="tblTraCuuDotBanGiao", timeout=20)
        
        # Chọn dòng đầu tiên
        safe_click_row_css(driver, wait)

        # Mở 'Chức năng' → 'Xem danh sách'
        logger.log("📂 Mở 'Chức năng' → 'Xem danh sách'…")
        wait.until(EC.element_to_be_clickable((By.ID, "drop1"))).click()
        wait.until(EC.element_to_be_clickable((By.ID, "btnXemDanhSach"))).click()

        # Nếu người dùng chọn một phase cụ thể, chỉ chạy phase đó 1 lần
        if selected_phase_key != "ALL":
            if selected_phase_key == "CHUA_XU_LY":
                run_phase(driver, wait, logger,
                          phase_name=STATUS_MAP["CHUA_XU_LY"][0],
                          status_value=STATUS_MAP["CHUA_XU_LY"][1],
                          phase_key="CHUA_XU_LY")
            elif selected_phase_key == "DA_DUYET_KHONG_XL":
                run_phase(driver, wait, logger,
                          phase_name=STATUS_MAP["DA_DUYET_KHONG_XL"][0],
                          status_value=STATUS_MAP["DA_DUYET_KHONG_XL"][1],
                          phase_key="DA_DUYET_KHONG_XL")
        else: # Nếu chọn "Tất cả", lặp lại 2 phase cho đến khi hết
            while True:
                total1 = run_phase(driver, wait, logger,
                                phase_name=STATUS_MAP["CHUA_XU_LY"][0],
                                status_value=STATUS_MAP["CHUA_XU_LY"][1],
                                phase_key="CHUA_XU_LY")
                total2 = run_phase(driver, wait, logger,
                                phase_name=STATUS_MAP["DA_DUYET_KHONG_XL"][0],
                                status_value=STATUS_MAP["DA_DUYET_KHONG_XL"][1],
                                phase_key="DA_DUYET_KHONG_XL")
                if total1 == 0 and total2 == 0:
                    logger.log("✅ Cả hai phase đều hết hồ sơ. Dừng hoàn toàn.")
                    break
                logger.log("🔁 Phát hiện còn hồ sơ mới sau phase 2 ⇒ quay lại kiểm tra lại phase 1…")

        logger.log("🏁 Hoàn tất toàn bộ 2 giai đoạn.")

    except Exception as ex:
        logger.log(f"❌ Có lỗi xảy ra: {ex}")
        logger.log(traceback.format_exc())
    finally:
        try:
            if driver:
                driver.quit()
        except Exception:
            pass

# ============== TKINTER UI ==============
def main():
    root = tk.Tk()
    root.title("Tự động duyệt - MPLIS (Đắk Lắk / Phú Yên)")
    root.geometry("780x540")

    frm = ttk.Frame(root, padding=12)
    frm.pack(fill="x")

    ttk.Label(frm, text="Username:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
    ent_user = ttk.Entry(frm, width=36)
    ent_user.grid(row=0, column=1, sticky="w", padx=4, pady=4)
    ent_user.insert(0, "dla.tuannhqt")

    ttk.Label(frm, text="Password:").grid(row=1, column=0, sticky="e", padx=4, pady=4)
    ent_pass = ttk.Entry(frm, width=36, show="•")
    ent_pass.grid(row=1, column=1, sticky="w", padx=4, pady=4)
    ent_pass.insert(0, "NNMT@2025")

    ttk.Label(frm, text="CODE (Mã đợt bàn giao):").grid(row=2, column=0, sticky="e", padx=4, pady=4)
    ent_code = ttk.Entry(frm, width=36)
    ent_code.grid(row=2, column=1, sticky="w", padx=4, pady=4)
    ent_code.insert(0, "8919bd18-177b-48fa-9ae9-9846d359a048")

    ttk.Label(frm, text="Chọn tỉnh:").grid(row=4, column=0, sticky="e", padx=4, pady=4)
    province_cb = ttk.Combobox(frm, state="readonly", width=33)
    province_cb["values"] = ["Đắk Lắk", "Phú Yên"]
    province_cb.current(0)
    province_cb.grid(row=4, column=1, sticky="w", padx=4, pady=4)

    ttk.Label(frm, text="Trạng thái:").grid(row=5, column=0, sticky="e", padx=4, pady=4)
    phase_cb = ttk.Combobox(frm, state="readonly", width=33)
    phase_map = {
        "Tất cả các trạng thái": "ALL",
        f"1. {STATUS_MAP['CHUA_XU_LY'][0]}": "CHUA_XU_LY",
        f"2. {STATUS_MAP['DA_DUYET_KHONG_XL'][0]}": "DA_DUYET_KHONG_XL",
    }
    phase_cb["values"] = list(phase_map.keys())
    phase_cb.current(0)
    phase_cb.grid(row=5, column=1, sticky="w", padx=4, pady=4)

    ttk.Label(frm, text="Trạng thái:").grid(row=5, column=0, sticky="e", padx=4, pady=4)
    phase_cb = ttk.Combobox(frm, state="readonly", width=33)
    phase_map = {
        "Tất cả các trạng thái": "ALL",
        f"1. {STATUS_MAP['CHUA_XU_LY'][0]}": "CHUA_XU_LY",
        f"2. {STATUS_MAP['DA_DUYET_KHONG_XL'][0]}": "DA_DUYET_KHONG_XL",
    }
    phase_cb["values"] = list(phase_map.keys())
    phase_cb.current(0)
    phase_cb.grid(row=5, column=1, sticky="w", padx=4, pady=4)

    btn_run = ttk.Button(frm, text="Chạy tự động")
    btn_run.grid(row=6, column=1, sticky="w", padx=4, pady=8)

    txt = tk.Text(root, height=18, state="disabled", bg="#0f1115", fg="#e5e7eb")
    txt.pack(fill="both", expand=True, padx=12, pady=(0, 12))
    logger = UILogger(txt)

    def on_run():
        username = ent_user.get().strip()
        password = ent_pass.get()
        code = ent_code.get().strip()
        province = province_cb.get()
        selected_phase_name = phase_cb.get()
        selected_phase_key = phase_map[selected_phase_name]

        if not username or not password or not code:
            messagebox.showerror("Thiếu thông tin", "Vui lòng nhập đủ Username, Password và CODE.")
            return

        if province == "Phú Yên":
            base_url = "https://phy.mplis.gov.vn/dc/TichHopDongBoDuLieu/QuanLyDotBanGiao"
        else:
            base_url = "https://dla.mplis.gov.vn/dc/TichHopDongBoDuLieu/QuanLyDotBanGiao"

        btn_run.config(state="disabled")
        logger.log(f"=== BẮT ĐẦU CHẠY ({province}) ===")

        def runner():
            run_bot(username, password, code, logger, base_url, selected_phase_key)
            btn_run.after(0, lambda: btn_run.config(state="normal"))

        threading.Thread(target=runner, daemon=True).start()

    btn_run.configure(command=on_run)
    root.mainloop()

if __name__ == "__main__":
    main()