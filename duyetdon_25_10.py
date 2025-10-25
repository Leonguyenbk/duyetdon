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
def _is_enabled_vakata_item(driver, a_el):
    """Một item bật nếu KHÔNG có class disabled/aria-disabled."""
    try:
        cls = (a_el.get_attribute("class") or "").lower()
        aria = (a_el.get_attribute("aria-disabled") or "").lower()
        li = a_el.find_element(By.XPATH, "./ancestor::li[1]")
        li_cls = (li.get_attribute("class") or "").lower()
        return ("disabled" not in cls) and ("disabled" not in li_cls) and (aria not in ["true", "1"])
    except Exception:
        return False

def _ensure_node_selected(driver, anchor):
    """jsTree yêu cầu node được select thì menu mới bật."""
    try:
        li = anchor.find_element(By.XPATH, "./ancestor::li[1]")
        selected = "jstree-clicked" in (anchor.get_attribute("class") or "")
        if not selected:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor)
            try:
                anchor.click()
            except Exception:
                driver.execute_script("arguments[0].click();", anchor)
            # chờ chọn xong
            WebDriverWait(driver, 3).until(
                lambda d: "jstree-clicked" in (anchor.get_attribute("class") or "")
                         or "jstree-selected" in (li.get_attribute("class") or "")
            )
    except Exception:
        pass

def _open_context_menu(driver, anchor):
    """Mở menu ngữ cảnh ổn định."""
    ActionChains(driver).move_to_element(anchor).pause(0.05).context_click(anchor).perform()
    # menu visible
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((
        By.XPATH, "//ul[contains(@class,'vakata-context')][contains(@style,'display') and not(contains(@style,'display: none'))]"
    )))

def context_click_when_enabled(driver, anchor, rel=None, label=None,
                               logger=None, modal=None):
    """
    Thực hiện chuột phải vào một 'anchor' và click vào một item trong menu ngữ cảnh.
    Rút gọn: chỉ thử 1 lần, nếu item bị disabled thì nudge Next→Back luôn (không lặp lại 10 lần).
    """
    _ensure_node_selected(driver, anchor)

    try:
        # Mở menu
        _open_context_menu(driver, anchor)
        menu = driver.find_element(By.XPATH,
            "//ul[contains(@class,'vakata-context')][contains(@style,'display') and not(contains(@style,'display: none'))]"
        )

        # Tìm item trong menu
        item = None
        if rel is not None:
            els = menu.find_elements(By.CSS_SELECTOR, f"a[rel='{rel}']")
            if els:
                item = els[0]
        if item is None and label:
            els = menu.find_elements(By.XPATH, f".//a[normalize-space()='{label}']") or \
                  menu.find_elements(By.XPATH, f".//a[contains(normalize-space(.), '{label}')]")
            if els:
                item = els[0]

        # Kiểm tra trạng thái item
        if item and _is_enabled_vakata_item(driver, item):
            if logger: logger(f"   ✓ Menu '{label or rel}' đã bật, tiến hành click.")
            try:
                item.click()
            except (ElementClickInterceptedException, ElementNotInteractableException):
                driver.execute_script("arguments[0].click();", item)
            return True
        else:
            if logger:
                logger(f"   ⚠️ Menu '{label or rel}' bị disabled hoặc không thấy. Thử Next→Back để refresh.")
            _hard_close_all_context_menus(driver)

            # Nếu có modal thì thử refresh
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
                            if logger: logger("   ✓ Menu đã bật sau nudge! Click thực hiện.")
                            item.click()
                            return True
                        else:
                            if logger: logger("   (Menu vẫn không bật sau nudge.)")
                    except Exception as e:
                        if logger: logger(f"   (Lỗi khi thử lại sau nudge: {e.__class__.__name__})")

            return False

    except Exception as e:
        if logger: logger(f"   ❌ Lỗi khi thao tác context menu: {e.__class__.__name__}")
        _hard_close_all_context_menus(driver)
        return False

def wait_xuly_modal(driver, timeout=20):
    """
    Đợi modal Xử lý đơn đăng ký hiển thị; trả về WebElement modal.
    Modal có id động bắt đầu bằng 'mdlXuLyDonDangKy-'.
    """
    wait = WebDriverWait(driver, timeout)
    driver.switch_to.default_content()
    modal = wait.until(EC.visibility_of_element_located((
        By.CSS_SELECTOR, "div.modal.modal-fullscreen.in[id^='mdlXuLyDonDangKy-'][style*='display: block']"
    )))
    # đảm bảo body không còn overlay che click
    try:
        WebDriverWait(driver, 5).until(lambda d: d.execute_script("return (window.jQuery? jQuery.active:0)") == 0)
    except Exception:
        pass
    return modal

def wait_jstree_ready_in(container_el, timeout=20):
    """
    Đợi #treeDonDangKy trong container có ít nhất một anchor khác 'Không có dữ liệu'.
    """
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
    """
    Trả về <a> node 'Thông tin đăng ký' (trong đó text ở <b> bên trong).
    Linh hoạt với phần tử phụ như <div id='elementStatus'>.
    """
    xpaths = [
        ".//a[.//b[normalize-space()='Thông tin đăng ký']]",                     # case phổ biến
        ".//a[normalize-space()='Thông tin đăng ký']",                           # đôi khi text flatten
        ".//a[contains(normalize-space(.), 'Thông tin đăng ký')]",               # lỏng
    ]
    for xp in xpaths:
        els = tree_el.find_elements(By.XPATH, xp)
        if els:
            return els[0]
    raise NoSuchElementException("Không tìm thấy anchor 'Thông tin đăng ký' trong jsTree.")


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

def quick_confirm_if_present(driver, root_el=None, soft_timeout=1.2):
    """
    Tìm & bấm nút xác nhận nếu có (SweetAlert2/Bootstrap). KHÔNG raise TimeoutException.
    Trả về True nếu đã bấm xác nhận; False nếu không thấy gì để bấm.
    root_el: nếu truyền modal WebElement, chỉ tìm trong đó (ổn định hơn).
    """
    try:
        scope = root_el if root_el is not None else driver
        sw = WebDriverWait(driver, soft_timeout)

        # 1) SweetAlert2 .swal2-confirm
        btns = scope.find_elements(By.CSS_SELECTOR, ".swal2-container .swal2-confirm")
        if not btns:
            # 2) Bootstrap modal primary
            btns = scope.find_elements(By.CSS_SELECTOR, ".modal.in .btn-primary, .modal.show .btn-primary")

        if not btns:
            # 3) Theo text tiếng Việt/English phổ biến
            xp = ".//button[normalize-space()='Đồng ý' or normalize-space()='Xác nhận' or normalize-space()='OK' or normalize-space()='Có' or normalize-space()='Yes']"
            try:
                btns = scope.find_elements(By.XPATH, xp)
            except Exception:
                btns = []

        if not btns:
            # Không thấy gì → coi như không có confirm
            return False

        # Chọn nút hiển thị được
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

        # Đảm bảo không bị backdrop che
        try:
            driver.execute_script("""
                document.querySelectorAll('.modal-backdrop, .swal2-container, .jquery-loading-modal__bg')
                    .forEach(el=>{ el.style.pointerEvents='auto'; });
            """)
        except Exception:
            pass

        # Thử click thường
        try:
            cand.click()
            return True
        except Exception:
            pass

        # Thử JS click
        try:
            driver.execute_script("arguments[0].click();", cand)
            return True
        except Exception:
            pass

        # Thử phím Enter vào phần tử đang focus/active
        try:
            driver.switch_to.active_element.send_keys(Keys.ENTER)
            return True
        except Exception:
            pass

        return False
    except Exception:
        # Tuyệt đối không để propagate TimeoutException từ waits bên trong
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

def context_click_jstree_pick(driver, wait, node_text: str,
                              menu_text: str, logger: UILogger = None):
    """
    Tìm một node trong cây jstree theo text, nhấp chuột phải và chọn menu.
    """
    anchor_xpath = f"//a[contains(@class, 'jstree-anchor') and normalize-space(.)='{node_text}']"

    # 1) Đảm bảo đang ở frame có node cần tìm
    if not switch_to_frame_having(driver, By.XPATH, anchor_xpath, timeout=8):
        # Nếu không thấy, thử chuyển đến frame bất kỳ có jstree
        switched = switch_to_frame_having(driver, By.CLASS_NAME, "jstree-anchor", timeout=5)
        if not switched and logger:
            logger.log(f"⚠️ Không tìm thấy iframe chứa jstree hoặc node '{node_text}'.")

    # 2) Mở rộng cây thư mục để đảm bảo node nhìn thấy được
    try:
        # Dùng presence_of_element_located để lấy element ngay cả khi nó chưa visible
        anchor_for_script = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, anchor_xpath)))
        driver.execute_script(""" 
            try {
                var a = arguments[0];
                if (!a) return;
                var li = a.closest('li');
                var tree = li.closest('.jstree, .jstree-default, .jstree-container-ul');
                var inst = (window.jQuery && tree) ? jQuery(tree).jstree(true) : null;
                if (inst) {
                    // Mở tất cả các node cha để đảm bảo node con hiển thị
                    inst.open_node(li, null, true); 
                }
            } catch(e) { console.error('Jstree open_node failed:', e); }
        """, anchor_for_script)
        time.sleep(0.5) # Chờ animation mở cây
    except TimeoutException:
        if logger:
            logger.log(f"   (Không tìm thấy node '{node_text}' để mở rộng, có thể nó đã hiển thị hoặc tên node không đúng)")

    # 3) Lấy anchor và thực hiện context click
    try:
        anchor = wait.until(EC.visibility_of_element_located((By.XPATH, anchor_xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", anchor) # Cuộn vào view
        wait.until(EC.element_to_be_clickable((By.XPATH, anchor_xpath))) # Chờ có thể click
        ActionChains(driver).context_click(anchor).perform()
    except TimeoutException as e:
        if logger:
            logger.log(f"❌ Không thể tìm thấy hoặc tương tác với node '{node_text}' sau khi chờ.")
            logger.log("   Gợi ý: Kiểm tra lại tên node, hoặc đảm bảo nó không bị che khuất.")
        raise e # Ném lại lỗi để dừng script

    # 4) Chờ menu vakata hiện + click item theo text
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "ul.vakata-context")))
    # Dùng normalize-space để match chính xác text, tránh lỗi khoảng trắng
    item_xpath = f"//ul[contains(@class,'vakata-context')]//a[normalize-space(.)='{menu_text}']"
    
    try:
        # Thử chờ click được trước
        item = wait.until(EC.element_to_be_clickable((By.XPATH, item_xpath)))
        item.click()
    except (TimeoutException, ElementClickInterceptedException) as e:
        # Nếu không click được, thử tìm sự hiện diện và click bằng JS
        if logger:
            logger.log(f"   (Không thể click trực tiếp menu '{menu_text}', thử click bằng Javascript. Lỗi: {e.__class__.__name__})")
        try:
            # Chờ element có trong DOM, không cần visible hoặc clickable
            item = wait.until(EC.presence_of_element_located((By.XPATH, item_xpath)))
            driver.execute_script("arguments[0].click();", item)
        except TimeoutException:
            if logger:
                logger.log(f"   (Không tìm thấy menu '{menu_text}' ngay cả với Javascript.)")
            raise # Ném lại lỗi gốc

def _hard_close_all_context_menus(driver):
    """Đóng sạch mọi menu vakata + overlay trước khi mở lại."""
    try:
        # 1) click ra ngoài: giảm xác suất menu còn mở
        driver.execute_script("document.body.click();")
    except Exception:
        pass
    try:
        # 2) gửi ESC (đôi khi menu dùng keydown ESC để đóng)
        try:
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except Exception:
            pass
        # 3) ép ẩn mọi menu vakata còn sót + các lớp overlay gây chặn click
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
    """Mở menu ngữ cảnh ổn định, sau khi đóng sạch các menu cũ."""
    _hard_close_all_context_menus(driver)  # ✨ mới thêm
    # chọn node trước khi right-click (jsTree hay yêu cầu selected)
    try:
        if "jstree-clicked" not in (anchor.get_attribute("class") or ""):
            try:
                anchor.click()
            except Exception:
                driver.execute_script("arguments[0].click();", anchor)
            time.sleep(0.05)
    except Exception:
        pass

    # di chuột vào chính giữa anchor rồi context click
    ActionChains(driver).move_to_element(anchor).pause(0.05).context_click(anchor).perform()

    # chờ menu hiện và thực sự "mở"
    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((
        By.XPATH,
        "//ul[contains(@class,'vakata-context')][contains(@style,'display') and not(contains(@style,'display: none'))]"
    )))


def extract_ma_don_from_tree(tree_el):
    """
    Từ #treeDonDangKy, lấy chuỗi 'Mã đơn: ...' (phục vụ so sánh khi chuyển hồ sơ).
    """
    try:
        el = tree_el.find_element(By.XPATH, ".//a[starts-with(normalize-space(.), 'Mã đơn:')]")
        return (el.text or "").strip()
    except Exception:
        # fallback: ráp text toàn cây (ít tin cậy hơn)
        try:
            return (tree_el.text or "").strip()
        except Exception:
            return ""
    
def click_step_backward(modal):
    """Nhấn nút '◀' (btnStepBackward)."""
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
    """
    Thử Next -> chờ đổi hồ sơ -> Back về hồ sơ cũ (hoặc ngược lại nếu không Next được).
    Trả về True nếu đã đi-về thành công (DOM được refresh).
    """
    def log(m): 
        (logger and logger(m))

    try:
        ma0 = current_ma_don_in_thicong(modal, timeout=8)
    except Exception:
        ma0 = ""

    # Ưu tiên Next→Back
    if click_step_forward(modal):
        # đợi đổi hồ sơ
        try:
            WebDriverWait(driver, change_timeout).until(
                lambda d: (lambda x: x and x != ma0)(current_ma_don_in_thicong(modal, timeout=8))
            )
        except TimeoutException:
            # không đổi được → coi như fail
            try: click_step_backward(modal)
            except Exception: pass
            return False

        # quay lại hồ sơ cũ
        if not click_step_backward(modal):
            return False
        # đợi về lại ma0
        try:
            WebDriverWait(driver, change_timeout).until(
                lambda d: current_ma_don_in_thicong(modal, timeout=8) == ma0
            )
        except TimeoutException:
            return False

        log and log("   ↩️ Đã nudge Next→Back để refresh trạng thái.")
        return True

    # Nếu không Next được, thử Back→Next
    if click_step_backward(modal):
        try:
            WebDriverWait(driver, change_timeout).until(
                lambda d: (lambda x: x and x != ma0)(current_ma_don_in_thicong(modal, timeout=8))
            )
        except TimeoutException:
            # quay lại vị trí cũ nếu có thể
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

        log and log("   ↩️ Đã nudge Back→Next để refresh trạng thái.")
        return True

    return False

def current_ma_don_in_thicong(modal, timeout=15):
    """Đọc 'Mã đơn:' hiện tại từ module Thi công."""
    module_thicong = modal.find_element(By.CSS_SELECTOR, "#vModuleThiCong[vmodule-name='xulydondangky']")
    tree = wait_jstree_ready_in(module_thicong, timeout=timeout)
    return extract_ma_don_from_tree(tree)


def process_current_record(driver, wait, logger, modal):
    """
    - Thi công: Thêm vào dữ liệu vận hành (rel=2).
    - Vận hành: Cập nhật lịch sử tất cả (rel=4).
    """
    try:
        # Lấy lại 2 module mỗi vòng (DOM có thể thay đổi sau khi chuyển hồ sơ)
        module_thicong  = modal.find_element(By.CSS_SELECTOR, "#vModuleThiCong[vmodule-name='xulydondangky']")
        module_vanhanh  = modal.find_element(By.CSS_SELECTOR, "#vModuleVanHanh[vmodule-name='xulydondangky']")

        # 1) Đợi jsTree có dữ liệu
        tree_thicong = wait_jstree_ready_in(module_thicong, timeout=30)
        tree_vanhanh = wait_jstree_ready_in(module_vanhanh, timeout=30)

        # --- Bước 1: Thi công → 'Thêm vào dữ liệu vận hành' (rel=2) ---
        logger.log("   - Bước 1: Thêm vào dữ liệu vận hành...")
        anchor_tc = find_tt_dangky_anchor(tree_thicong)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor_tc)
        
        # Sử dụng hàm context click đã được cải tiến
        ok_tc = context_click_when_enabled(
            driver, anchor_tc,
            rel=2, label="Thêm vào dữ liệu vận hành",
            logger=logger.log, modal=modal
        )
        
        if ok_tc:
            quick_confirm_if_present(driver, root_el=modal, soft_timeout=1.8)
            logger.log("     ✓ Hoàn tất bước 1.")
        else:
            # Lỗi đã được log chi tiết trong hàm con, chỉ cần log kết quả cuối
            logger.log("     ❌ Không thể hoàn tất 'Thêm vào dữ liệu vận hành'.")
        
        time.sleep(0.5)  # Chờ một chút trước khi xử lý phần Vận hành

        # --- Bước 2: Vận hành → 'Cập nhật lịch sử tất cả' (rel=4) ---
        logger.log("   - Bước 2: Cập nhật lịch sử tất cả...")
        anchor_vh = find_tt_dangky_anchor(tree_vanhanh)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor_vh)

        ok_vh = context_click_when_enabled(
            driver, anchor_vh,
            rel=4, label="Cập nhật lịch sử tất cả",
            logger=logger.log, modal=modal
        )

        if ok_vh:
            quick_confirm_if_present(driver, root_el=modal, soft_timeout=1.8)
            logger.log("     ✓ Hoàn tất bước 2.")
        else:
            logger.log("     ❌ Không thể hoàn tất 'Cập nhật lịch sử tất cả'.")

        time.sleep(0.5)  # Chờ một chút trước khi bấm bước tiếp
    except Exception as e:
        logger.log(f"   🔥 Lỗi nghiêm trọng trong process_current_record: {e.__class__.__name__}")
        logger.log(traceback.format_exc())


def click_step_forward(modal):
    """
    Nhấn nút '▶' (btnStepForward) trong modal. Trả về False nếu nút bị disable.
    """
    try:
        btn = modal.find_element(By.ID, "btnStepForward")
    except NoSuchElementException:
        return False

    # Nếu bị disable (thuộc tính disabled hoặc class chứa 'disabled')
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


def wait_ajax_idle(driver, max_wait=15, check_interval=0.3, log=None):
    """
    Đợi cho tất cả các request AJAX / fetch / loading overlay kết thúc.
    Hỗ trợ cả jQuery, fetch, axios, và các modal overlay phổ biến.
    - driver: WebDriver
    - max_wait: thời gian tối đa (giây)
    - check_interval: thời gian chờ giữa các lần kiểm tra
    - log: hàm log (nếu có)
    """
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
            return True  # Nếu script lỗi (chưa có jQuery chẳng hạn) → coi như idle

    # Hook global counter cho fetch / XHR nếu chưa có
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
                        return origFetch(...args)
                            .finally(() => window.pendingFetchCount--);
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
            elif time.time() - idle_since >= 0.8:  # ổn định ít nhất 0.8s
                return True
        else:
            last_busy = True
        time.sleep(check_interval)

    if log:
        log("⚠️ Hết thời gian chờ AJAX.")
    raise TimeoutException("AJAX requests không idle sau {:.1f}s".format(max_wait))

def _get_tt_anchor_and_key(modal, timeout=15):
    """
    Lấy anchor 'Thông tin đăng ký' và 1 khóa ổn định từ <li> cha.
    Khóa ưu tiên: li@id > li@data-id > anchor.text.
    """
    module_thicong = modal.find_element(By.CSS_SELECTOR, "#vModuleThiCong[vmodule-name='xulydondangky']")
    tree = wait_jstree_ready_in(module_thicong, timeout=timeout)
    anchor = find_tt_dangky_anchor(tree)
    li = anchor.find_element(By.XPATH, "./ancestor::li[1]")
    key = (li.get_attribute("id") or li.get_attribute("data-id") or (anchor.text or "")).strip()
    return anchor, key

def _wait_switched_to_new_record(driver, modal, old_anchor, old_key, timeout=12):
    """
    Sau khi bấm Next, đợi đến khi:
      - old_anchor trở thành stale (DOM cũ biến mất), HOẶC
      - khóa (key) của node 'Thông tin đăng ký' đổi khác old_key.
    Trả về True nếu đã chuyển; False nếu không.
    """
    end = time.time() + timeout
    # thử nhanh: staleness_of phần tử cũ
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
            # Trong lúc chuyển, modal/cây có thể tạm thời chưa sẵn sàng
            pass
        time.sleep(0.2)
    return False


# ============== BOT CORE ==============
def run_bot(username, password, code, start_page, logger: UILogger, base_url):
    driver = None
    try:
        logger.log("🚀 Khởi động Chrome…")
        options = Options()
        options.add_argument("--start-maximized")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(driver, 20)

        # 👉 Dùng link theo tỉnh được chọn
        driver.get(base_url)
        logger.log(f"🌐 Mở trang: {base_url}")

        # Đăng nhập
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(username)
        driver.find_element(By.NAME, "password").send_keys(password)
        driver.find_element(By.NAME, "password").send_keys(Keys.ENTER)
        logger.log("🔐 Đang đăng nhập…")
        # Nếu có captcha/xác minh thủ công, dừng lại cho người dùng thao tác
        messagebox.showinfo("Xác minh",
                            "Nếu có xác minh thủ công (captcha/SSO), hãy hoàn tất trên trình duyệt rồi bấm OK để tiếp tục.")

        # Tìm CODE đợt bàn giao
        try:
            old_first_row = driver.find_element(By.CSS_SELECTOR, "#tblTraCuuDotBanGiao tbody tr")
        except Exception:
            old_first_row = None

        driver.find_element(By.ID, "txtTraCuuDotBanGiao").send_keys(code)
        driver.find_element(By.ID, "txtTraCuuDotBanGiao").send_keys(Keys.ENTER)
        logger.log(f"🔎 Đang tìm kiếm CODE: {code}")

        if old_first_row:
            WebDriverWait(driver, 20).until(EC.staleness_of(old_first_row))

        # Chọn dòng đầu tiên
        safe_click_row_css(driver, wait)

        # Mở menu Chức năng → Xem danh sách
        logger.log("📂 Mở 'Chức năng' → 'Xem danh sách'…")
        wait.until(EC.element_to_be_clickable((By.ID, "drop1"))).click()
        wait.until(EC.element_to_be_clickable((By.ID, "btnXemDanhSach"))).click()

        # Bấm 'Xử lý đơn'
        btn_xu_ly = wait.until(EC.presence_of_element_located((By.ID, "btnXuLyDon")))
        driver.execute_script("arguments[0].scrollIntoView(true);", btn_xu_ly)
        driver.execute_script("arguments[0].click();", btn_xu_ly)

        logger.log("⏳ Đợi modal 'Xử lý đơn đăng ký' hiển thị…")
        modal = wait_xuly_modal(driver, timeout=25)

        logger.log("🔁 Bắt đầu duyệt tuần tự từng hồ sơ (StepForward)…")

        # Lấy Mã đơn ban đầu (chỉ để log), logic chuyển trang dùng KEY + staleness
        module_thicong = modal.find_element(By.CSS_SELECTOR, "#vModuleThiCong[vmodule-name='xulydondangky']")
        tree_thicong = wait_jstree_ready_in(module_thicong, timeout=30)
        current_ma_don = extract_ma_don_from_tree(tree_thicong)

        index = 1
        while True:
            logger.log(f"— Hồ sơ #{index}: {current_ma_don or '(không rõ)'}")

            # Xử lý hồ sơ hiện tại
            try:
                process_current_record(driver, wait, logger, modal)
            except Exception as e:
                logger.log(f"   ⚠️ Lỗi khi xử lý hồ sơ: {e.__class__.__name__}: {e}. Tiếp tục hồ sơ kế tiếp…")

            # LẤY KHÓA TRƯỚC KHI NEXT: anchor + key ổn định của node 'Thông tin đăng ký'
            try:
                old_anchor, old_key = _get_tt_anchor_and_key(modal, timeout=8)
            except Exception:
                old_anchor, old_key = None, ""

            # Thử sang hồ sơ kế tiếp (tối đa 5 lần click nếu nút chưa sẵn sàng)
            next_clicked = False
            for attempt in range(1, 6):
                if click_step_forward(modal):
                    next_clicked = True
                    break
                logger.log(f"   (Click Next lần {attempt} thất bại, chờ 1s rồi thử lại...)")
                time.sleep(1)

            if not next_clicked:
                logger.log("⛔ Hết hồ sơ (không thể click nút ▶ sau nhiều lần thử). Kết thúc.")
                break

            # Chờ AJAX/overlay lắng bớt (giảm race condition)
            try:
                wait_ajax_idle(driver, max_wait=0.5, log=logger.log)
            except Exception:
                pass

            # Đợi thực sự chuyển sang hồ sơ mới bằng staleness/key thay đổi
            switched = _wait_switched_to_new_record(driver, modal, old_anchor, old_key, timeout=12)
            if not switched:
                # Thử nudge Next→Back rồi chờ đổi lại lần nữa (tùy chọn)
                try:
                    nudge_by_next_back(driver, modal, logger=logger.log)
                    switched = _wait_switched_to_new_record(driver, modal, old_anchor, old_key, timeout=8)
                except Exception:
                    pass

            if not switched:
                logger.log("⛔ Không chuyển được sang hồ sơ mới (DOM/KEY không đổi). Kết thúc.")
                break

            # Cập nhật thông tin cho vòng lặp kế tiếp
            try:
                module_thicong = modal.find_element(By.CSS_SELECTOR, "#vModuleThiCong[vmodule-name='xulydondangky']")
                tree_thicong = wait_jstree_ready_in(module_thicong, timeout=20)
                current_ma_don = extract_ma_don_from_tree(tree_thicong)  # có thể rỗng, nhưng KEY đã khác là đủ
            except Exception:
                current_ma_don = ""

            index += 1

        logger.log("✅ Hoàn tất toàn bộ hồ sơ trong phiên.")

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

    # --- Các trường nhập ---
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

    ttk.Label(frm, text="Trang bắt đầu:").grid(row=3, column=0, sticky="e", padx=4, pady=4)
    ent_start = ttk.Entry(frm, width=12)
    ent_start.grid(row=3, column=1, sticky="w", padx=4, pady=4)
    ent_start.insert(0, "1")

    # --- Thêm Combobox chọn tỉnh ---
    ttk.Label(frm, text="Chọn tỉnh:").grid(row=4, column=0, sticky="e", padx=4, pady=4)
    province_cb = ttk.Combobox(frm, state="readonly", width=33)
    province_cb["values"] = ["Đắk Lắk", "Phú Yên"]
    province_cb.current(0)  # mặc định là Đắk Lắk
    province_cb.grid(row=4, column=1, sticky="w", padx=4, pady=4)

    # --- Nút chạy ---
    btn_run = ttk.Button(frm, text="Chạy tự động")
    btn_run.grid(row=5, column=1, sticky="w", padx=4, pady=8)

    # --- Vùng log ---
    txt = tk.Text(root, height=18, state="disabled", bg="#0f1115", fg="#e5e7eb")
    txt.pack(fill="both", expand=True, padx=12, pady=(0, 12))
    logger = UILogger(txt)

    # --- Hàm click nút ---
    def on_run():
        username = ent_user.get().strip()
        password = ent_pass.get()
        code = ent_code.get().strip()
        province = province_cb.get()
        start_page = ent_start.get().strip()

        if not username or not password or not code:
            messagebox.showerror("Thiếu thông tin", "Vui lòng nhập đủ Username, Password và CODE.")
            return

        try:
            start_page = int(start_page)
        except:
            messagebox.showerror("Lỗi", "Trang bắt đầu phải là số nguyên ≥ 1.")
            return

        # Chọn URL theo tỉnh
        if province == "Phú Yên":
            base_url = "https://phy.mplis.gov.vn/dc/TichHopDongBoDuLieu/QuanLyDotBanGiao"
        else:
            base_url = "https://dla.mplis.gov.vn/dc/TichHopDongBoDuLieu/QuanLyDotBanGiao"

        # Chạy bot trong luồng riêng
        btn_run.config(state="disabled")
        logger.log(f"=== BẮT ĐẦU CHẠY ({province}) ===")

        def runner():
            run_bot(username, password, code, start_page, logger, base_url)
            btn_run.after(0, lambda: btn_run.config(state="normal"))

        threading.Thread(target=runner, daemon=True).start()

    btn_run.configure(command=on_run)
    root.mainloop()

if __name__ == "__main__":
    main()