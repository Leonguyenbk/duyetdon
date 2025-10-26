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

def process_current_record(driver, wait, logger, modal):
    try:
        module_thicong  = modal.find_element(By.CSS_SELECTOR, "#vModuleThiCong[vmodule-name='xulydondangky']")
        module_vanhanh  = modal.find_element(By.CSS_SELECTOR, "#vModuleVanHanh[vmodule-name='xulydondangky']")
        tree_thicong = wait_jstree_ready_in(module_thicong, timeout=30)
        tree_vanhanh = wait_jstree_ready_in(module_vanhanh, timeout=30)

        logger.log("   - B1: Thêm vào dữ liệu vận hành…")
        anchor_tc = find_tt_dangky_anchor(tree_thicong)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor_tc)
        ok_tc = context_click_when_enabled(
            driver, anchor_tc, rel=2, label="Thêm vào dữ liệu vận hành",
            logger=logger.log, modal=modal
        )
        if ok_tc:
            quick_confirm_if_present(driver, root_el=modal, soft_timeout=1.8)
            logger.log("     ✓ Hoàn tất B1.")
        else:
            logger.log("     ❌ Lỗi B1.")

        time.sleep(1)

        logger.log("   - B2: Cập nhật lịch sử tất cả…")
        anchor_vh = find_tt_dangky_anchor(tree_vanhanh)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor_vh)
        ok_vh = context_click_when_enabled(
            driver, anchor_vh, rel=4, label="Cập nhật lịch sử tất cả",
            logger=logger.log, modal=modal
        )
        if ok_vh:
            quick_confirm_if_present(driver, root_el=modal, soft_timeout=1.8)
            logger.log("     ✓ Hoàn tất B2.")
        else:
            logger.log("     ❌ Lỗi B2.")

        time.sleep(1)
    except Exception as e:
        logger.log(f"   🔥 Lỗi process_current_record: {e.__class__.__name__}")
        logger.log(traceback.format_exc())

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
                        return origFetch(...args)
                            .finally(() => window.pendingFetchCount--;
                        );
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

def open_filter_and_select_chua_xu_ly(driver, wait, table_id="formSearchDSDDK", logger=None):
    log = (logger.log if logger else print)
    log("🔍 Mở bộ lọc tìm kiếm mở rộng…")
    btn_filter = wait.until(EC.element_to_be_clickable((By.ID, "btnShowFormSearchDSDDK")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_filter)
    try:
        btn_filter.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn_filter)
    time.sleep(1.2)
    log("🧩 Chọn trạng thái 'Chưa xử lý'…")
    sel = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "select[name='trangThai']")))
    try:
        driver.execute_script("""
            const sel = document.querySelector("select[name='trangThai']");
            if (sel) {
                sel.value = '0';
                sel.dispatchEvent(new Event('change', {bubbles:true}));
                if (window.jQuery) jQuery(sel).trigger('change');
            }
        """)
    except Exception:
        pass
    try:
        select_element = driver.find_element(By.CSS_SELECTOR, '#formSearchDSDDK > div > div > fieldset > div > div:nth-child(2) > div > div:nth-child(3) > div > select')
        dropdown = Select(select_element)
        dropdown.select_by_value("0")
    except Exception:
        pass
    log("🔎 Bấm nút 'Tìm kiếm'…")
    btn_search = wait.until(EC.element_to_be_clickable((By.ID, "btnTimKiemMoRong")))
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn_search)
    try:
        btn_search.click()
    except Exception:
        driver.execute_script("arguments[0].click();", btn_search)
    try:
        old_row = driver.find_element(By.CSS_SELECTOR, f"#{table_id} tbody tr")
        WebDriverWait(driver, 10).until(EC.staleness_of(old_row))
    except Exception:
        pass
    try:
        wait_for_table_loaded(driver, table_id=table_id, timeout=20)
    except Exception:
        pass
    log("✅ Đã lọc và tải danh sách 'Chưa xử lý'.")

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
        m = re.search(r"tổng\s*số\s+([\d\.,\s]+)", text, flags=re.IGNORECASE)
        if m:
            return _normalize_int(m.group(1))
        nums = re.findall(r"[\d\.,]+", text)
        return _normalize_int(nums[-1]) if nums else 0
    except Exception:
        return 0

# ============== PHỤC HỒI CHỈ KHI NEXT LỖI: ĐÓNG MODAL ⇒ (CÓ THÌ) TÌM KIẾM ⇒ ĐỌC TỔNG ⇒ XỬ LÝ ĐƠN ==============
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

def resubmit_current_search_if_present(driver, wait, logger=None):
    """
    KHÔNG mở 'Tìm kiếm mở rộng'.
    Nếu nút 'Tìm kiếm' hiện có (btnTimKiemMoRong) xuất hiện thì bấm để refresh; nếu không có thì bỏ qua.
    """
    try:
        btn = wait.until(EC.presence_of_element_located((By.ID, "btnTimKiemMoRong")))
        vis_ok = driver.execute_script("""
            const b = arguments[0]; const r = b.getBoundingClientRect();
            const s = getComputedStyle(b);
            return r.width>0 && r.height>0 && s.display!=='none' && s.visibility!=='hidden';
        """, btn)
        if vis_ok:
            if logger: logger.log("🔁 Bấm lại 'Tìm kiếm' (không mở Tìm kiếm mở rộng).")
            try: btn.click()
            except Exception: driver.execute_script("arguments[0].click();", btn)
            try:
                wait_for_table_loaded(driver, table_id="tblTTThuaDat", timeout=15)
            except Exception:
                pass
            return True
    except Exception:
        pass
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
      3) Đọc lại tổng số hồ sơ 'Chưa xử lý'
      4) Nếu còn hồ sơ → bấm 'Xử lý đơn' để mở modal mới
    Trả về (modal_moi hoặc None, remaining_after).
    """
    ok = close_xuly_modal_if_open(driver, soft_timeout=6)
    logger and logger.log("   ↪️ Đóng modal " + ("✓" if ok else "✗ (không mở/đóng không được)"))

    _ = resubmit_current_search_if_present(driver, wait, logger)

    remaining_after = get_total_records_from_info(driver, info_id="tblDanhSachGoiTinDongBo_info")
    logger and logger.log(f"   📊 Sau refresh: còn {remaining_after if remaining_after else 0} hồ sơ 'Chưa xử lý'.")

    if remaining_after == 0:
        logger and logger.log("✅ Không còn hồ sơ nào để xử lý. Dừng.")
        return None, 0

    try:
        modal = reopen_modal_from_list(driver, wait, logger, timeout=25)
        logger and logger.log("   🔄 Đã mở lại modal sau refresh.")
        return modal, remaining_after
    except Exception as e:
        logger and logger.log(f"   ⛔ Không thể mở lại modal: {e.__class__.__name__}")
        return None, remaining_after

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

        # Mở 'Chức năng' → 'Xem danh sách'
        logger.log("📂 Mở 'Chức năng' → 'Xem danh sách'…")
        wait.until(EC.element_to_be_clickable((By.ID, "drop1"))).click()
        wait.until(EC.element_to_be_clickable((By.ID, "btnXemDanhSach"))).click()

        # Lọc 'Chưa xử lý' lần đầu
        logger.log("🧰 Lọc & áp dụng: Chưa xử lý…")
        open_filter_and_select_chua_xu_ly(driver, wait, table_id="formSearchDSDDK", logger=logger)
        time.sleep(1.0)

        total_records = get_total_records_from_info(driver, info_id="tblDanhSachGoiTinDongBo_info")
        if total_records > 0:
            logger.log(f"📊 Tổng 'Chưa xử lý' ban đầu: {total_records}.")
        else:
            logger.log("⚠️ Không đọc được tổng hồ sơ ban đầu. Bot sẽ chạy cho đến khi hết (theo refresh).")

        # Mở modal 'Xử lý đơn'
        btn_xu_ly = wait.until(EC.presence_of_element_located((By.ID, "btnXuLyDon")))
        driver.execute_script("arguments[0].scrollIntoView(true);", btn_xu_ly)
        driver.execute_script("arguments[0].click();", btn_xu_ly)

        logger.log("⏳ Đợi modal 'Xử lý đơn đăng ký' hiển thị…")
        modal = wait_xuly_modal(driver, timeout=25)

        logger.log("🔁 Bắt đầu duyệt tuần tự từng hồ sơ (StepForward)…")
        processed = 0

        while True:
            # Log mã đơn hiện tại (nếu lấy được)
            try:
                module_thicong = modal.find_element(By.CSS_SELECTOR, "#vModuleThiCong[vmodule-name='xulydondangky']")
                tree_thicong = wait_jstree_ready_in(module_thicong, timeout=20)
                current_ma_don = extract_ma_don_from_tree(tree_thicong)
            except Exception:
                current_ma_don = ""
            logger.log(f"— Đang xử lý hồ sơ: {current_ma_don or '(không rõ)'}")

            try:
                old_anchor, old_key = _get_tt_anchor_and_key(modal, timeout=8)
            except Exception:
                old_anchor, old_key = None, ""

            # Xử lý hồ sơ hiện tại
            try:
                process_current_record(driver, wait, logger, modal)
            except Exception as e:
                logger.log(f"   ⚠️ Lỗi process_current_record: {e.__class__.__name__}: {e}. Vẫn tiếp tục…")
            time.sleep(0.5)

            processed += 1

            # Thử Next (tối đa 5 lần)
            next_clicked = False
            for attempt in range(1, 6):
                if click_step_forward(modal):
                    next_clicked = True
                    break
                logger.log(f"   (Click Next lần {attempt} chưa được, chờ 1s rồi thử lại…)"); time.sleep(1)

            # Nếu không Next được → refresh + đo tổng → dừng nếu hết
            if not next_clicked:
                logger.log("   ⛔ Không Next được. Đóng modal ⇒ refresh danh sách ⇒ đọc tổng ⇒ 'Xử lý đơn'.")
                modal2, remaining_after = recover_close_resubmit_reopen(driver, wait, logger)
                if remaining_after == 0:
                    logger.log(f"✅ Đã xử lý {processed} hồ sơ. Không còn hồ sơ 'Chưa xử lý'. Kết thúc.")
                    break
                if modal2 is None:
                    logger.log(f"   ⛔ Phục hồi thất bại. Đã xử lý {processed} hồ sơ. Dừng.")
                    break
                modal = modal2
                continue

            # Đợi thực sự chuyển sang hồ sơ khác
            try:
                wait_ajax_idle(driver, max_wait=0.8, log=logger.log)
            except Exception:
                pass

            switched = _wait_switched_to_new_record(driver, modal, old_anchor, old_key, timeout=10)
            if not switched:
                # Thử nudge nhanh
                try:
                    nudge_by_next_back(driver, modal, logger=logger.log)
                    switched = _wait_switched_to_new_record(driver, modal, old_anchor, old_key, timeout=6)
                except Exception:
                    pass

            # Nếu vẫn không chuyển → refresh + đo tổng → dừng nếu hết
            if not switched:
                logger.log("   ⛔ Không chuyển được hồ sơ. Đóng modal ⇒ refresh danh sách ⇒ đọc tổng ⇒ 'Xử lý đơn'.")
                modal2, remaining_after = recover_close_resubmit_reopen(driver, wait, logger)
                if remaining_after == 0:
                    logger.log(f"✅ Đã xử lý {processed} hồ sơ. Không còn hồ sơ 'Chưa xử lý'. Kết thúc.")
                    break
                if modal2 is None:
                    logger.log(f"   ⛔ Phục hồi thất bại. Đã xử lý {processed} hồ sơ. Dừng.")
                    break
                modal = modal2
                continue

        logger.log("🏁 Hoàn tất phiên xử lý.")

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

    ttk.Label(frm, text="Trang bắt đầu:").grid(row=3, column=0, sticky="e", padx=4, pady=4)
    ent_start = ttk.Entry(frm, width=12)
    ent_start.grid(row=3, column=1, sticky="w", padx=4, pady=4)
    ent_start.insert(0, "1")

    ttk.Label(frm, text="Chọn tỉnh:").grid(row=4, column=0, sticky="e", padx=4, pady=4)
    province_cb = ttk.Combobox(frm, state="readonly", width=33)
    province_cb["values"] = ["Đắk Lắk", "Phú Yên"]
    province_cb.current(0)
    province_cb.grid(row=4, column=1, sticky="w", padx=4, pady=4)

    btn_run = ttk.Button(frm, text="Chạy tự động")
    btn_run.grid(row=5, column=1, sticky="w", padx=4, pady=8)

    txt = tk.Text(root, height=18, state="disabled", bg="#0f1115", fg="#e5e7eb")
    txt.pack(fill="both", expand=True, padx=12, pady=(0, 12))
    logger = UILogger(txt)

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

        if province == "Phú Yên":
            base_url = "https://phy.mplis.gov.vn/dc/TichHopDongBoDuLieu/QuanLyDotBanGiao"
        else:
            base_url = "https://dla.mplis.gov.vn/dc/TichHopDongBoDuLieu/QuanLyDotBanGiao"

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
