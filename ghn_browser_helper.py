import os
import sys
import asyncio
import ctypes
import subprocess
from playwright.async_api import Browser, BrowserContext, Page, async_playwright

try:
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8')
except Exception:
    pass

SESSION_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ghn_session.json")
USER_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "playwright_user_data")
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120"
LOCK_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ghn_browser.lock")

_asyncio_lock = asyncio.Lock()
_persistent_playwright = None
_persistent_context = None

def cleanup_orphaned_chrome():
    if os.name == 'nt':
        try:
            out = subprocess.check_output('wmic process where "name=\'chrome.exe\'" get commandline,processid', shell=True, text=True, errors='ignore')
            for line in out.splitlines():
                if 'playwright_user_data' in line:
                    parts = line.strip().split()
                    if parts:
                        pid = parts[-1]
                        if pid.isdigit():
                            print(f"[System] Dọn dẹp tiến trình Chrome cũ PID {pid}...")
                            subprocess.run(f"taskkill /F /PID {pid}", shell=True, capture_output=True)
        except Exception as e:
            print(f"[!] Cảnh báo dọn dẹp Chrome cũ: {e}")

def is_pid_running(pid):
    if not pid:
        return False
    if os.name == 'nt':
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            SYNCHRONIZE = 0x00100000
            process = kernel32.OpenProcess(SYNCHRONIZE, False, int(pid))
            if process:
                kernel32.CloseHandle(process)
                return True
            return False
        except Exception:
            return False
    else:
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False

async def acquire_lock():
    await _asyncio_lock.acquire()
    while True:
        try:
            if os.path.exists(LOCK_FILE):
                try:
                    with open(LOCK_FILE, "r") as f:
                        pid = int(f.read().strip())
                except Exception:
                    pid = None
                
                if pid is not None and pid != os.getpid() and not is_pid_running(pid):
                    print(f"[System] Phát hiện lock file cũ từ PID {pid} đã chết. Tiến hành xóa.")
                    try:
                        os.remove(LOCK_FILE)
                    except Exception:
                        pass
                elif pid == os.getpid():
                    print(f"[System] Lock file thuộc về process hiện tại. Tiến hành ghi đè.")
                    try:
                        os.remove(LOCK_FILE)
                    except Exception:
                        pass
                else:
                    await asyncio.sleep(0.5)
                    continue
            
            fd = os.open(LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(fd, "w") as f:
                f.write(str(os.getpid()))
            break
        except FileExistsError:
            await asyncio.sleep(0.5)

def release_lock():
    try:
        if os.path.exists(LOCK_FILE):
            with open(LOCK_FILE, "r") as f:
                pid = int(f.read().strip())
            if pid == os.getpid():
                os.remove(LOCK_FILE)
    except Exception:
        pass
    try:
        if _asyncio_lock.locked():
            _asyncio_lock.release()
    except Exception:
        pass

async def get_shared_context() -> BrowserContext:
    """Return a singleton persistent Playwright context.

    The first call creates the context by launching Chromium (or Chrome if available)
    with a user data directory. Subsequent calls return the already‑created context.
    The function is thread‑safe via an asyncio lock and a lock‑file to guard against
    multiple processes on the same host.
    """
    global _persistent_playwright, _persistent_context

    # If we already have a context, verify it is still alive and connected.
    if _persistent_context:
        try:
            if _persistent_context.browser and not _persistent_context.browser.is_connected():
                raise Exception("Browser disconnected")
            _ = _persistent_context.pages
            return _persistent_context
        except Exception as e:
            print(f"[System] Persistent context không còn hiệu lực ({e}), khởi tạo lại...")
            _persistent_context = None
            if _persistent_playwright:
                try:
                    await _persistent_playwright.stop()
                except Exception:
                    pass
                _persistent_playwright = None

    # Acquire the cross‑process lock to ensure only one creator.
    await acquire_lock()
    try:
        # Double‑check after acquiring the lock.
        if _persistent_context:
            try:
                if _persistent_context.browser and not _persistent_context.browser.is_connected():
                    raise Exception("Browser disconnected")
                _ = _persistent_context.pages
                return _persistent_context
            except Exception:
                _persistent_context = None

        _persistent_playwright = await async_playwright().start()
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        launch_kwargs = {
            "user_data_dir": USER_DATA_DIR,
            "headless": True,
            "viewport": {"width": 1440, "height": 900},
            "user_agent": USER_AGENT,
            "permissions": ["clipboard-read", "clipboard-write"],
        }
        if os.path.exists(chrome_path):
            launch_kwargs["executable_path"] = chrome_path
            print("[System] Sử dụng Google Chrome chính thức để giảm OTP.")
        else:
            print("[System] Chrome không có, dùng Playwright Chromium.")

        try:
            _persistent_context = await _persistent_playwright.chromium.launch_persistent_context(**launch_kwargs)
        except Exception as e:
            if "ProcessSingleton" in str(e) or "already in use" in str(e):
                print(f"[System] Phát hiện xung đột Profile Directory ({e}). Đang tự động giải phóng Chrome...")
                cleanup_orphaned_chrome()
                await asyncio.sleep(1)
                _persistent_context = await _persistent_playwright.chromium.launch_persistent_context(**launch_kwargs)
            else:
                raise e
        return _persistent_context
    except Exception as e:
        print(f"[!] Lỗi khi khởi tạo persistent context: {e}")
        raise e
    finally:
        # Always release the lock – context may be None on error.
        release_lock()

# Backward compatible alias used by existing code.
async def get_ghn_context(browser: Browser = None) -> BrowserContext:
    """Compatibility wrapper.

    Existing callers pass a temporary ``browser`` instance which we no longer need.
    We close that temporary browser (if possible) and then return the shared context.
    """
    if browser is not None:
        try:
            await browser.close()
        except Exception:
            pass
    return await get_shared_context()

async def save_session(context: BrowserContext):
    """
    Lưu trữ trạng thái hiện tại của context vào file session (bao gồm cookies và localStorage)
    """
    try:
        await context.storage_state(path=SESSION_FILE)
        print(f"[🎉] Đã cập nhật và lưu session state thành công vào: {SESSION_FILE}")
    except Exception as e:
        print(f"[!] Không thể lưu session state: {e}")

async def check_and_login(page: Page, context: BrowserContext, target_url: str, chat_id=None) -> bool:
    """
    Kiểm tra trạng thái đăng nhập bằng cách truy cập target_url.
    Nếu bị redirect sang trang login/sso, sẽ tự động thực hiện login_flow và lưu session.
    """
    logged_in = False
    
    try:
        print("[System] Đang kiểm tra session...")
        await page.goto(target_url, wait_until="domcontentloaded", timeout=25000)
        await page.wait_for_timeout(4000)
        current_url = page.url
        print(f"[System] URL sau khi kiểm tra: {current_url}")
        if current_url != "about:blank" and "login" not in current_url.lower() and "sso" not in current_url.lower():
            login_form = await page.query_selector("input[type='password'], button:has-text('Đăng nhập')")
            if not login_form:
                print("[🎉] Session còn hiệu lực!")
                logged_in = True
            else:
                print("[System] Trang chứa form Đăng nhập, session đã hết hạn.")
    except Exception as e:
        print(f"[!] Cảnh báo khi kiểm tra session: {e}")
        try:
            current_url = page.url
            print(f"[System] URL hiện tại sau khi bị timeout: {current_url}")
            if current_url != "about:blank" and "login" not in current_url.lower() and "sso" not in current_url.lower():
                print("[🎉] Session vẫn còn hiệu lực (dựa trên URL hiện tại)!")
                logged_in = True
            else:
                print("[System] URL không hợp lệ (about:blank), tiến hành mở lại trang mục tiêu...")
                await page.goto(target_url, wait_until="domcontentloaded", timeout=30000)
                await page.wait_for_timeout(3000)
                current_url = page.url
                if current_url != "about:blank" and "login" not in current_url.lower() and "sso" not in current_url.lower():
                    logged_in = True
        except Exception as e2:
            print(f"[!] Thử mở lại trang thất bại: {e2}")
            
    if not logged_in:
        from ghn_automation import login_flow
        success = await login_flow(page, context, chat_id=chat_id)
        if not success:
            return False
            
        print("[System] Điều hướng tới trang mục tiêu...")
        await page.goto(target_url, wait_until="domcontentloaded", timeout=30000)
        await page.wait_for_timeout(5000)
        
    return True

