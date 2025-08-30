# -*- coding: utf-8 -*-
import os, sys, json, tempfile, subprocess, pathlib, time, threading, io, traceback
from datetime import datetime
from typing import Optional

import psutil
import pyperclip
from pynput import keyboard
import win32gui, win32process, win32com.client
from win32com.client import gencache
import pythoncom  # 线程内初始化 COM

# 托盘
import pystray
from PIL import Image, ImageDraw

# 可选通知
try:
    from win10toast import ToastNotifier
    TOASTER = ToastNotifier()
except Exception:
    TOASTER = None

APP_NAME = "MD→DOCX HotPaste"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
LOG_PATH = os.path.join(BASE_DIR, "md2docx.log")

DEFAULT_CONFIG = {
    "hotkey": "<ctrl>+b",
    "pandoc_path": "pandoc",
    "reference_docx": None,  # 可选：Pandoc 参考模板；不需要就设为 None
    "save_dir": r"%USERPROFILE%\Documents\md2docx_paste",
    "keep_file": False,
    "insert_target": "auto",  # auto|word|wps|none
    "notify": True
}

_state = {
    "enabled": True,
    "listener": None,
    "icon": None,
    "hotkey_str": DEFAULT_CONFIG["hotkey"],
    "config": DEFAULT_CONFIG.copy(),
    "last_ok": True,
    # 新增：异步触发防抖 + 互斥
    "running": False,
    "last_fire": 0.0,
}

# 触发防抖时间（秒）
FIRE_DEBOUNCE_SEC = 1

# ------------- 工具 -------------

def log(msg: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}\n"
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line)
    except Exception:
        pass

def load_config():
    cfg = DEFAULT_CONFIG.copy()
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                user = json.load(f)
            for k, v in user.items():
                cfg[k] = v
        except Exception as e:
            log(f"Load config error: {e}")
    cfg["save_dir"] = os.path.expandvars(cfg["save_dir"])
    _state["hotkey_str"] = cfg.get("hotkey", DEFAULT_CONFIG["hotkey"])
    _state["config"] = cfg
    return cfg

def save_config():
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(_state["config"], f, ensure_ascii=False, indent=2)
    except Exception as e:
        log(f"Save config error: {e}")

def notify(title, msg, ok=True):
    if not _state["config"].get("notify", True):
        return
    try:
        if _state["icon"]:
            _state["icon"].icon = make_icon(ok=ok, flash=True)
            _state["icon"].visible = True
            def restore():
                time.sleep(0.8)
                _state["icon"].icon = make_icon(ok=ok)
            threading.Thread(target=restore, daemon=True).start()
    except Exception:
        pass
    if TOASTER:
        try:
            TOASTER.show_toast(title, msg, duration=3, threaded=True, icon_path=None)
        except Exception:
            pass

def get_foreground_process_name() -> str:
    hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        return ""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        p = psutil.Process(pid)
        return os.path.basename(p.exe()).lower()
    except Exception:
        return ""

def active_target() -> str:
    name = get_foreground_process_name()
    if "winword" in name:
        return "word"
    if "wps" in name:
        return "wps"
    return ""

def ensure_dir(path: str):
    pathlib.Path(path).mkdir(parents=True, exist_ok=True)

def generate_output_path(keep_file: bool, save_dir: str) -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = f"md_paste_{ts}.docx"
    if keep_file:
        ensure_dir(save_dir)
        return os.path.join(save_dir, fname)
    else:
        return os.path.join(tempfile.gettempdir(), fname)

def run_pandoc(md_text: str, out_docx: str, pandoc: str, reference: Optional[str]):
    with tempfile.TemporaryDirectory(prefix="md2docx_") as td:
        in_md = os.path.join(td, "in.md")
        with open(in_md, "w", encoding="utf-8", newline="\n") as f:
            f.write(md_text)
        cmd = [
            pandoc, in_md,
            "--from", "markdown+tex_math_dollars+raw_tex",
            "--to", "docx",
            "-o", out_docx,
            "--highlight-style", "tango"
        ]
        if reference:
            cmd += ["--reference-doc", reference]
        
        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW  # 隐藏窗口
            creationflags = subprocess.CREATE_NO_WINDOW              # 不创建控制台

        res = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            shell=False,
            startupinfo=startupinfo,
            creationflags=creationflags
        )
        if res.returncode != 0:
            raise RuntimeError(res.stderr.strip() or res.stdout or "Pandoc failed")

# ------------- COM/插入 -------------

def ensure_com(func):
    def wrapper(*args, **kwargs):
        pythoncom.CoInitialize()
        try:
            return func(*args, **kwargs)
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    return wrapper

@ensure_com
def insert_into_word(docx_path: str) -> bool:
    """使用 InsertFile，将 DOCX 插入到 Word 当前光标位置。"""
    try:
        try:
            app = win32com.client.GetActiveObject("Word.Application")
        except Exception:
            app = gencache.EnsureDispatch("Word.Application")
        if getattr(app, "Documents", None) is None or app.Documents.Count == 0:
            return False
        sel = getattr(app, "Selection", None)
        if sel is None:
            return False
        sel.InsertFile(docx_path)
        return True
    except Exception as e:
        log(f"Insert Word failed: {e}")
        return False

@ensure_com
def insert_into_wps(docx_path: str) -> bool:
    """使用 InsertFile，将 DOCX 插入到 WPS 当前光标位置。"""
    for progid in ("kwps.Application", "wps.Application"):
        try:
            try:
                app = win32com.client.GetActiveObject(progid)
            except Exception:
                app = win32com.client.Dispatch(progid)
        except Exception as e:
            log(f"Try {progid} failed: {e}")
            continue
        try:
            sel = getattr(getattr(app, "ActiveWindow", app), "Selection", None)
            if sel is None:
                log(f"{progid} has no Selection")
                continue
            sel.InsertFile(docx_path)
            return True
        except Exception as e:
            log(f"Try {progid} InsertFile failed: {e}")
            continue
    return False

# ------------- 异步触发（关键修复） -------------

def _hotkey_trigger_async():
    """
    仅由键盘钩子线程调用：快速返回，把重活丢到后台线程；做防抖与互斥。
    """
    now = time.time()
    # 防抖：短时间重复按键直接忽略
    if now - _state.get("last_fire", 0.0) < FIRE_DEBOUNCE_SEC:
        return
    _state["last_fire"] = now

    # 互斥：已有任务在跑就不再启动
    if _state.get("running", False):
        return

    def worker():
        _state["running"] = True
        try:
            do_convert_and_insert()  # 真正的工作
        except Exception:
            log(traceback.format_exc())
        finally:
            _state["running"] = False

    threading.Thread(target=worker, daemon=True).start()

# ------------- 热键 -------------

def _start_hotkey_listener():
    hotkey_def = _state["hotkey_str"]

    def _on_hotkey():
        if _state["enabled"]:
            try:
                _hotkey_trigger_async()  # 立即返回，不在钩子线程里做重活
            except Exception:
                log(traceback.format_exc())

    mapping = { hotkey_def: _on_hotkey }
    listener = keyboard.GlobalHotKeys(mapping)
    listener.daemon = True
    listener.start()
    _state["listener"] = listener
    log(f"Hotkey started: {hotkey_def}")

def _stop_hotkey_listener():
    if _state["listener"]:
        try:
            _state["listener"].stop()
        except Exception:
            pass
    _state["listener"] = None
    log("Hotkey stopped")

def restart_hotkey():
    _stop_hotkey_listener()
    _start_hotkey_listener()

# ------------- 托盘图标 -------------

def make_icon(ok=True, flash=False):
    size = (64, 64)
    img = Image.new("RGBA", size, (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    bg = (30, 30, 30, 255)
    d.rectangle([0, 0, size[0], size[1]], fill=bg)
    color = (60, 200, 80, 255) if ok else (220, 70, 70, 255)
    if flash:
        color = tuple(min(255, int(c * 1.3)) if i < 3 else c for i, c in enumerate(color))
    d.ellipse([10, 10, 54, 54], fill=color)
    return img

# ------------- 托盘菜单动作 -------------

def on_toggle_enabled(icon, item):
    _state["enabled"] = not _state["enabled"]
    icon.icon = make_icon(ok=_state["enabled"])
    notify(APP_NAME, "已启用热键" if _state["enabled"] else "已暂停热键", ok=_state["enabled"])

def on_target_auto(icon, item):
    _state["config"]["insert_target"] = "auto"
    save_config()
    notify(APP_NAME, "插入目标：Auto", ok=True)

def on_target_word(icon, item):
    _state["config"]["insert_target"] = "word"
    save_config()
    notify(APP_NAME, "插入目标：Word", ok=True)

def on_target_wps(icon, item):
    _state["config"]["insert_target"] = "wps"
    save_config()
    notify(APP_NAME, "插入目标：WPS", ok=True)

def on_target_none(icon, item):
    _state["config"]["insert_target"] = "none"
    save_config()
    notify(APP_NAME, "仅生成，不插入", ok=True)

def on_toggle_keep(icon, item):
    v = not _state["config"].get("keep_file", False)
    _state["config"]["keep_file"] = v
    save_config()
    notify(APP_NAME, "保留文件：开启" if v else "保留文件：关闭", ok=True)

def on_open_save_dir(icon, item):
    path = _state["config"].get("save_dir", DEFAULT_CONFIG["save_dir"])
    path = os.path.expandvars(path)
    ensure_dir(path)
    os.startfile(path)

def on_open_log(icon, item):
    if not os.path.exists(LOG_PATH):
        open(LOG_PATH, "w", encoding="utf-8").close()
    os.startfile(LOG_PATH)

def on_edit_config(icon, item):
    if not os.path.exists(CONFIG_PATH):
        save_config()
    os.startfile(CONFIG_PATH)

def on_reload(icon, item):
    load_config()
    restart_hotkey()
    notify(APP_NAME, "配置已重载", ok=True)

def on_quit(icon, item):
    _stop_hotkey_listener()
    icon.stop()

def build_menu():
    cfg = _state["config"]
    return pystray.Menu(
        pystray.MenuItem("启用热键", on_toggle_enabled, checked=lambda item: _state["enabled"]),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            "插入目标",
            pystray.Menu(
                pystray.MenuItem("Auto", on_target_auto, checked=lambda i: cfg.get("insert_target")=="auto"),
                pystray.MenuItem("Word", on_target_word, checked=lambda i: cfg.get("insert_target")=="word"),
                pystray.MenuItem("WPS", on_target_wps, checked=lambda i: cfg.get("insert_target")=="wps"),
                pystray.MenuItem("None (仅生成)", on_target_none, checked=lambda i: cfg.get("insert_target")=="none"),
            )
        ),
        pystray.MenuItem("保留生成文件", on_toggle_keep, checked=lambda item: cfg.get("keep_file", False)),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("打开保存目录", on_open_save_dir),
        pystray.MenuItem("查看日志", on_open_log),
        pystray.MenuItem("编辑配置", on_edit_config),
        pystray.MenuItem("重载配置/热键", on_reload),
        pystray.MenuItem("退出", on_quit)
    )

# ------------- 核心流程 -------------

@ensure_com
def do_convert_and_insert():
    cfg = _state["config"]
    try:
        md = pyperclip.paste()
        if not md or not md.strip():
            notify(APP_NAME, "剪贴板为空，未处理。", ok=False)
            return

        out_docx = generate_output_path(cfg.get("keep_file", False),
                                        cfg.get("save_dir", DEFAULT_CONFIG["save_dir"]))
        run_pandoc(
            md_text=md,
            out_docx=out_docx,
            pandoc=cfg.get("pandoc_path", "pandoc"),
            reference=cfg.get("reference_docx")
        )

        target = cfg.get("insert_target", "auto")
        if target == "auto":
            target = active_target() or "word"

        inserted = False
        if target == "word":
            inserted = insert_into_word(out_docx)
        elif target == "wps":
            inserted = insert_into_wps(out_docx)
        elif target == "none":
            inserted = False

        # 清理
        if not cfg.get("keep_file", False):
            try:
                if inserted or target != "none":
                    time.sleep(0.5)
                if os.path.exists(out_docx):
                    os.remove(out_docx)
            except Exception as e:
                log(f"Cleanup temp failed: {e}")

        if target == "none":
            notify(APP_NAME, "已生成 DOCX（仅生成，不插入）。", ok=True)
        elif inserted:
            notify(APP_NAME, f"已插入到 {target.upper()}。", ok=True)
        else:
            notify(APP_NAME, f"未能插入到 {target.upper()}，请确认软件已打开且有光标。", ok=False)

    except Exception:
        buf = io.StringIO()
        traceback.print_exc(file=buf)
        log(buf.getvalue())
        notify(APP_NAME, "转换失败，请查看日志。", ok=False)

# ------------- 主程序 -------------

def main():
    load_config()
    restart_hotkey()
    icon = pystray.Icon(APP_NAME, make_icon(ok=True), APP_NAME, build_menu())
    _state["icon"] = icon
    icon.run()

if __name__ == "__main__":
    main()
