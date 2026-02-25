# -*- coding: utf-8 -*-
# VS after PdbGen (restart + socket): одна кнопка "VS" в конце тулбара 'pdbgen'.
# По клику: мягко закрывает ТОЛЬКО VS с указанным .sln и запускает новую VS с ним.
# Даблклик в IDA отправляет текущий адрес в VS и фокусирует окно.
#
# ВНИМАНИЕ:
#  - Автоконнект к VS при старте IDA выключен (см. AUTO_CONNECT_ON_START = False).
#  - Экшен "Link" (Ctrl+Alt+L) включает/выключает фоновый приём адресов от VS (порт 8888).
#  - Совместимо с IDA 7.x–9.x, Python 3.x. Windows-only. Без PyQt/PySide.

import os
import time
import json
import threading
import socket
import ctypes
from ctypes import wintypes
import subprocess

import idaapi
import ida_kernwin as kw
import idc

# === Настройки ===
ACTION_NAME     = "czar.vs.afterpdb.restart"
ACTION_LABEL    = "VS"
ACTION_TIP      = "Рестарт Visual Studio с re.sln (не трогая другие)"
HOTKEY          = "Ctrl+Alt+V"

PDBGEN_TOOLBAR  = "pdbgen"        # тулбар, который создаёт PdbGenerator
FALLBACK_TB     = "MainToolBar"   # если 'pdbgen' не появится
TIMER_STEP_MS   = 300
MAX_TRIES       = 20              # ~6 секунд ожидания

# Управление автоконнектом:
AUTO_CONNECT_ON_START = False      # НЕ подключаться к VS на старте IDA

# Ваш solution:
SOLUTION        = r"d:\VSProjects\re\re.sln"
SOL_BASENAME    = os.path.splitext(os.path.basename(SOLUTION))[0]  # 're'

# Путь к devenv. Можно оставить "devenv.exe" (будет резолвиться через PATH/where).
DEVENV          = "devenv.exe"
# DEVENV        = r"C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe"

# Закрытие найденной инстанции:
CLOSE_TIMEOUT_SEC    = 8.0         # ждём мягкого закрытия (WM_CLOSE)
FORCE_KILL_IF_HANGS  = False       # форс-килл по таймауту (по умолчанию False)

# TCP-взаимодействие с VS:
VS_HOST        = "localhost"
VS_PORT        = 8891
RECONNECT_DELAY_S = 0.8            # задержка перед попыткой переподключения
SOCK_RECV_BUFSZ = 4096
# Для фокуса окна VS через точное имя; если не найдём — fallback по эвристике.
VS_WIN_CLASS   = "wndclass_desked_gsk"
VS_WIN_TITLE   = "re (Debugging) - Microsoft Visual Studio"
# ==============

_state = {"tries": 0, "attached_to": None, "timer_live": False}
_sock_thread = None
_sock_stop = threading.Event()
_view_hook = None

def _msg(s): kw.msg("[VS-AfterPdbGen] %s\n" % s)

# --------- WinAPI (ctypes) ---------
user32   = ctypes.WinDLL("user32",   use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
shell32  = ctypes.WinDLL("shell32",  use_last_error=True)

EnumWindows  = user32.EnumWindows
GetWindowTextLengthW = user32.GetWindowTextLengthW
GetWindowTextW = user32.GetWindowTextW
IsWindowVisible = user32.IsWindowVisible
GetWindow = user32.GetWindow
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
PostMessageW = user32.PostMessageW
FindWindowW = user32.FindWindowW
SetForegroundWindow = user32.SetForegroundWindow

OpenProcess = kernel32.OpenProcess
CloseHandle = kernel32.CloseHandle
WaitForSingleObject = kernel32.WaitForSingleObject
QueryFullProcessImageNameW = kernel32.QueryFullProcessImageNameW
ShellExecuteW = shell32.ShellExecuteW

EnumWindows.argtypes = [ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM), wintypes.LPARAM]
GetWindowTextLengthW.argtypes = [wintypes.HWND]
GetWindowTextW.argtypes = [wintypes.HWND, ctypes.c_wchar_p, ctypes.c_int]
IsWindowVisible.argtypes = [wintypes.HWND]
GetWindow.argtypes = [wintypes.HWND, ctypes.c_uint]
GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
FindWindowW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR]
SetForegroundWindow.argtypes = [wintypes.HWND]
OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
CloseHandle.argtypes = [wintypes.HANDLE]
WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, ctypes.c_wchar_p, ctypes.POINTER(wintypes.DWORD)]
ShellExecuteW.argtypes = [wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_int]
ShellExecuteW.restype = wintypes.HINSTANCE

# consts
GW_OWNER  = 4
WM_CLOSE  = 0x0010
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
SYNCHRONIZE = 0x00100000
WAIT_OBJECT_0 = 0

# --------- Поиск целевой инстанции VS ---------

def _get_proc_exe(pid):
    h = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return None
    try:
        buf = ctypes.create_unicode_buffer(32768)
        n = wintypes.DWORD(len(buf))
        if QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(n)):
            return buf.value
        return None
    finally:
        CloseHandle(h)

def _wait_pid_exit(pid, timeout_sec):
    h = OpenProcess(SYNCHRONIZE, False, pid)
    if not h:
        return True
    try:
        r = WaitForSingleObject(h, int(timeout_sec * 1000))
        return r == WAIT_OBJECT_0
    finally:
        CloseHandle(h)

def _enum_top_windows():
    res = []
    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def _cb(hwnd, lparam):
        if not IsWindowVisible(hwnd): return True
        if GetWindow(hwnd, GW_OWNER): return True
        n = GetWindowTextLengthW(hwnd)
        if n <= 0: return True
        buf = ctypes.create_unicode_buffer(n + 1)
        GetWindowTextW(hwnd, buf, len(buf))
        title = buf.value
        pid = wintypes.DWORD(0)
        GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        res.append((hwnd, pid.value, title))
        return True
    EnumWindows(_cb, 0)
    return res

def _find_vs_pid_by_window(sol_basename):
    name = sol_basename.lower()
    for hwnd, pid, title in _enum_top_windows():
        t = (title or "").strip()
        low = t.lower()
        if "microsoft visual studio" not in low:
            continue
        exe = (_get_proc_exe(pid) or "").lower()
        if not exe.endswith("devenv.exe"):
            continue
        if " - microsoft visual studio" in low:
            prefix = low.split(" - microsoft visual studio", 1)[0]
            candidate = prefix.split(" - ")[-1].strip()
            cand = candidate.lower()
            if cand == name or cand.endswith("\\" + name) or cand.endswith("/" + name):
                return pid
    return None

def _ps_list_devenv_cmdlines():
    try:
        cmd = [
            "powershell", "-NoProfile", "-NonInteractive", "-Command",
            "$p=Get-CimInstance Win32_Process -Filter \"name='devenv.exe'\" | "
            "Select-Object ProcessId,CommandLine; $p | ConvertTo-Json -Compress"
        ]
        out = subprocess.check_output(cmd, stderr=subprocess.DEVNULL)
        txt = out.decode("utf-8", errors="ignore").strip()
        if not txt: return []
        data = json.loads(txt)
        if isinstance(data, dict): data = [data]
        res = []
        for it in data:
            pid = it.get("ProcessId")
            cl  = it.get("CommandLine") or ""
            if isinstance(pid, int):
                res.append({"Pid": pid, "Cmd": cl})
        return res
    except Exception:
        return []

def _find_vs_pid_by_cmdline(sln_path):
    want = os.path.normcase(os.path.abspath(sln_path))
    for it in _ps_list_devenv_cmdlines():
        cl = (it["Cmd"] or "")
        cl_low = os.path.normcase(cl.replace('"', '').replace("'", ""))
        if want in cl_low:
            return it["Pid"]
    return None

# ---------- Запуск VS (через ShellExecute, устойчиво) ----------

def _resolve_devenv():
    p = DEVENV
    if os.path.isabs(p) and os.path.isfile(p):
        return p
    # попробуем where
    try:
        out = subprocess.check_output(["cmd", "/c", "where", "devenv"], stderr=subprocess.DEVNULL)
        for ln in out.decode("utf-8", errors="ignore").splitlines():
            ln = ln.strip()
            if os.path.isfile(ln) and ln.lower().endswith("devenv.exe"):
                return ln
    except Exception:
        pass
    return p  # ShellExecute тоже поищет в PATH

def _shell_execute_open(file_, params=None, cwd=None, show=1):
    h = shell32.ShellExecuteW(None, "open", file_, params or None, cwd or None, int(show))
    if h <= 32:
        raise OSError(f"ShellExecute failed, code={h}")

def _open_solution_fresh():
    dev = _resolve_devenv()
    try:
        _shell_execute_open(dev, SOLUTION, None, 1)
        _msg(f"Запускаю: {dev} {SOLUTION}")
    except Exception as e:
        try:
            os.startfile(SOLUTION)
            _msg(f"Открываю по ассоциации: {SOLUTION}")
        except Exception:
            kw.warning(f"[VS-AfterPdbGen] Не удалось запустить Visual Studio:\n{e}\nDEVENV={dev}")

def _close_vs_instance(pid):
    # найдём главное окно процесса и отправим WM_CLOSE
    hwnd_main = None
    for hwnd, p, title in _enum_top_windows():
        if p == pid:
            hwnd_main = hwnd
            break
    if hwnd_main:
        try:
            PostMessageW(hwnd_main, WM_CLOSE, 0, 0)
            _msg(f"Отправлен WM_CLOSE PID={pid}, ждём до {CLOSE_TIMEOUT_SEC:.1f}с")
            if _wait_pid_exit(pid, CLOSE_TIMEOUT_SEC):
                _msg("Инстанция VS закрылась.")
                return True
        except Exception:
            pass
    _msg("Инстанция VS не закрылась (возможно, диалог сохранения).")
    if FORCE_KILL_IF_HANGS:
        try:
            subprocess.call(["taskkill", "/PID", str(pid), "/T", "/F"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            time.sleep(0.5)
            _msg("Сделан форс-килл.")
        except Exception:
            pass
        return True
    return False

def _restart_solution_instance():
    if not os.path.isfile(SOLUTION):
        kw.warning(f"Файл solution не найден:\n{SOLUTION}")
        return
    pid = _find_vs_pid_by_window(SOL_BASENAME) or _find_vs_pid_by_cmdline(SOLUTION)
    if pid:
        _msg(f"Найдена инстанция VS с '{SOL_BASENAME}': PID={pid}")
        closed = _close_vs_instance(pid)
        if not closed:
            _msg("Не закрыли (оставляем как есть). Открываю новую с .sln.")
        _open_solution_fresh()
    else:
        _msg("Инстанция с нужным solution не найдена — просто открываю .sln.")
        _open_solution_fresh()

# ---------- Фокус окна VS ----------

def _focus_vs_window_exact():
    """Пытаемся сфокусироваться по точному классу и заголовку."""
    try:
        hwnd = FindWindowW(VS_WIN_CLASS, VS_WIN_TITLE)
        if hwnd:
            SetForegroundWindow(hwnd)
            return True
    except Exception:
        pass
    return False

def _focus_vs_window_heuristic():
    """Если точный заголовок не подошёл — ищем VS с нужным sol по заголовку."""
    for hwnd, pid, title in _enum_top_windows():
        t = (title or "")
        low = t.lower()
        if "microsoft visual studio" not in low:
            continue
        if (" " + SOL_BASENAME.lower() + " ") in (" " + low + " "):  # грубая эвристика
            try:
                SetForegroundWindow(hwnd)
                return True
            except Exception:
                pass
    return False

def focus_vs_window():
    if not _focus_vs_window_exact():
        _focus_vs_window_heuristic()

# ---------- Socket: авто-реконнект и приём адресов ----------

def _recv_loop_forever():
    """Подключаемся к VS, ждём строки с hex-адресами и прыгаем на них в IDA. Автореконнект."""
    while not _sock_stop.is_set():
        try:
            with socket.create_connection((VS_HOST, VS_PORT), timeout=3.0) as s:
                s.settimeout(0.5)  # для регулярной проверки stop-события
                _msg(f"Подключено к VS на {VS_HOST}:{VS_PORT}")
                buf = b""
                while not _sock_stop.is_set():
                    try:
                        chunk = s.recv(SOCK_RECV_BUFSZ)
                        if not chunk:
                            break
                        buf += chunk
                        while b"\n" in buf:
                            line, buf = buf.split(b"\n", 1)
                            line = line.strip()
                            if not line:
                                continue
                            try:
                                ea = int(line.decode("utf-8", errors="ignore"), 16)
                            except ValueError:
                                _msg(f"[Socket] Неверная строка: {line!r}")
                                continue
                            def _jump():
                                idc.jumpto(ea)
                            kw.execute_sync(_jump, kw.MFF_FAST)
                    except socket.timeout:
                        continue
        except Exception:
            # можно раскомментить для подробностей:
            # _msg(f"[Socket] Подключение не удалось: {e}")
            pass
        if _sock_stop.is_set():
            break
        _msg("Соединение закрыто. Повторная попытка...")
        if _sock_stop.wait(RECONNECT_DELAY_S):
            break

def _is_socket_running():
    return (_sock_thread is not None) and _sock_thread.is_alive() and not _sock_stop.is_set()

def start_socket_thread():
    """Запустить фоновый поток приёма адресов (если не запущен)."""
    global _sock_thread
    if _is_socket_running():
        return
    _sock_stop.clear()
    t = threading.Thread(target=_recv_loop_forever, name="VSRecv", daemon=True)
    t.start()
    _sock_thread = t
    _msg("Приём адресов: включен (автоподключение к VS).")

def stop_socket_thread():
    """Остановить поток приёма адресов (если запущен)."""
    global _sock_thread
    if _sock_thread and _sock_thread.is_alive():
        _sock_stop.set()
        try:
            _sock_thread.join(timeout=1.0)
        except Exception:
            pass
    _sock_thread = None
    _msg("Приём адресов: выключен.")

# ---------- Hook на даблклик: отправка адреса в VS + фокус ----------

class _DoubleClickHook(kw.View_Hooks):
    def view_dblclick(self, view, event):
        self.send_to_vs()
        return False  # передаём обработку дальше

    def send_to_vs(self):
        try:
            ea = idc.get_screen_ea()
            if ea == idc.BADADDR:
                return
            txt = f"{ea:X}"  # без \n — как в вашем исходнике
            with socket.create_connection((VS_HOST, VS_PORT), timeout=0.5) as s:
                s.sendall(txt.encode("utf-8"))
        except Exception as e:
            _msg(f"[send] Ошибка отправки адреса: {e}")
        # Пытаемся сфокусировать VS
        focus_vs_window()

def install_view_hook():
    global _view_hook
    if _view_hook is None:
        _view_hook = _DoubleClickHook()
        if _view_hook.hook():
            _msg("[*] View Hook активирован (даблклик -> отправка в VS)")
        else:
            _msg("[ERROR] Не удалось активировать View Hook")

def uninstall_view_hook():
    global _view_hook
    if _view_hook is not None:
        try:
            _view_hook.unhook()
        except Exception:
            pass
        _view_hook = None

# ---------- IDA UI: кнопка после PdbGen ----------

def _attach_last(toolbar_name):
    # отцепить от всех и прицепить в конец указанного тулбара
    for tb in (PDBGEN_TOOLBAR, FALLBACK_TB):
        try: kw.detach_action_from_toolbar(tb, ACTION_NAME)
        except Exception: pass
    ok = kw.attach_action_to_toolbar(toolbar_name, ACTION_NAME)
    _msg("attach_action_to_toolbar('%s') -> %s" % (toolbar_name, ok))
    if ok:
        try: kw.detach_action_from_toolbar(toolbar_name, ACTION_NAME)
        except Exception: pass
        kw.attach_action_to_toolbar(toolbar_name, ACTION_NAME)
        _msg("re-attached to '%s' (tail)" % toolbar_name)
    return ok

class _Handler(kw.action_handler_t):
    def __init__(self):
        kw.action_handler_t.__init__(self)
    def activate(self, ctx):
        if os.name != "nt":
            kw.warning("Плагин рассчитан на Windows.")
            return 1
        if not _is_socket_running():
            start_socket_thread()
        _restart_solution_instance()
        # После рестарта VS: подберём окно в фокус
        focus_vs_window()
        return 1
    def update(self, ctx):
        return kw.AST_ENABLE_ALWAYS if os.name == "nt" else kw.AST_DISABLE

def _register_action():
    try: kw.unregister_action(ACTION_NAME)
    except Exception: pass
    # icon_id = -1 (без иконки), чтобы не пользоваться «магическим» 199
    desc = kw.action_desc_t(ACTION_NAME, ACTION_LABEL, _Handler(), HOTKEY, ACTION_TIP, -1)
    ok = kw.register_action(desc)
    _msg("register_action -> %s" % ok)
    return ok

# ---------- Экшен Link: тумблер приёма адресов ----------

LINK_ACTION_NAME  = "czar.vs.afterpdb.link"
LINK_ACTION_LABEL = "Link"
LINK_ACTION_TIP   = "Вкл/выкл приём адресов от VS (порт 8888)"
LINK_HOTKEY       = "Ctrl+Alt+L"

class _LinkHandler(kw.action_handler_t):
    def __init__(self):
        kw.action_handler_t.__init__(self)
    def activate(self, ctx):
        if _is_socket_running():
            stop_socket_thread()
        else:
            start_socket_thread()
        return 1
    def update(self, ctx):
        return kw.AST_ENABLE_ALWAYS

def _register_link_action():
    try: kw.unregister_action(LINK_ACTION_NAME)
    except Exception: pass
    desc = kw.action_desc_t(LINK_ACTION_NAME, LINK_ACTION_LABEL, _LinkHandler(), LINK_HOTKEY, LINK_ACTION_TIP, -1)
    ok = kw.register_action(desc)
    _msg("register_action(Link) -> %s" % ok)
    return ok

# ---------- Таймер: ожидание появления тулбара PdbGen ----------

def _probe_and_attach():
    if _state["attached_to"] == PDBGEN_TOOLBAR:
        _state["timer_live"] = False
        return -1
    ok = _attach_last(PDBGEN_TOOLBAR)
    if ok:
        _state["attached_to"] = PDBGEN_TOOLBAR
        _state["timer_live"] = False
        return -1
    _state["tries"] += 1
    if _state["tries"] >= MAX_TRIES:
        _attach_last(FALLBACK_TB)
        _state["timer_live"] = False
        return -1
    return TIMER_STEP_MS

# ---------- UI Hooks ----------

class _Hooks(kw.UI_Hooks):
    def ready_to_run(self):
        _register_action()
        _register_link_action()
        if not _attach_last(PDBGEN_TOOLBAR):
            if not _state["timer_live"]:
                kw.register_timer(TIMER_STEP_MS, _probe_and_attach)
                _state["timer_live"] = True
        return 0  # ВАЖНО: IDA 9.2 ожидает int, а не None

_hooks = None

# ---------- Жизненный цикл плагина ----------

def _install():
    global _hooks
    if AUTO_CONNECT_ON_START:
        start_socket_thread()
    else:
        _msg("Автоподключение к VS при старте: выключено (Ctrl+Alt+L — включить).")
    install_view_hook()
    _register_action()
    _register_link_action()
    if not _attach_last(PDBGEN_TOOLBAR):
        kw.register_timer(TIMER_STEP_MS, _probe_and_attach)
        _state["timer_live"] = True
    _hooks = _Hooks()
    try:
        _hooks.hook()
        _msg("UI_Hooks.hook -> ok")
    except Exception as e:
        _msg(f"UI_Hooks.hook -> fail: {e}")
    _msg("Плагин готов. Даблклик — отправка в VS. Кнопка 'VS' — рестарт re.sln.")

def _uninstall():
    global _hooks
    try:
        for tb in (PDBGEN_TOOLBAR, FALLBACK_TB):
            kw.detach_action_from_toolbar(tb, ACTION_NAME)
    except Exception:
        pass
    try: kw.unregister_action(ACTION_NAME)
    except Exception: pass
    try: kw.unregister_action(LINK_ACTION_NAME)
    except Exception: pass
    if _hooks:
        try: _hooks.unhook()
        except Exception: pass
    _hooks = None
    uninstall_view_hook()
    stop_socket_thread()
    _msg("uninstall complete")

class vs_after_pdbgen_restart_plugin_t(idaapi.plugin_t):
    flags = idaapi.PLUGIN_PROC
    comment = "VS button after PdbGen (restart + socket)"
    help = "Restarts only the VS instance with re.sln and keeps socket link"
    wanted_name = "VS after PdbGen (restart)"
    wanted_hotkey = HOTKEY
    def init(self): _install(); return idaapi.PLUGIN_KEEP
    def run(self, arg): kw.process_ui_action(ACTION_NAME)
    def term(self): _uninstall()

def PLUGIN_ENTRY():
    return vs_after_pdbgen_restart_plugin_t()
