# -*- coding: utf-8 -*-
# Mahashe Env Manager (CLI + GUI)
# GUI: customtkinter, стиль 1:1 Mahashe Install Hub (тёмно-синяя тема по умолчанию)
# Функции:
# - Вкладка PATH: поиск, добавление папки, редактирование, удаление, удаление дубликатов, удаление несуществующих, сохранение в PATH (User + Machine если есть админ)
# - Вкладка Переменные среды: список переменных (User / Machine / Both), двойной клик -> большое окно редактирования, создание/удаление
# CLI:
#   -h / -help                     справка
#   -gui                           открыть GUI
#   -list [--scope user|machine|both]           (также поддерживается --score как алиас)
#   -get  <NAME> [--scope user|machine]        (также поддерживается --score как алиас)
#   -set  <NAME> <VALUE...> [--scope user|machine]  VALUE может быть без кавычек, до флагов
#   -del  <NAME> [--scope user|machine]        (также поддерживается --score как алиас)
#   -addpath <PATH...> [--scope user|machine|both]
#   -rmpath  <PATH...> [--scope user|machine|both]
#   -deduppath [--scope user|machine|both]
#   -prunepath [--scope user|machine|both]
#
# Windows only.

import os
import sys
import ctypes
import subprocess
from collections import Counter
from typing import Dict, List, Tuple, Optional

try:
    import winreg
except Exception:
    winreg = None

# PyInstaller: обеспечить доступ к Tcl/Tk при onefile до импорта customtkinter
if getattr(sys, "frozen", False):
    _base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    os.environ.setdefault("TCL_LIBRARY", os.path.join(_base, "tcl", "tcl8.6"))
    os.environ.setdefault("TK_LIBRARY", os.path.join(_base, "tcl", "tk8.6"))

import customtkinter as ctk


APP_TITLE = "Mahashe Path_Tool Helper"


# =========================
# CLI helpers
# =========================

def eprint(msg: str) -> None:
    try:
        sys.stderr.write(str(msg).rstrip() + "\n")
        sys.stderr.flush()
    except Exception:
        pass


def okprint(msg: str) -> None:
    try:
        sys.stdout.write(str(msg).rstrip() + "\n")
        sys.stdout.flush()
    except Exception:
        pass


def exit_with(code: int, message: Optional[str] = None) -> int:
    if message:
        if code == 0:
            okprint(message)
        else:
            eprint(message)
    return code


# =========================
# THEME (как Install Hub)
# =========================

THEMES = {
    "Тёмно-синяя": {
        "BG": "#0c1533",
        "CARD": "#0f1e47",
        "BORDER": "#223265",
        "BLUE": "#1d4ed8",
        "TEXT": "#ffffff",
        "MUTED": "#cbd5e1",
        "OK": "#2bd576",
        "WARN": "#ffb454",
        "BAD": "#e04b59",
    },
}

DEFAULT_THEME = "Тёмно-синяя"


def _theme():
    return THEMES[DEFAULT_THEME]


# =========================
# ADMIN / BROADCAST
# =========================

def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def elevate_if_needed():
    if os.name != "nt":
        return
    try:
        if ctypes.windll.shell32.IsUserAnAdmin():
            return
    except Exception:
        return

    params = subprocess.list2cmdline(sys.argv[1:])
    exe = sys.executable
    workdir = os.path.dirname(os.path.abspath(sys.argv[0]))
    rc = ctypes.windll.shell32.ShellExecuteW(None, "runas", exe, params, workdir, 1)
    if isinstance(rc, int) and rc > 32:
        sys.exit(0)


def broadcast_env_change():
    if os.name != "nt":
        return
    try:
        HWND_BROADCAST = 0xFFFF
        WM_SETTINGCHANGE = 0x1A
        SMTO_ABORTIFHUNG = 0x0002
        ctypes.windll.user32.SendMessageTimeoutW(
            HWND_BROADCAST,
            WM_SETTINGCHANGE,
            0,
            "Environment",
            SMTO_ABORTIFHUNG,
            5000,
            None
        )
    except Exception:
        pass


# =========================
# REGISTRY HELPERS
# =========================

HKCU_ENV = (winreg.HKEY_CURRENT_USER, r"Environment") if winreg else None
HKLM_ENV = (winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment") if winreg else None


def _require_windows_registry():
    if winreg is None or os.name != "nt":
        raise RuntimeError("Требуется Windows (winreg недоступен).")


def _open_env_key(scope: str, access: int):
    _require_windows_registry()
    if scope == "user":
        root, path = HKCU_ENV
        return winreg.CreateKeyEx(root, path, 0, access)
    if scope == "machine":
        root, path = HKLM_ENV
        return winreg.CreateKeyEx(root, path, 0, access)
    raise ValueError("scope must be user|machine")


def list_env(scope: str) -> Dict[str, Tuple[str, int]]:
    """
    returns {name: (value, reg_type)}
    """
    _require_windows_registry()
    out: Dict[str, Tuple[str, int]] = {}
    access = winreg.KEY_READ
    with _open_env_key(scope, access) as k:
        i = 0
        while True:
            try:
                name, val, vtype = winreg.EnumValue(k, i)
            except OSError:
                break
            out[str(name)] = (str(val), int(vtype))
            i += 1
    return out


def get_env(scope: str, name: str) -> Optional[Tuple[str, int]]:
    _require_windows_registry()
    try:
        with _open_env_key(scope, winreg.KEY_READ) as k:
            val, vtype = winreg.QueryValueEx(k, name)
            return (str(val), int(vtype))
    except Exception:
        return None


def set_env(scope: str, name: str, value: str) -> None:
    _require_windows_registry()
    if scope == "machine" and not is_admin():
        raise PermissionError("Требуются права администратора для записи в MACHINE.")
    vtype = winreg.REG_EXPAND_SZ if "%" in (value or "") else winreg.REG_SZ
    with _open_env_key(scope, winreg.KEY_SET_VALUE) as k:
        winreg.SetValueEx(k, name, 0, vtype, value)
    broadcast_env_change()


def delete_env(scope: str, name: str) -> None:
    _require_windows_registry()
    if scope == "machine" and not is_admin():
        raise PermissionError("Требуются права администратора для удаления из MACHINE.")
    try:
        with _open_env_key(scope, winreg.KEY_SET_VALUE) as k:
            winreg.DeleteValue(k, name)
    except FileNotFoundError:
        pass
    except Exception:
        pass
    broadcast_env_change()


# =========================
# PATH LOGIC
# =========================

def _split_path(raw: str) -> List[str]:
    parts = []
    for p in (raw or "").split(os.pathsep):
        p = p.strip()
        if p:
            parts.append(p)
    return parts


def _join_path(parts: List[str]) -> str:
    return os.pathsep.join(parts)


def read_path(scope: str) -> List[str]:
    cur = get_env(scope, "Path")
    if not cur:
        return []
    return _split_path(cur[0])


def write_path(scope: str, parts: List[str]) -> None:
    set_env(scope, "Path", _join_path(parts))


def expand_exists(p: str) -> bool:
    try:
        ex = os.path.expandvars(p)
        return os.path.exists(ex)
    except Exception:
        return False


def dedup_keep_first(parts: List[str]) -> List[str]:
    seen = set()
    out = []
    for p in parts:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


def prune_nonexistent(parts: List[str]) -> List[str]:
    return [p for p in parts if expand_exists(p)]


def add_path_once(parts: List[str], new_p: str) -> List[str]:
    if new_p and new_p not in parts:
        parts.append(new_p)
    return parts


def rm_path_exact(parts: List[str], target: str) -> List[str]:
    return [p for p in parts if p != target]


# =========================
# TOAST (Install Hub style)
# =========================

class ToastManager:
    def __init__(self, root, theme):
        self.root = root
        self.theme = theme
        self.active: List[ctk.CTkToplevel] = []

    def show(self, title: str, text: str, ms=3000, width=380):
        win = ctk.CTkToplevel(self.root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.attributes("-alpha", 0.0)
        win.configure(fg_color=self.theme["CARD"])

        frame = ctk.CTkFrame(
            win,
            fg_color=self.theme["CARD"],
            corner_radius=12,
            border_width=1,
            border_color=self.theme["BORDER"],
        )
        frame.pack(fill="both", expand=True)

        title_lbl = ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=13, weight="bold"), text_color=self.theme["TEXT"])
        body_lbl = ctk.CTkLabel(frame, text=text, font=ctk.CTkFont(size=12), text_color=self.theme["MUTED"], justify="left")
        title_lbl.pack(anchor="w", padx=14, pady=(10, 2))
        body_lbl.pack(anchor="w", padx=14, pady=(0, 12))

        win.update_idletasks()
        h = frame.winfo_reqheight() + 2
        x = self.root.winfo_screenwidth() - width - 16
        y = self.root.winfo_screenheight() - 16 - (len(self.active) * (h + 10)) - h
        win.geometry(f"{width}x{h}+{x}+{y}")

        self.active.append(win)
        self._fade(win, 0.0, 0.98, step=0.08, delay=20)
        win.after(ms, lambda: self._dismiss(win))

    def _fade(self, win, a_from, a_to, step=0.05, delay=25):
        a = a_from + step
        if (step > 0 and a >= a_to) or (step < 0 and a <= a_to):
            try:
                win.attributes("-alpha", a_to)
            except Exception:
                pass
            return
        try:
            win.attributes("-alpha", a)
        except Exception:
            return
        win.after(delay, lambda: self._fade(win, a, a_to, step, delay))

    def _dismiss(self, win):
        if not win.winfo_exists():
            return
        try:
            cur = float(win.attributes("-alpha"))
        except Exception:
            cur = 1.0
        self._fade(win, cur, 0.0, step=-0.08, delay=20)
        win.after(220, lambda: self._destroy(win))

    def _destroy(self, win):
        if win in self.active:
            self.active.remove(win)
        try:
            win.destroy()
        except Exception:
            pass


# =========================
# UI DIALOGS
# =========================

class BigEditDialog(ctk.CTkToplevel):
    def __init__(self, parent, theme, title: str, name: str = "", value: str = "", name_editable: bool = True):
        super().__init__(parent)
        self.theme = theme
        self.result = None

        self.title(title)
        self.geometry("760x420")
        self.resizable(False, False)
        self.configure(fg_color=theme["BG"])
        self.grab_set()
        self.transient(parent)

        card = ctk.CTkFrame(self, fg_color=theme["CARD"], corner_radius=15, border_width=1, border_color=theme["BORDER"])
        card.pack(fill="both", expand=True, padx=14, pady=14)
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=16, weight="bold"), text_color=theme["TEXT"]).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(12, 8)
        )

        ctk.CTkLabel(card, text="Имя:", font=ctk.CTkFont(size=13, weight="bold"), text_color=theme["TEXT"]).grid(
            row=1, column=0, sticky="w", padx=14, pady=(0, 6)
        )
        self.ent_name = ctk.CTkEntry(card, height=34)
        self.ent_name.grid(row=2, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 12))
        self.ent_name.insert(0, name or "")
        if not name_editable:
            self.ent_name.configure(state="disabled")

        ctk.CTkLabel(card, text="Значение:", font=ctk.CTkFont(size=13, weight="bold"), text_color=theme["TEXT"]).grid(
            row=3, column=0, sticky="w", padx=14, pady=(0, 6)
        )

        self.txt_val = ctk.CTkTextbox(card, height=220, corner_radius=12, border_width=1, border_color=theme["BORDER"])
        self.txt_val.grid(row=4, column=0, columnspan=2, sticky="nsew", padx=14, pady=(0, 10))
        self.txt_val.insert("1.0", value or "")

        btns = ctk.CTkFrame(card, fg_color=theme["CARD"])
        btns.grid(row=5, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 14))
        btns.grid_columnconfigure(0, weight=1)

        self.btn_ok = ctk.CTkButton(btns, text="Сохранить", corner_radius=15, command=self._ok, width=140)
        self.btn_cancel = ctk.CTkButton(btns, text="Отмена", corner_radius=15, command=self._cancel, width=120)

        self.btn_ok.pack(side="left")
        self.btn_cancel.pack(side="right")

        self._center(parent)

    def _center(self, parent):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _ok(self):
        name = (self.ent_name.get() or "").strip()
        val = self.txt_val.get("1.0", "end").rstrip("\n")
        if not name:
            return
        self.result = (name, val)
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# =========================
# MAIN APP (customtkinter)
# =========================

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.th = _theme()
        self.configure(fg_color=self.th["BG"])
        self.title(APP_TITLE + ("" if is_admin() else " — Без прав администратора"))
        self.geometry("1040x760")
        self.resizable(False, False)

        self.toaster = ToastManager(self, self.th)

        top = ctk.CTkFrame(self, fg_color=self.th["BG"], corner_radius=0)
        top.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(
            top, text="Mahashe Env Manager",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=self.th["TEXT"]
        ).pack(side="left")

        ctk.CTkButton(
            top, text="Обновить всё",
            corner_radius=15,
            command=self.refresh_all
        ).pack(side="right")

        self.tabs = ctk.CTkTabview(
            self,
            corner_radius=15,
            border_width=1,
            border_color=self.th["BORDER"],
            fg_color=self.th["BG"],
            segmented_button_fg_color=self.th["CARD"],
            segmented_button_selected_color=self.th["BLUE"],
            segmented_button_selected_hover_color="#2563eb",
            segmented_button_unselected_color=self.th["CARD"],
            segmented_button_unselected_hover_color="#142352",
            text_color="white",
        )
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.tab_path = self.tabs.add("PATH")
        self.tab_env = self.tabs.add("Переменные среды")

        self._build_path_tab()
        self._build_env_tab()

        self.refresh_all()

    # ---------------- PATH TAB ----------------

    def _card(self, parent):
        f = ctk.CTkFrame(parent, fg_color=self.th["CARD"], corner_radius=15, border_width=1, border_color=self.th["BORDER"])
        f._is_card = True
        return f

    def _build_path_tab(self):
        root = ctk.CTkFrame(self.tab_path, fg_color=self.th["BG"], corner_radius=0)
        root.pack(fill="both", expand=True, padx=10, pady=10)

        top_card = self._card(root)
        top_card.pack(fill="x", padx=6, pady=(0, 10))
        top_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(top_card, text="PATH редактор", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(12, 6)
        )

        self.path_scope = ctk.CTkOptionMenu(top_card, values=["both", "user", "machine"])
        self.path_scope.set("both")
        self.path_scope.grid(row=0, column=1, sticky="e", padx=14, pady=(12, 6))

        self.path_search = ctk.CTkEntry(top_card, placeholder_text="Поиск по PATH")
        self.path_search.grid(row=1, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 12))
        self.path_search.bind("<KeyRelease>", lambda e: self._path_apply_filter())

        self.path_list = ctk.CTkScrollableFrame(
            root,
            corner_radius=15,
            fg_color=self.th["BG"],
            border_width=0,
            scrollbar_button_color=self.th["BG"],
            scrollbar_button_hover_color=self.th["BG"],
            scrollbar_fg_color=self.th["BG"],
        )
        self.path_list.pack(fill="both", expand=True, padx=6, pady=(0, 10))

        bottom = self._card(root)
        bottom.pack(fill="x", padx=6, pady=(0, 0))

        b1 = ctk.CTkButton(bottom, text="Добавить папку", corner_radius=15, command=self.path_add_folder)
        b2 = ctk.CTkButton(bottom, text="Удалить выбранные", corner_radius=15, command=self.path_delete_selected)
        b3 = ctk.CTkButton(bottom, text="Удалить дубликаты", corner_radius=15, command=self.path_dedup)
        b4 = ctk.CTkButton(bottom, text="Удалить несуществующие", corner_radius=15, command=self.path_prune)
        b5 = ctk.CTkButton(bottom, text="Сохранить", corner_radius=15, command=self.path_apply)

        b1.pack(side="left", padx=12, pady=12)
        b2.pack(side="left", padx=(0, 10), pady=12)
        b3.pack(side="left", padx=(0, 10), pady=12)
        b4.pack(side="left", padx=(0, 10), pady=12)
        b5.pack(side="right", padx=12, pady=12)

        self._path_rows = []
        self._path_items: List[str] = []

    def _path_load(self) -> List[str]:
        scope = self.path_scope.get()
        items: List[str] = []
        if scope == "user":
            items = read_path("user")
        elif scope == "machine":
            items = read_path("machine")
        else:
            u = read_path("user")
            m = read_path("machine")
            seen = set()
            out = []
            for p in u + m:
                if p not in seen:
                    seen.add(p)
                    out.append(p)
            items = out
        return items

    def _path_rebuild(self):
        for w in self.path_list.winfo_children():
            w.destroy()
        self._path_rows.clear()

        items = self._path_items
        counts = Counter(items)
        flt = (self.path_search.get() or "").strip().lower()

        for p in items:
            if flt and flt not in p.lower():
                continue

            valid = expand_exists(p)
            dup = counts[p]

            row = ctk.CTkFrame(self.path_list, fg_color=self.th["CARD"], corner_radius=14, border_width=1, border_color=self.th["BORDER"])
            row.pack(fill="x", padx=6, pady=6)
            row.grid_columnconfigure(2, weight=1)

            var = ctk.BooleanVar(value=False)
            cb = ctk.CTkCheckBox(row, text="", variable=var, width=24)
            cb.grid(row=0, column=0, padx=(12, 8), pady=10)

            dot = ctk.CTkLabel(row, text="●", width=12)
            color = self.th["WARN"] if dup > 1 else (self.th["OK"] if valid else self.th["BAD"])
            dot.configure(text_color=color, font=ctk.CTkFont(size=14, weight="bold"))
            dot.grid(row=0, column=1, padx=(0, 8), pady=10)

            ent = ctk.CTkEntry(row, height=34)
            ent.grid(row=0, column=2, sticky="ew", padx=(0, 8), pady=10)
            ent.insert(0, p)
            ent.configure(state="readonly")

            btn = ctk.CTkButton(row, text="⋮", width=44, corner_radius=12, command=lambda pp=p: self.path_row_menu(pp))
            btn.grid(row=0, column=3, padx=(0, 12), pady=10)

            self._path_rows.append((p, var))

    def _path_apply_filter(self):
        self._path_rebuild()

    def path_row_menu(self, p: str):
        win = ctk.CTkToplevel(self)
        win.title("PATH")
        win.geometry("360x180")
        win.resizable(False, False)
        win.configure(fg_color=self.th["BG"])
        win.grab_set()
        win.transient(self)

        card = self._card(win)
        card.pack(fill="both", expand=True, padx=14, pady=14)

        ctk.CTkLabel(card, text="Действия", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=14, pady=(12, 8))
        ctk.CTkLabel(card, text=p, wraplength=320, text_color=self.th["MUTED"]).pack(anchor="w", padx=14, pady=(0, 10))

        def do_edit():
            win.destroy()
            self.path_edit(p)

        def do_del():
            win.destroy()
            self._path_items = rm_path_exact(self._path_items, p)
            self._path_rebuild()

        btns = ctk.CTkFrame(card, fg_color=self.th["CARD"])
        btns.pack(fill="x", padx=14, pady=(0, 14))
        ctk.CTkButton(btns, text="Редактировать", corner_radius=15, command=do_edit).pack(side="left")
        ctk.CTkButton(btns, text="Удалить", corner_radius=15, command=do_del).pack(side="right")

        win.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = win.winfo_width()
        h = win.winfo_height()
        win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def path_add_folder(self):
        from tkinter import filedialog
        d = filedialog.askdirectory(parent=self, title="Выберите папку для PATH")
        if not d:
            return
        self._path_items = add_path_once(self._path_items, d)
        self._path_rebuild()
        self.toaster.show("PATH", "Путь добавлен в список", ms=2200)

    def path_edit(self, p: str):
        dlg = BigEditDialog(self, self.th, "Редактирование пути (PATH)", name="PathItem", value=p, name_editable=False)
        self.wait_window(dlg)
        if not dlg.result:
            return
        _, new_val = dlg.result
        new_val = (new_val or "").strip().splitlines()[0].strip()
        if not new_val:
            return
        try:
            idx = self._path_items.index(p)
            self._path_items[idx] = new_val
            self._path_rebuild()
            self.toaster.show("PATH", "Путь обновлён", ms=2200)
        except ValueError:
            pass

    def path_delete_selected(self):
        selected = [p for (p, v) in self._path_rows if v.get()]
        if not selected:
            self.toaster.show("PATH", "Нечего удалять", ms=2000)
            return
        s = set(selected)
        self._path_items = [p for p in self._path_items if p not in s]
        self._path_rebuild()
        self.toaster.show("PATH", f"Удалено: {len(selected)}", ms=2200)

    def path_dedup(self):
        self._path_items = dedup_keep_first(self._path_items)
        self._path_rebuild()
        self.toaster.show("PATH", "Дубликаты удалены", ms=2200)

    def path_prune(self):
        before = len(self._path_items)
        self._path_items = prune_nonexistent(self._path_items)
        removed = before - len(self._path_items)
        self._path_rebuild()
        self.toaster.show("PATH", f"Удалено несуществующих: {removed}", ms=2400)

    def path_apply(self):
        scope = self.path_scope.get()
        try:
            if scope == "user":
                write_path("user", self._path_items)
                self.toaster.show("PATH", "Сохранено: USER", ms=2400)
            elif scope == "machine":
                if not is_admin():
                    self.toaster.show("PATH", "Нужен админ для MACHINE", ms=2600)
                    return
                write_path("machine", self._path_items)
                self.toaster.show("PATH", "Сохранено: MACHINE", ms=2400)
            else:
                write_path("user", self._path_items)
                if is_admin():
                    write_path("machine", self._path_items)
                    self.toaster.show("PATH", "Сохранено: USER + MACHINE", ms=2600)
                else:
                    self.toaster.show("PATH", "Сохранено: USER (MACHINE требует админ)", ms=3000)
        except PermissionError:
            self.toaster.show("PATH", "Отказано в доступе (админ)", ms=2800)
        except Exception as e:
            self.toaster.show("PATH", f"Ошибка: {e}", ms=3400)

    # ---------------- ENV TAB ----------------

    def _build_env_tab(self):
        root = ctk.CTkFrame(self.tab_env, fg_color=self.th["BG"], corner_radius=0)
        root.pack(fill="both", expand=True, padx=10, pady=10)

        top_card = self._card(root)
        top_card.pack(fill="x", padx=6, pady=(0, 10))
        top_card.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(top_card, text="Переменные среды", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(12, 6)
        )

        self.env_scope = ctk.CTkOptionMenu(top_card, values=["user", "machine", "both"], command=lambda _: self.env_reload())
        self.env_scope.set("user")
        self.env_scope.grid(row=0, column=1, sticky="e", padx=14, pady=(12, 6))

        self.env_search = ctk.CTkEntry(top_card, placeholder_text="Поиск по переменным (имя/значение)")
        self.env_search.grid(row=1, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 12))
        self.env_search.bind("<KeyRelease>", lambda e: self.env_rebuild())

        self.env_list = ctk.CTkScrollableFrame(
            root,
            corner_radius=15,
            fg_color=self.th["BG"],
            border_width=0,
            scrollbar_button_color=self.th["BG"],
            scrollbar_button_hover_color=self.th["BG"],
            scrollbar_fg_color=self.th["BG"],
        )
        self.env_list.pack(fill="both", expand=True, padx=6, pady=(0, 10))

        bottom = self._card(root)
        bottom.pack(fill="x", padx=6, pady=0)

        ctk.CTkButton(bottom, text="Создать переменную", corner_radius=15, command=self.env_create).pack(
            side="left", padx=12, pady=12
        )
        ctk.CTkButton(bottom, text="Обновить список", corner_radius=15, command=self.env_reload).pack(
            side="right", padx=12, pady=12
        )

        self._env_data = []  # list of tuples (scope, name, value, regtype)

    def env_reload(self):
        self._env_data.clear()
        scope = self.env_scope.get()

        def load_one(sc):
            try:
                d = list_env(sc)
                for k, (v, t) in d.items():
                    self._env_data.append((sc, k, v, t))
            except PermissionError:
                pass
            except Exception:
                pass

        if scope in ("user", "both"):
            load_one("user")
        if scope in ("machine", "both"):
            load_one("machine")

        self._env_data.sort(key=lambda x: (x[1].lower(), x[0]))
        self.env_rebuild()
        self.toaster.show("Переменные среды", "Список обновлён", ms=1700)

    def env_rebuild(self):
        for w in self.env_list.winfo_children():
            w.destroy()

        flt = (self.env_search.get() or "").strip().lower()

        for sc, name, val, _t in self._env_data:
            if flt and (flt not in name.lower()) and (flt not in (val or "").lower()):
                continue

            card = ctk.CTkFrame(self.env_list, fg_color=self.th["CARD"], corner_radius=15, border_width=1, border_color=self.th["BORDER"])
            card.pack(fill="x", padx=6, pady=6)
            card.grid_columnconfigure(1, weight=1)

            badge = "USER" if sc == "user" else "MACHINE"
            badge_color = self.th["BLUE"] if sc == "user" else "#6b7280"

            b = ctk.CTkLabel(card, text=badge, text_color=self.th["TEXT"])
            b.configure(font=ctk.CTkFont(size=11, weight="bold"))
            b.grid(row=0, column=0, padx=(12, 10), pady=10, sticky="w")

            lbl = ctk.CTkLabel(card, text=name, font=ctk.CTkFont(size=13, weight="bold"), text_color=self.th["TEXT"])
            lbl.grid(row=0, column=1, padx=(0, 10), pady=(10, 2), sticky="w")

            short = (val or "").replace("\r", "").replace("\n", " ")
            if len(short) > 140:
                short = short[:140] + "…"
            sub = ctk.CTkLabel(card, text=short, font=ctk.CTkFont(size=12), text_color=self.th["MUTED"], wraplength=720, justify="left")
            sub.grid(row=1, column=1, padx=(0, 10), pady=(0, 10), sticky="w")

            for w in (card, lbl, sub, b):
                w.bind("<Double-Button-1>", lambda _e, s=sc, n=name: self.env_edit_open(s, n))

            btn = ctk.CTkButton(card, text="⋮", width=44, corner_radius=12, command=lambda s=sc, n=name: self.env_row_menu(s, n))
            btn.grid(row=0, column=2, rowspan=2, padx=(0, 12), pady=10, sticky="e")

            try:
                b.configure(text_color="white")
                b._text_label.configure(bg=badge_color)  # type: ignore[attr-defined]
            except Exception:
                pass

    def env_row_menu(self, scope: str, name: str):
        win = ctk.CTkToplevel(self)
        win.title("Переменная")
        win.geometry("380x200")
        win.resizable(False, False)
        win.configure(fg_color=self.th["BG"])
        win.grab_set()
        win.transient(self)

        card = self._card(win)
        card.pack(fill="both", expand=True, padx=14, pady=14)

        ctk.CTkLabel(card, text="Действия", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=14, pady=(12, 8))
        ctk.CTkLabel(card, text=f"{scope.upper()}  •  {name}", text_color=self.th["MUTED"]).pack(anchor="w", padx=14, pady=(0, 10))

        def do_edit():
            win.destroy()
            self.env_edit_open(scope, name)

        def do_del():
            win.destroy()
            self.env_delete(scope, name)

        btns = ctk.CTkFrame(card, fg_color=self.th["CARD"])
        btns.pack(fill="x", padx=14, pady=(0, 14))
        ctk.CTkButton(btns, text="Редактировать", corner_radius=15, command=do_edit).pack(side="left")
        ctk.CTkButton(btns, text="Удалить", corner_radius=15, command=do_del).pack(side="right")

        win.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = win.winfo_width()
        h = win.winfo_height()
        win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def env_create(self):
        dlg = BigEditDialog(self, self.th, "Создать переменную среды", name="", value="", name_editable=True)
        self.wait_window(dlg)
        if not dlg.result:
            return
        name, val = dlg.result
        scope = self.env_scope.get()
        try:
            if scope == "both":
                set_env("user", name, val)
                if is_admin():
                    set_env("machine", name, val)
                    self.toaster.show("Переменные", f"Создано: USER + MACHINE: {name}", ms=2600)
                else:
                    self.toaster.show("Переменные", f"Создано: USER (MACHINE требует админ): {name}", ms=3200)
            else:
                if scope == "machine" and not is_admin():
                    self.toaster.show("Переменные", "Нужен админ для MACHINE", ms=2800)
                    return
                set_env(scope, name, val)
                self.toaster.show("Переменные", f"Создано/обновлено: {scope.upper()}: {name}", ms=2600)
        except Exception as e:
            self.toaster.show("Переменные", f"Ошибка: {e}", ms=3600)

        self.env_reload()

    def env_edit_open(self, scope: str, name: str):
        cur = get_env(scope, name)
        if not cur:
            self.toaster.show("Переменные", "Переменная не найдена", ms=2400)
            return
        value, _t = cur

        dlg = BigEditDialog(self, self.th, f"Редактирование переменной ({scope.upper()})", name=name, value=value, name_editable=False)
        self.wait_window(dlg)
        if not dlg.result:
            return
        _n, new_val = dlg.result

        try:
            if scope == "machine" and not is_admin():
                self.toaster.show("Переменные", "Нужен админ для MACHINE", ms=2800)
                return
            set_env(scope, name, new_val)
            self.toaster.show("Переменные", f"Сохранено: {scope.upper()}: {name}", ms=2400)
        except Exception as e:
            self.toaster.show("Переменные", f"Ошибка: {e}", ms=3600)

        self.env_reload()

    def env_delete(self, scope: str, name: str):
        try:
            if scope == "machine" and not is_admin():
                self.toaster.show("Переменные", "Нужен админ для MACHINE", ms=2800)
                return
            delete_env(scope, name)
            self.toaster.show("Переменные", f"Удалено: {scope.upper()}: {name}", ms=2400)
        except Exception as e:
            self.toaster.show("Переменные", f"Ошибка: {e}", ms=3600)
        self.env_reload()

    # ---------------- COMMON ----------------

    def refresh_all(self):
        try:
            self._path_items = self._path_load()
            self._path_rebuild()
        except Exception as e:
            self.toaster.show("PATH", f"Ошибка чтения PATH: {e}", ms=3800)

        try:
            self.env_reload()
        except Exception as e:
            self.toaster.show("Переменные", f"Ошибка чтения: {e}", ms=3800)


# =========================
# CLI
# =========================

def print_help():
    exe = os.path.basename(sys.argv[0] or "env_manager.py")
    txt = f"""
{APP_TITLE}

ENV:
  {exe} -set  <NAME> <VALUE...> [--scope user|machine]   создать/обновить переменную
  {exe} -del  <NAME> [--scope user|machine]              удалить переменную
  {exe} -get  <NAME> [--scope user|machine]              получить значение
  {exe} -list [--scope user|machine|both]                вывести список

PATH:
  {exe} -addpath <PATH...> [--scope user|machine|both]   добавить элемент в PATH
  {exe} -rmpath  <PATH...> [--scope user|machine|both]   удалить точное совпадение из PATH
  {exe} -deduppath [--scope user|machine|both]           удалить дубликаты в PATH
  {exe} -prunepath [--scope user|machine|both]           удалить несуществующие из PATH

GUI:
  {exe} -gui

HELP:
  {exe} -h
  {exe} -help

Алиасы:
  --score работает как --scope

Коды возврата:
  0  OK
  1  Не найдено/ошибка
  2  Неверные аргументы/неподдерживаемая команда
  5  Нет прав (нужен запуск от администратора)
"""
    print(txt.strip())


def _arg_value_any(args: List[str], keys: List[str], default: Optional[str] = None) -> Optional[str]:
    for key in keys:
        if key in args:
            i = args.index(key)
            if i + 1 < len(args):
                return args[i + 1]
    return default


def _scope_from_args(args: List[str], default: str) -> str:
    sc = _arg_value_any(args, ["--scope", "--score"], default) or default
    sc = sc.lower().strip()
    if sc not in ("user", "machine", "both"):
        return default
    return sc


def _take_value_until_flags(tokens: List[str]) -> str:
    out = []
    for t in tokens:
        if t.startswith("--"):
            break
        out.append(t)
    return " ".join(out).strip()


def cli_run() -> Optional[int]:
    _require_windows_registry()

    args = sys.argv[1:]
    if not args:
        return None

    a0 = args[0].lower()

    if a0 in ("-h", "-help", "--help", "/?"):
        print_help()
        return 0

    if a0 == "-gui":
        return None

    # ENV
    if a0 == "-list":
        sc = _scope_from_args(args, "both")
        try:
            if sc == "both":
                u = list_env("user")
                okprint("=== USER ===")
                for k in sorted(u.keys(), key=str.lower):
                    okprint(f"{k}={u[k][0]}")
                okprint("")
                okprint("=== MACHINE ===")
                if is_admin():
                    m = list_env("machine")
                    for k in sorted(m.keys(), key=str.lower):
                        okprint(f"{k}={m[k][0]}")
                else:
                    okprint("(нет доступа без админа)")
                    return 5
                return 0

            if sc == "machine" and not is_admin():
                return exit_with(5, "ERROR: Нет прав. Запусти от администратора для scope=machine.")

            d = list_env(sc)
            for k in sorted(d.keys(), key=str.lower):
                okprint(f"{k}={d[k][0]}")
            return 0
        except Exception as e:
            return exit_with(1, f"ERROR: list failed: {e}")

    if a0 == "-get":
        if len(args) < 2:
            return exit_with(2, "ERROR: -get требует <NAME>.")
        name = args[1]
        sc = _scope_from_args(args, "user")
        if sc == "both":
            return exit_with(2, "ERROR: -get не поддерживает scope=both. Используй user или machine.")
        if sc == "machine" and not is_admin():
            return exit_with(5, "ERROR: Нет прав. Запусти от администратора для scope=machine.")
        try:
            v = get_env(sc, name)
            if not v:
                return exit_with(1, f"ERROR: Переменная не найдена: {sc}:{name}")
            okprint(v[0])
            return 0
        except Exception as e:
            return exit_with(1, f"ERROR: get failed: {e}")

    if a0 == "-set":
        if len(args) < 3:
            return exit_with(2, "ERROR: -set требует <NAME> <VALUE...>.")
        name = args[1]
        sc = _scope_from_args(args, "user")
        if sc == "both":
            return exit_with(2, "ERROR: -set не поддерживает scope=both. Используй user или machine.")
        value = _take_value_until_flags(args[2:])
        if not value:
            return exit_with(2, "ERROR: -set: пустое VALUE.")
        if sc == "machine" and not is_admin():
            return exit_with(5, "ERROR: Нет прав. Запусти от администратора для scope=machine.")
        try:
            set_env(sc, name, value)
            return exit_with(0, f"OK: set {sc}:{name}")
        except PermissionError as e:
            return exit_with(5, f"ERROR: {e}")
        except Exception as e:
            return exit_with(1, f"ERROR: set failed: {e}")

    if a0 == "-del":
        if len(args) < 2:
            return exit_with(2, "ERROR: -del требует <NAME>.")
        name = args[1]
        sc = _scope_from_args(args, "user")
        if sc == "both":
            return exit_with(2, "ERROR: -del не поддерживает scope=both. Используй user или machine.")
        if sc == "machine" and not is_admin():
            return exit_with(5, "ERROR: Нет прав. Запусти от администратора для scope=machine.")
        try:
            delete_env(sc, name)
            return exit_with(0, f"OK: del {sc}:{name}")
        except PermissionError as e:
            return exit_with(5, f"ERROR: {e}")
        except Exception as e:
            return exit_with(1, f"ERROR: del failed: {e}")

    # PATH
    if a0 in ("-addpath", "-rmpath"):
        if len(args) < 2:
            return exit_with(2, f"ERROR: {a0} требует <PATH...>.")
        sc = _scope_from_args(args, "both")
        p = _take_value_until_flags(args[1:])
        if not p:
            return exit_with(2, f"ERROR: {a0}: пустой PATH.")

        targets = []
        if sc in ("user", "both"):
            targets.append("user")
        if sc in ("machine", "both"):
            targets.append("machine")

        wrote_machine = False
        try:
            for t in targets:
                if t == "machine" and not is_admin():
                    continue
                parts = read_path(t)
                parts = add_path_once(parts, p) if a0 == "-addpath" else rm_path_exact(parts, p)
                write_path(t, parts)
                if t == "machine":
                    wrote_machine = True

            if sc in ("machine", "both") and not is_admin():
                return exit_with(5, "ERROR: Нет прав. Запусти от администратора для scope=machine.")
            return exit_with(0, f"OK: {a0} ({sc})" + ("" if wrote_machine or sc == "user" else ""))
        except PermissionError:
            return exit_with(5, "ERROR: Нет прав. Запусти от администратора для записи в MACHINE.")
        except Exception as e:
            return exit_with(1, f"ERROR: {a0} failed: {e}")

    if a0 in ("-deduppath", "-prunepath"):
        sc = _scope_from_args(args, "both")

        targets = []
        if sc in ("user", "both"):
            targets.append("user")
        if sc in ("machine", "both"):
            targets.append("machine")

        try:
            for t in targets:
                if t == "machine" and not is_admin():
                    continue
                parts = read_path(t)
                parts = dedup_keep_first(parts) if a0 == "-deduppath" else prune_nonexistent(parts)
                write_path(t, parts)

            if sc in ("machine", "both") and not is_admin():
                return exit_with(5, "ERROR: Нет прав. Запусти от администратора для scope=machine.")
            return exit_with(0, f"OK: {a0} ({sc})")
        except PermissionError:
            return exit_with(5, "ERROR: Нет прав. Запусти от администратора для записи в MACHINE.")
        except Exception as e:
            return exit_with(1, f"ERROR: {a0} failed: {e}")

    return exit_with(2, f"ERROR: Команда не поддерживается: {args[0]}")


# =========================
# ENTRYPOINT
# =========================

def main():
    if os.name != "nt" or winreg is None:
        eprint("ERROR: Требуется Windows.")
        sys.exit(1)

    # CLI first
    try:
        rc = cli_run()
        if rc is not None:
            sys.exit(rc)
    except Exception as e:
        eprint(f"ERROR: CLI crashed: {e}")
        # фолбэк: если CLI упал — открываем GUI

    elevate_if_needed()
    App().mainloop()


if __name__ == "__main__":
    main()
