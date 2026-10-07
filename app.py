"""
Sortirovka — графическое приложение для сортировки товаров по моделям.

Запуск для проверки:   python app.py
Сборка в exe:          см. build.bat
"""
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog

import sortirovka_core as core

# Перетаскивание файлов (drag & drop) — необязательная библиотека.
# Если её нет, приложение всё равно работает: файл выбирается по клику.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except Exception:
    HAS_DND = False

# ---------- цвета и шрифт ----------
BG = "#0E1013"            # фон окна
CARD = "#161A20"          # фон карточек
CARD_HOVER = "#1B2030"    # карточка при наведении
BORDER = "#2A303A"        # рамки
ACCENT = "#7C83FF"        # главный акцент
ACCENT_HOVER = "#939AFF"
ICON_BG = "#222850"
TEXT = "#E8EAF0"
MUTED = "#8B93A3"
OK = "#3DDC97"
WARN = "#F5B849"
ERR = "#FF6B6B"
FONT = "Segoe UI"


# ---------- вспомогательные функции ----------

def plural(n, forms):
    """Склонение: plural(3, ('короб', 'короба', 'коробов')) -> 'короба'."""
    n10, n100 = n % 10, n % 100
    if n10 == 1 and n100 != 11:
        return forms[0]
    if 2 <= n10 <= 4 and not 12 <= n100 <= 14:
        return forms[1]
    return forms[2]


def fmt(n):
    """3440 -> '3 440'."""
    return f"{n:,}".replace(",", " ")


def short(name, limit=34):
    """Длинное имя файла сокращаем посередине, чтобы влезло в карточку."""
    if len(name) <= limit:
        return name
    half = limit // 2
    return name[:half - 1] + "…" + name[-half:]


def open_path(path):
    """Открывает файл программой по умолчанию (Excel)."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except OSError:
        pass


def reveal_in_folder(path):
    """Открывает папку и выделяет в ней файл."""
    try:
        if sys.platform.startswith("win"):
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path])
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(path)])
    except OSError:
        pass


def round_rect(canvas, x1, y1, x2, y2, r, **kw):
    """Прямоугольник со скруглёнными углами (на Canvas такого нет из коробки)."""
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
           x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
           x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return canvas.create_polygon(pts, smooth=True, **kw)


# ---------- виджеты ----------

class FlatButton(tk.Label):
    """Плоская кнопка (Label с наведением), выглядит одинаково на любой Windows."""

    def __init__(self, master, text, command, k, primary=True):
        self._bg = ACCENT if primary else BORDER
        self._hover = ACCENT_HOVER if primary else "#363D4A"
        super().__init__(
            master, text=text, bg=self._bg, fg="#FFFFFF" if primary else TEXT,
            font=(FONT, 10, "bold"), padx=int(16 * k), pady=int(7 * k), cursor="hand2",
        )
        self.bind("<Enter>", lambda e: self.config(bg=self._hover))
        self.bind("<Leave>", lambda e: self.config(bg=self._bg))
        self.bind("<Button-1>", lambda e: command())


class DropZone(tk.Canvas):
    """Большая зона: сюда перетаскивают файл или кликают, чтобы выбрать."""

    def __init__(self, master, k, on_click, has_dnd):
        super().__init__(master, width=int(584 * k), height=int(230 * k),
                         bg=BG, highlightthickness=0, cursor="hand2")
        self.k = k
        self.on_click = on_click
        self.has_dnd = has_dnd
        self.hover = False
        self.busy = False
        self.file_name = ""
        self.phase = 0.0           # положение бегущей полоски (0..1)
        self._animating = False

        self.bind("<Configure>", lambda e: self.redraw())
        self.bind("<Enter>", lambda e: self.set_hover(True))
        self.bind("<Leave>", lambda e: self.set_hover(False))
        self.bind("<Button-1>", lambda e: self.on_click())

    # --- состояния ---
    def set_hover(self, value):
        self.hover = value
        self.redraw()

    def set_file(self, name):
        self.file_name = name
        self.redraw()

    def set_busy(self, value):
        self.busy = value
        if value and not self._animating:
            self._animating = True
            self._tick()
        self.config(cursor="watch" if value else "hand2")
        self.redraw()

    def _tick(self):
        """Анимация бегущей полоски, пока идёт обработка."""
        if not self.busy:
            self._animating = False
            return
        self.phase = (self.phase + 0.025) % 1.0
        self.redraw()
        self.after(30, self._tick)

    # --- рисование ---
    def redraw(self):
        self.delete("all")
        w, h, k = self.winfo_width(), self.winfo_height(), self.k
        if w < 20:                       # окно ещё не показано
            return
        cx, cy = w / 2, h / 2
        active = self.hover and not self.busy

        style = dict(fill=CARD_HOVER if active else CARD,
                     outline=ACCENT if (active or self.busy) else BORDER, width=2)
        if not self.busy:
            style["dash"] = (7, 5)
        round_rect(self, 2, 2, w - 2, h - 2, 24 * k, **style)

        if self.busy:
            self.create_text(cx, cy - 22 * k, text="Обрабатываю…",
                             font=(FONT, 15, "bold"), fill=TEXT)
            self.create_text(cx, cy + 6 * k, text=short(self.file_name, 46),
                             font=(FONT, 10), fill=MUTED)
            # дорожка и бегущая полоска
            x0, x1 = cx - 130 * k, cx + 130 * k
            y0, y1 = cy + 38 * k, cy + 44 * k
            self.create_rectangle(x0, y0, x1, y1, fill=BORDER, outline="")
            bar = 90 * k
            start = x0 - bar + (x1 - x0 + bar) * self.phase
            a, b = max(x0, start), min(x1, start + bar)
            if b > a:
                self.create_rectangle(a, y0, b, y1, fill=ACCENT, outline="")
            return

        # иконка: кружок со стрелкой вверх
        iy = cy - 42 * k
        r = 30 * k
        self.create_oval(cx - r, iy - r, cx + r, iy + r, fill=ICON_BG, outline="")
        line = dict(width=max(2, int(3 * k)), fill=ACCENT, capstyle="round")
        self.create_line(cx, iy + 12 * k, cx, iy - 12 * k, **line)
        self.create_line(cx - 10 * k, iy - 2 * k, cx, iy - 12 * k, **line)
        self.create_line(cx + 10 * k, iy - 2 * k, cx, iy - 12 * k, **line)

        if self.has_dnd:
            title, sub = "Перетащите Excel-файлы сюда", "можно сразу несколько · или нажмите, чтобы выбрать"
        else:
            title, sub = "Нажмите, чтобы выбрать Excel-файлы", "можно выбрать сразу несколько · формат .xlsx"
        self.create_text(cx, cy + 28 * k, text=title, font=(FONT, 15, "bold"), fill=TEXT)
        self.create_text(cx, cy + 56 * k, text=sub, font=(FONT, 10), fill=MUTED)


# ---------- приложение ----------

class App:
    def __init__(self, root, dnd):
        self.root = root
        self.dnd = dnd
        self.busy = False
        self.q = queue.Queue()

        # коэффициент масштаба экрана (на ноутбуках с 125–150% размеры должны расти)
        self.k = root.winfo_fpixels("1i") / 96.0

        root.title("Sortirovka")
        root.configure(bg=BG)
        root.geometry(f"{self.px(660)}x{self.px(580)}")
        root.minsize(self.px(560), self.px(520))

        self._build_ui()

        if dnd:
            root.drop_target_register(DND_FILES)
            root.dnd_bind("<<DropEnter>>", self._on_drop_enter)
            root.dnd_bind("<<DropLeave>>", self._on_drop_leave)
            root.dnd_bind("<<Drop>>", self._on_drop)

        if sys.platform.startswith("win"):
            root.after(50, self._dark_titlebar)

    def px(self, value):
        return int(round(value * self.k))

    # --- интерфейс ---
    def _build_ui(self):
        pad = self.px(38)

        header = tk.Frame(self.root, bg=BG)
        header.pack(fill="x", padx=pad, pady=(self.px(32), 0))
        tk.Label(header, text="SORTIROVKA", bg=BG, fg=ACCENT,
                 font=(FONT, 9, "bold")).pack(anchor="w")
        tk.Label(header, text="Сортировка по моделям", bg=BG, fg=TEXT,
                 font=(FONT, 24, "bold")).pack(anchor="w", pady=(self.px(2), 0))
        tk.Label(header, text="Загрузите файлы Excel (любого из двух форматов) — получите сводку по моделям",
                 bg=BG, fg=MUTED, font=(FONT, 10)).pack(anchor="w", pady=(self.px(4), 0))

        self.zone = DropZone(self.root, self.k, self.choose_files, self.dnd)
        self.zone.pack(fill="x", padx=pad, pady=(self.px(22), 0))

        # сюда добавляются карточки с результатом
        self.results = tk.Frame(self.root, bg=BG)
        self.results.pack(fill="x", padx=pad, pady=(self.px(18), 0))

        tk.Label(self.root, bg=BG, fg=MUTED, font=(FONT, 9),
                 text="Один общий файл сохраняется рядом с исходным: имя_sortirovka.xlsx"
                 ).pack(side="bottom", pady=(0, self.px(16)))

    def _dark_titlebar(self):
        """Тёмная шапка окна на Windows 10/11 (если не получится — не страшно)."""
        try:
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            value = ctypes.c_int(1)
            for attr in (20, 19):
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, attr, ctypes.byref(value), ctypes.sizeof(value))
        except Exception:
            pass

    # --- получение файлов ---
    def choose_files(self):
        if self.busy:
            return
        paths = filedialog.askopenfilenames(
            title="Выберите один или несколько Excel-файлов",
            filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")])
        if paths:
            self.start(list(paths))

    def _on_drop_enter(self, event):
        self.zone.set_hover(True)
        return event.action

    def _on_drop_leave(self, event):
        self.zone.set_hover(False)
        return event.action

    def _on_drop(self, event):
        self.zone.set_hover(False)
        # event.data — строка вида "{C:/мои файлы/a.xlsx} C:/b.xlsx"; splitlist разбирает её правильно
        self.start(list(self.root.tk.splitlist(event.data)))
        return event.action

    # --- обработка ---
    def start(self, paths):
        if self.busy:
            return
        self.clear_results()

        # один и тот же файл дважды не нужен — иначе цифры в сводке задвоятся
        unique, seen = [], set()
        for p in paths:
            key = os.path.normcase(os.path.abspath(p))
            if key not in seen:
                seen.add(key)
                unique.append(p)

        # если среди файлов есть неподходящий — ничего не считаем,
        # иначе итоговая сводка получилась бы неполной, и это легко не заметить
        invalid = [p for p in unique
                   if not p.lower().endswith(".xlsx") or os.path.basename(p).startswith("~$")]
        if invalid:
            for p in invalid:
                self.add_error(os.path.basename(p) or p,
                               "Нужен файл Excel в формате .xlsx. Если у вас .xls — сохраните его "
                               "как «Книга Excel (.xlsx)». Уберите этот файл и загрузите всё заново.")
            self._fit_height()
            return

        self.busy = True
        self.zone.set_file(os.path.basename(unique[0]))
        self.zone.set_busy(True)
        threading.Thread(target=self._worker, args=(unique,), daemon=True).start()
        self.root.after(80, self._poll)

    def _worker(self, paths):
        """Работает в отдельном потоке, чтобы окно не зависало."""
        def progress(name, i, total):
            label = name if total == 1 else f"{name}  ({i} из {total})"
            self.q.put(("start", label))

        try:
            for res in core.process_files(paths, progress):
                self.q.put(("ok", res))
        except core.SortError as e:
            self.q.put(("err", ("Не удалось обработать", str(e))))
        except Exception as e:
            self.q.put(("err", ("Не удалось обработать", f"Что-то пошло не так: {e}")))
        self.q.put(("done", None))

    def _poll(self):
        """Забирает сообщения из потока и обновляет окно (в основном потоке)."""
        try:
            while True:
                kind, data = self.q.get_nowait()
                if kind == "start":
                    self.zone.set_file(data)
                elif kind == "ok":
                    self.add_result(data)
                elif kind == "err":
                    self.add_error(*data)
                elif kind == "done":
                    self.busy = False
                    self.zone.set_busy(False)
                    self._fit_height()
                    return
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    # --- карточки результата ---
    def clear_results(self):
        for child in self.results.winfo_children():
            child.destroy()

    def _make_card(self, dot_color):
        card = tk.Frame(self.results, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="x", pady=(0, self.px(10)))
        card.columnconfigure(1, weight=1)
        tk.Label(card, text="●", fg=dot_color, bg=CARD, font=(FONT, 12)
                 ).grid(row=0, column=0, sticky="n", padx=(self.px(16), self.px(10)), pady=self.px(14))
        body = tk.Frame(card, bg=CARD)
        body.grid(row=0, column=1, sticky="ew", pady=self.px(12))
        return card, body

    def add_result(self, res):
        card, body = self._make_card(OK)
        name = os.path.basename(res["output"])

        tk.Label(body, text=short(name), bg=CARD, fg=TEXT,
                 font=(FONT, 11, "bold"), anchor="w").pack(anchor="w")

        files = f'{res["files"]} {plural(res["files"], ("файл", "файла", "файлов"))}'
        models = f'{res["models"]} {plural(res["models"], ("модель", "модели", "моделей"))}'
        boxes = f'{fmt(res["boxes"])} {plural(res["boxes"], ("короб", "короба", "коробов"))}'
        pairs = f'{fmt(res["pairs"])} {plural(res["pairs"], ("пара", "пары", "пар"))}'
        tk.Label(body, text="  ·  ".join(([files] if res["files"] > 1 else []) + [models, boxes, pairs]), bg=CARD, fg=MUTED,
                 font=(FONT, 9), anchor="w").pack(anchor="w", pady=(self.px(2), 0))

        tk.Label(body, text="Формат: " + core.KIND_NAMES[res["kind"]] + (f'  ·  {res["period"]}' if res.get("period") else ""),
                 bg=CARD, fg=MUTED,
                 font=(FONT, 9), anchor="w").pack(anchor="w")

        # замечания: красным/жёлтым — то, что стоит проверить, серым — просто информация
        wrap = self.px(330)
        color = WARN if res["warn_level"] == "bad" else MUTED
        for text in res["warnings"][:6]:
            tk.Label(body, text=text, bg=CARD, fg=color, font=(FONT, 9), anchor="w",
                     justify="left", wraplength=wrap).pack(anchor="w", pady=(self.px(3), 0))
        if len(res["warnings"]) > 6:
            tk.Label(body, text=f'…и ещё замечаний: {len(res["warnings"]) - 6} (они в файле, лист «Поставки»)',
                     bg=CARD, fg=MUTED, font=(FONT, 9), anchor="w").pack(anchor="w", pady=(self.px(3), 0))

        buttons = tk.Frame(card, bg=CARD)
        buttons.grid(row=0, column=2, padx=(self.px(8), self.px(16)), pady=self.px(12))
        out = res["output"]
        FlatButton(buttons, "Открыть", lambda: open_path(out), self.k).pack(side="left")
        FlatButton(buttons, "В папке", lambda: reveal_in_folder(out), self.k,
                   primary=False).pack(side="left", padx=(self.px(8), 0))

    def add_error(self, name, message):
        card, body = self._make_card(ERR)
        tk.Label(body, text=short(name), bg=CARD, fg=TEXT,
                 font=(FONT, 11, "bold"), anchor="w").pack(anchor="w")
        tk.Label(body, text=message, bg=CARD, fg=ERR, font=(FONT, 9),
                 anchor="w", justify="left", wraplength=self.px(470)
                 ).pack(anchor="w", pady=(self.px(3), 0))

    def _fit_height(self):
        """Если карточек много — увеличиваем окно, чтобы всё было видно."""
        self.root.update_idletasks()
        need = self.root.winfo_reqheight()
        if need > self.root.winfo_height():
            self.root.geometry(f"{self.root.winfo_width()}x{need}")


def main():
    # чёткий текст на экранах с масштабом 125–150% (только Windows)
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    dnd = HAS_DND
    try:
        root = TkinterDnD.Tk() if dnd else tk.Tk()
    except Exception:
        # не загрузилась библиотека перетаскивания — работаем без неё
        dnd = False
        root = tk.Tk()

    App(root, dnd)
    root.mainloop()


if __name__ == "__main__":
    main()
