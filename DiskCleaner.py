# disk_cleaner_full.py
"""
Disk Cleaner (tkinter)
- Сканирует рекурсивно все файлы на выбранном диске
- Показывает: Название, Размер, Тип, Системный (Да/Нет), Процесс(если бл.), % от диска
- Сортировка по столбцам, фильтр по типу
- Удаление (с проверкой системных файлов и завершением процессов)
- Создание ярлыка на рабочем столе (pywin32)
"""

import os
import sys
import shutil
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
import psutil
import ctypes
import math

# Для ярлыка (опционально)
try:
    import pythoncom
    import win32com.client
except Exception:
    pythoncom = None
    win32com = None

# ---------- Constants / Helpers ----------
FILE_ATTRIBUTE_SYSTEM = 0x4

def human_size(size):
    try:
        size = float(size)
    except:
        return "0 B"
    for unit in ['B','KB','MB','GB','TB','PB']:
        if size < 1024.0:
            return f"{size:,.1f} {unit}"
        size /= 1024.0
    return f"{size:.1f} PB"

def list_drives():
    drives = []
    try:
        bitmask = ctypes.cdll.kernel32.GetLogicalDrives()
        for i in range(26):
            if bitmask & (1 << i):
                drives.append(f"{chr(65+i)}:\\")
    except Exception:
        pass
    return drives

def detect_type(path: Path):
    try:
        if path.is_dir():
            return "Папка"
    except Exception:
        # могут быть проблемы доступа
        return "Папка"
    ext = path.suffix.lower()
    if ext in [".exe", ".msi"]:
        return "Программа"
    elif ext in [".txt", ".docx", ".pdf", ".xls", ".xlsx", ".ppt", ".pptx", ".rtf", ".odt"]:
        return "Документ"
    else:
        return "Файл"

def is_system_file(path):
    try:
        attrs = ctypes.windll.kernel32.GetFileAttributesW(str(path))
        if attrs == -1:
            return False
        return bool(attrs & FILE_ATTRIBUTE_SYSTEM)
    except Exception:
        return False

def find_processes_locking(path):
    result = []
    for proc in psutil.process_iter(['pid','name']):
        try:
            for f in proc.open_files():
                try:
                    if os.path.abspath(f.path) == os.path.abspath(path):
                        result.append((proc.pid, proc.name()))
                except Exception:
                    pass
        except (psutil.AccessDenied, psutil.NoSuchProcess):
            continue
        except Exception:
            continue
    return result

def get_folder_size_bytes(path):
    total = 0
    for root, dirs, files in os.walk(path, onerror=lambda e: None):
        for f in files:
            try:
                fp = os.path.join(root, f)
                total += os.path.getsize(fp)
            except Exception:
                pass
    return total

# ---------- GUI ----------
class DiskCleaner(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Disk Cleaner — безопасный очиститель")
        self.geometry("1500x800")
        self.configure(bg="#f3f3f3")

        # Header
        header = ttk.Label(self, text="Disk Cleaner — безопасный очиститель",
                           font=("Segoe UI", 18, "bold"))
        header.pack(side="top", fill="x", pady=8)

        # Controls
        ctrl_frame = ttk.Frame(self)
        ctrl_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(ctrl_frame, text="Диск:").pack(side="left")
        self.disk_combo = ttk.Combobox(ctrl_frame, values=list_drives(), width=10)
        self.disk_combo.pack(side="left", padx=6)
        self.btn_scan = ttk.Button(ctrl_frame, text="Сканировать полностью", command=self.scan_disk_full)
        self.btn_scan.pack(side="left", padx=6)
        self.btn_refresh = ttk.Button(ctrl_frame, text="Обновить диски", command=self.refresh_drives)
        self.btn_refresh.pack(side="left", padx=6)

        ttk.Label(ctrl_frame, text="Фильтр:").pack(side="left", padx=(12,2))
        self.type_filter = ttk.Combobox(ctrl_frame, values=["Все","Файл","Программа","Документ","Папка"], width=12)
        self.type_filter.current(0)
        self.type_filter.pack(side="left")
        self.type_filter.bind("<<ComboboxSelected>>", lambda e: self.apply_filter())

        self.btn_shortcut = ttk.Button(ctrl_frame, text="Создать ярлык", command=self.create_shortcut)
        self.btn_shortcut.pack(side="left", padx=12)

        # Treeview columns
        columns = ("name","size","percent","type","system","process")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        self.tree.heading("name", text="Название", command=lambda: self.sort_tree("name", False))
        self.tree.heading("size", text="Размер", command=lambda: self.sort_tree("size", True))
        self.tree.heading("percent", text="% от диска", command=lambda: self.sort_tree("percent", True))
        self.tree.heading("type", text="Тип", command=lambda: self.sort_tree("type", False))
        self.tree.heading("system", text="Системный", command=lambda: self.sort_tree("system", False))
        self.tree.heading("process", text="Процесс (если заблокирован)")
        self.tree.column("name", width=600, anchor="w")
        self.tree.column("size", width=140, anchor="e")
        self.tree.column("percent", width=100, anchor="e")
        self.tree.column("type", width=120, anchor="center")
        self.tree.column("system", width=100, anchor="center")
        self.tree.column("process", width=420, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        # Bottom controls
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=10, pady=6)
        self.btn_delete = ttk.Button(bottom, text="Удалить выбранное", command=self.delete_selected)
        self.btn_delete.pack(side="left")
        self.btn_export = ttk.Button(bottom, text="Экспорт списка в CSV", command=self.export_csv)
        self.btn_export.pack(side="left", padx=6)
        self.status = ttk.Label(bottom, text="Готово.")
        self.status.pack(side="right")

        # Store items: list of tuples (path, meta) where meta is dict {name,size_bytes,type,system,process,percent}
        self.all_items = []
        self.disk_total_bytes = 1  # чтобы деление на 0 не произошло

    def refresh_drives(self):
        self.disk_combo["values"] = list_drives()

    def scan_disk_full(self):
        drive = self.disk_combo.get()
        if not drive:
            messagebox.showwarning("Ошибка", "Выберите диск (например, C:\\)!")
            return
        # clear
        self.tree.delete(*self.tree.get_children())
        self.all_items.clear()
        self.status.config(text="Сбор информации о диске...")
        # compute disk total
        try:
            usage = shutil.disk_usage(drive)
            self.disk_total_bytes = usage.total
        except Exception:
            self.disk_total_bytes = 1
        # start scanning thread
        t = threading.Thread(target=self._scan_thread, args=(drive,), daemon=True)
        t.start()

    def _scan_thread(self, drive):
        # Walk all files. For speed we enumerate files; for folders we can decide not to insert every folder,
        # but user asked to scan by file weights -> we insert files primarily. We'll still mark folders when desirable.
        count = 0
        try:
            for root, dirs, files in os.walk(drive, onerror=lambda e: None):
                # files first
                for f in files:
                    path = os.path.join(root, f)
                    try:
                        size_b = os.path.getsize(path)
                    except Exception:
                        size_b = 0
                    type_name = detect_type(Path(path))
                    system_flag = is_system_file(path)
                    percent = (size_b / self.disk_total_bytes) * 100 if self.disk_total_bytes else 0.0
                    meta = {
                        "name": f,
                        "size_bytes": size_b,
                        "type": type_name,
                        "system": "Да" if system_flag else "Нет",
                        "process": "",
                        "percent": percent
                    }
                    self.all_items.append((path, meta))
                    # insert to tree (string-format)
                    self.tree.insert("", "end", iid=path,
                                     values=(meta["name"], human_size(size_b), f"{meta['percent']:.3f}%", meta["type"], meta["system"], ""))
                    count += 1
                    if count % 500 == 0:
                        self.status.config(text=f"Просканировано файлов: {count} (последняя папка: {root})")
                # optionally also include empty folders or top-level folder sizes if wanted — skipped for performance
                self.status.config(text=f"Просканировано файлов: {count} (последняя папка: {root})")
        except Exception as e:
            # возможные ошибки доступа просто покажем в статусе
            self.status.config(text=f"Сканирование прервано: {e}")
            return
        self.status.config(text=f"Сканирование завершено. Всего файлов: {len(self.all_items)}")
        # применим текущий фильтр (чтобы отображать только нужные элементы)
        self.apply_filter()

    def apply_filter(self):
        filter_type = self.type_filter.get()
        self.tree.delete(*self.tree.get_children())
        for path, meta in self.all_items:
            if filter_type == "Все" or meta["type"] == filter_type:
                self.tree.insert("", "end", iid=path,
                                 values=(meta["name"], human_size(meta["size_bytes"]),
                                         f"{meta['percent']:.3f}%", meta["type"], meta["system"], meta["process"]))

    def sort_tree(self, col, numeric_desc):
        # Build list of (key, iid) to sort
        items = self.tree.get_children("")
        def key_fn(iid):
            val = self.tree.set(iid, col)
            if col == "size":
                # parse human size back to bytes using stored meta
                for path, meta in self.all_items:
                    if path == iid:
                        return meta["size_bytes"]
                return 0
            if col == "percent":
                for path, meta in self.all_items:
                    if path == iid:
                        return meta["percent"]
                return 0
            if col == "name":
                return self.tree.set(iid, "name").lower()
            if col == "type":
                return self.tree.set(iid, "type")
            if col == "system":
                return 0 if self.tree.set(iid, "system") == "Да" else 1
            return self.tree.set(iid, col)
        reverse = numeric_desc
        sorted_items = sorted(items, key=key_fn, reverse=reverse)
        for index, iid in enumerate(sorted_items):
            self.tree.move(iid, "", index)
        # toggle next time
        # rebind heading to invert order when clicked again
        # (we will set command to call with opposite)
        if col == "size":
            self.tree.heading("size", command=lambda: self.sort_tree("size", not numeric_desc))
        elif col == "percent":
            self.tree.heading("percent", command=lambda: self.sort_tree("percent", not numeric_desc))
        elif col == "name":
            self.tree.heading("name", command=lambda: self.sort_tree("name", not numeric_desc))
        elif col == "type":
            self.tree.heading("type", command=lambda: self.sort_tree("type", not numeric_desc))
        elif col == "system":
            self.tree.heading("system", command=lambda: self.sort_tree("system", not numeric_desc))

    def delete_selected(self):
        items = self.tree.selection()
        if not items:
            messagebox.showinfo("Инфо", "Ничего не выбрано.")
            return
        # check system files
        for iid in items:
            if self.tree.set(iid, "system") == "Да":
                messagebox.showwarning("Предупреждение", "Выбраны системные файлы — удаление запрещено.")
                return
        paths = list(items)
        if not messagebox.askyesno("Подтвердите", f"Удалить {len(paths)} элементов?"):
            return
        t = threading.Thread(target=self._delete_thread, args=(paths,), daemon=True)
        t.start()

    def _delete_thread(self, paths):
        for path in paths:
            p = Path(path)
            self.status.config(text=f"Удаление: {p}")
            try:
                if p.is_file():
                    os.remove(p)
                elif p.is_dir():
                    shutil.rmtree(p)
                # remove from tree and all_items
                try:
                    self.tree.delete(path)
                except Exception:
                    pass
                self.all_items = [(pth, meta) for pth, meta in self.all_items if pth != path]
            except Exception as e:
                # try detect processes locking
                owners = find_processes_locking(path)
                if owners:
                    proc_text = ", ".join(f"{name}({pid})" for pid, name in owners)
                    try:
                        self.tree.set(path, "process", proc_text)
                    except Exception:
                        pass
                    # ask user to terminate
                    if messagebox.askyesno("Файл занят", f"Файл {path} занят процессами:\n{proc_text}\nЗавершить процессы?"):
                        for pid, name in owners:
                            try:
                                psutil.Process(pid).terminate()
                            except Exception:
                                pass
                        # try delete again
                        try:
                            if p.exists():
                                if p.is_file():
                                    os.remove(p)
                                else:
                                    shutil.rmtree(p)
                                try:
                                    self.tree.delete(path)
                                except Exception:
                                    pass
                                self.all_items = [(pth, meta) for pth, meta in self.all_items if pth != path]
                                self.status.config(text=f"Удалено: {path}")
                        except Exception as e2:
                            messagebox.showerror("Ошибка", f"Не удалось удалить {path} после завершения процессов:\n{e2}")
                else:
                    messagebox.showerror("Ошибка", f"Не удалось удалить {path}:\n{e}")
        self.status.config(text="Удаление завершено.")

    def create_shortcut(self):
        if pythoncom is None or win32com is None:
            messagebox.showerror("Ошибка", "Нужен pywin32 (pip install pywin32)")
            return
        try:
            desktop = Path.home() / "Desktop"
            shortcut_path = desktop / "Disk Cleaner.lnk"
            target = sys.executable
            script = str(Path(__file__).resolve())
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(str(shortcut_path))
            shortcut.Targetpath = target
            shortcut.Arguments = f'"{script}"'
            shortcut.WorkingDirectory = str(Path(script).parent)
            shortcut.IconLocation = target
            shortcut.save()
            messagebox.showinfo("Ярлык создан", f"Ярлык на рабочем столе:\n{shortcut_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать ярлык:\n{e}")

    def export_csv(self):
        try:
            import csv
            out = Path.cwd() / "disk_scan_export.csv"
            with open(out, "w", newline="", encoding="utf-8") as fh:
                w = csv.writer(fh)
                w.writerow(["path","name","size_bytes","size_human","percent","type","system","process"])
                for path, meta in self.all_items:
                    w.writerow([path, meta["name"], meta["size_bytes"], human_size(meta["size_bytes"]),
                                f"{meta['percent']:.6f}", meta["type"], meta["system"], meta["process"]])
            messagebox.showinfo("Экспорт", f"CSV сохранён: {out}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать CSV:\n{e}")

# ---------- MAIN ----------
if __name__ == "__main__":
    # Требуем Windows
    if os.name != 'nt':
        tk.messagebox.showerror("Unsupported", "Это приложение предназначено для Windows.")
        sys.exit(1)
    app = DiskCleaner()
    app.mainloop()
