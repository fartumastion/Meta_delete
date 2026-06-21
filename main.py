import os
import shutil
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

SUPPORTED_EXT = {".docx", ".pptx", ".xlsx"}

NAMESPACES = {
    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'dcterms': 'http://purl.org/dc/terms/',
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
}

for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

PROPS_MAP = {
    'creator': ('core', 'dc:creator', 'Автор'),
    'lastModifiedBy': ('core', 'cp:lastModifiedBy', 'Последнее изменение (кем)'),
    'title': ('core', 'dc:title', 'Заголовок'),
    'subject': ('core', 'dc:subject', 'Тема'),
    'keywords': ('core', 'cp:keywords', 'Ключевые слова'),
    'description': ('core', 'dc:description', 'Описание'),
    'category': ('core', 'cp:category', 'Категория'),
    'contentStatus': ('core', 'cp:contentStatus', 'Статус содержимого'),
    'language': ('core', 'dc:language', 'Язык'),
    'revision': ('core', 'cp:revision', 'Номер версии'),
    'created': ('core', 'dcterms:created', 'Дата создания'),
    'modified': ('core', 'dcterms:modified', 'Дата изменения'),
    'lastPrinted': ('core', 'cp:lastPrinted', 'Последняя печать'),
    'Company': ('app', 'ep:Company', 'Компания'),
    'Manager': ('app', 'ep:Manager', 'Менеджер'),
}

EXT_COLORS = {
    ".docx": ("#1565C0", "#1976D2"),
    ".pptx": ("#6A1B9A", "#7B1FA2"),
    ".xlsx": ("#1B5E20", "#2E7D32"),
}

def get_file_properties(filepath: str) -> dict:
    props = {}
    for k, v in PROPS_MAP.items():
        props[k] = {"label": v[2], "value": ""}

    try:
        with zipfile.ZipFile(filepath, 'r') as zin:
            if 'docProps/core.xml' in zin.namelist():
                core_xml = zin.read('docProps/core.xml')
                root = ET.fromstring(core_xml)
                for k, v in PROPS_MAP.items():
                    if v[0] == 'core':
                        node = root.find(v[1], NAMESPACES)
                        if node is not None and node.text:
                            props[k]['value'] = node.text
                            
            if 'docProps/app.xml' in zin.namelist():
                app_xml = zin.read('docProps/app.xml')
                root = ET.fromstring(app_xml)
                for k, v in PROPS_MAP.items():
                    if v[0] == 'app':
                        node = root.find(v[1], NAMESPACES)
                        if node is not None and node.text:
                            props[k]['value'] = node.text
    except Exception as e:
        print(f"Ошибка чтения {filepath}: {e}")
        
    return props

def clear_file_properties(filepath: str, attrs_to_clear: list) -> int:
    cleared_count = 0
    fd, temp_path = tempfile.mkstemp(suffix=".zip")
    os.close(fd)

    try:
        with zipfile.ZipFile(filepath, 'r') as zin:
            with zipfile.ZipFile(temp_path, 'w') as zout:
                for item in zin.infolist():
                    if item.filename == 'docProps/core.xml':
                        xml_data = zin.read(item.filename)
                        root = ET.fromstring(xml_data)
                        modified = False
                        
                        for attr in attrs_to_clear:
                            if attr in PROPS_MAP and PROPS_MAP[attr][0] == 'core':
                                tag = PROPS_MAP[attr][1]
                                node = root.find(tag, NAMESPACES)
                                if node is not None:
                                    if attr in ('created', 'modified', 'lastPrinted', 'lastModifiedBy'):
                                        root.remove(node)
                                    else:
                                        node.text = ""
                                    modified = True
                                    cleared_count += 1
                                    
                        if modified:
                            zout.writestr(item, ET.tostring(root, encoding='utf-8', xml_declaration=True))
                        else:
                            zout.writestr(item, xml_data)

                    elif item.filename == 'docProps/app.xml':
                        xml_data = zin.read(item.filename)
                        root = ET.fromstring(xml_data)
                        modified = False
                        
                        for attr in attrs_to_clear:
                            if attr in PROPS_MAP and PROPS_MAP[attr][0] == 'app':
                                tag = PROPS_MAP[attr][1]
                                node = root.find(tag, NAMESPACES)
                                if node is not None:
                                    node.text = ""
                                    modified = True
                                    cleared_count += 1
                                    
                        if modified:
                            zout.writestr(item, ET.tostring(root, encoding='utf-8', xml_declaration=True))
                        else:
                            zout.writestr(item, xml_data)
                    else:
                        zout.writestr(item, zin.read(item.filename))

        shutil.move(temp_path, filepath)
    except Exception as e:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        raise e
        
    return cleared_count

class MetaCleanerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("MetaCleaner — Массовая очистка метаданных")
        self.geometry("1200x780")
        self.minsize(950, 680)
        self._files: list[str] = []
        self._current_file: str | None = None
        self._prop_vars: dict[str, ctk.BooleanVar] = {}
        self._file_buttons: dict[str, ctk.CTkButton] = {}
        self._build_layout()

    def _build_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_left_panel()
        self._build_right_panel()

    def _build_left_panel(self):
        panel = ctk.CTkFrame(self, width=300, corner_radius=0, fg_color=("gray92", "gray14"))
        panel.grid(row=0, column=0, sticky="nsew")
        panel.grid_propagate(False)
        panel.grid_rowconfigure(3, weight=1)
        panel.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(panel, fg_color=("gray85", "gray18"), corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="📁  Файлы", font=ctk.CTkFont(size=15, weight="bold")).grid(row=0, column=0, padx=16, pady=14, sticky="w")

        btn_add_row = ctk.CTkFrame(panel, fg_color="transparent")
        btn_add_row.grid(row=1, column=0, padx=12, pady=(10, 5), sticky="ew")
        btn_add_row.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(btn_add_row, text="＋ Файлы", command=self._add_files, height=30, corner_radius=6).grid(row=0, column=0, padx=(0, 3), sticky="ew")
        ctk.CTkButton(btn_add_row, text="＋ Папку", command=self._add_folder, height=30, corner_radius=6).grid(row=0, column=1, padx=(3, 0), sticky="ew")

        btn_ctrl_row = ctk.CTkFrame(panel, fg_color="transparent")
        btn_ctrl_row.grid(row=2, column=0, padx=12, pady=(0, 10), sticky="ew")
        btn_ctrl_row.grid_columnconfigure((0, 1), weight=1)
        ctk.CTkButton(btn_ctrl_row, text="✕ Убрать", command=self._remove_file, height=30, corner_radius=6, fg_color="#8B0000", hover_color="#B22222").grid(row=0, column=0, padx=(0, 3), sticky="ew")
        ctk.CTkButton(btn_ctrl_row, text="🗑 Очистить всё", command=self._clear_list, height=30, corner_radius=6, fg_color=("gray60", "gray40"), hover_color=("gray50", "gray30")).grid(row=0, column=1, padx=(3, 0), sticky="ew")

        self._file_scroll = ctk.CTkScrollableFrame(panel, fg_color="transparent", scrollbar_button_color=("gray70", "gray35"))
        self._file_scroll.grid(row=3, column=0, padx=8, pady=(0, 8), sticky="nsew")
        self._file_scroll.grid_columnconfigure(0, weight=1)

        self._status_lbl = ctk.CTkLabel(panel, text="Нет загруженных файлов", font=ctk.CTkFont(size=11), text_color="gray")
        self._status_lbl.grid(row=4, column=0, padx=14, pady=(0, 12), sticky="w")

    def _build_right_panel(self):
        panel = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        panel.grid(row=0, column=1, sticky="nsew", padx=(1, 0))
        panel.grid_rowconfigure(3, weight=1)
        panel.grid_columnconfigure(0, weight=1)

        info_frame = ctk.CTkFrame(panel, fg_color=("gray88", "gray16"), corner_radius=0)
        info_frame.grid(row=0, column=0, sticky="ew")
        info_frame.grid_columnconfigure(0, weight=1)
        self._fname_lbl = ctk.CTkLabel(info_frame, text="Выберите файл из списка слева", font=ctk.CTkFont(size=15, weight="bold"), anchor="w")
        self._fname_lbl.grid(row=0, column=0, padx=18, pady=(14, 2), sticky="ew")
        self._finfo_lbl = ctk.CTkLabel(info_frame, text="", font=ctk.CTkFont(size=11), text_color="gray", anchor="w")
        self._finfo_lbl.grid(row=1, column=0, padx=18, pady=(0, 12), sticky="ew")

        ctrl_row = ctk.CTkFrame(panel, fg_color="transparent")
        ctrl_row.grid(row=1, column=0, padx=16, pady=(10, 4), sticky="ew")
        ctk.CTkButton(ctrl_row, text="☑  Выбрать все", command=self._select_all, width=148, height=30, corner_radius=8).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            ctrl_row, text="☐  Снять выбор", command=self._deselect_all, width=148, height=30, corner_radius=8,
            fg_color=("gray75", "gray30"), hover_color=("gray65", "gray40"), text_color=("black", "white"),
        ).pack(side="left")

        self._props_scroll = ctk.CTkScrollableFrame(
            panel, label_text="  Свойства выбранного документа", label_font=ctk.CTkFont(weight="bold", size=13), corner_radius=10
        )
        self._props_scroll.grid(row=3, column=0, padx=16, pady=(4, 8), sticky="nsew")
        self._props_scroll.grid_columnconfigure(0, weight=1)

        action_frame = ctk.CTkFrame(panel, fg_color=("gray85", "gray18"), corner_radius=10)
        action_frame.grid(row=4, column=0, padx=16, pady=(4, 16), sticky="ew")
        action_frame.grid_columnconfigure(0, weight=1)
        action_frame.grid_columnconfigure(1, weight=1)

        single_act = ctk.CTkFrame(action_frame, fg_color="transparent")
        single_act.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        ctk.CTkLabel(single_act, text="Для ТЕКУЩЕГО файла:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(0, 5))
        
        self._clear_sel_btn = ctk.CTkButton(single_act, text="🗑 Очистить выбранные свойства", command=self._clear_selected, state="disabled")
        self._clear_sel_btn.pack(anchor="w", pady=2, fill="x")
        
        self._clear_all_btn = ctk.CTkButton(single_act, text="💥 Очистить ВСЕ свойства", command=self._clear_all, state="disabled", fg_color="#E67E22", hover_color="#CF6D17")
        self._clear_all_btn.pack(anchor="w", pady=2, fill="x")

        bulk_act = ctk.CTkFrame(action_frame, fg_color="transparent")
        bulk_act.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        ctk.CTkLabel(bulk_act, text="Для ВСЕХ файлов в списке:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", pady=(0, 5))
        
        self._bulk_clear_btn = ctk.CTkButton(bulk_act, text="🔥 Очистить ВСЕ у ВСЕХ файлов", command=self._bulk_clear_all, fg_color="#8B0000", hover_color="#B22222")
        self._bulk_clear_btn.pack(anchor="w", pady=2, fill="x")

        rename_frame = ctk.CTkFrame(bulk_act, fg_color="transparent")
        rename_frame.pack(anchor="w", pady=(10, 0), fill="x")
        
        self._rename_entry = ctk.CTkEntry(rename_frame, placeholder_text="Текст для добавления к имени")
        self._rename_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self._bulk_rename_btn = ctk.CTkButton(rename_frame, text="✏️ Переименовать", command=self._bulk_rename, width=120)
        self._bulk_rename_btn.pack(side="right")

        self._progress_frame = ctk.CTkFrame(panel, fg_color="transparent")
        self._progress_frame.grid(row=5, column=0, padx=16, pady=(0, 10), sticky="ew")
        self._progress_frame.grid_columnconfigure(1, weight=1)
        self._progress_frame.grid_remove() 

        self._result_lbl = ctk.CTkLabel(self._progress_frame, text="Обработка...", font=ctk.CTkFont(weight="bold"))
        self._result_lbl.grid(row=0, column=0, padx=(0, 10))
        self._progress_bar = ctk.CTkProgressBar(self._progress_frame)
        self._progress_bar.grid(row=0, column=1, sticky="ew")

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Выберите документы Office",
            filetypes=[("Документы Office", "*.docx *.pptx *.xlsx"), ("Все файлы", "*.*")]
        )
        self._register_files(paths)

    def _add_folder(self):
        folder_path = filedialog.askdirectory(title="Выберите папку с документами")
        if not folder_path:
            return
            
        found_files = []
        for root, _, files in os.walk(folder_path):
            for file in files:
                if Path(file).suffix.lower() in SUPPORTED_EXT:
                    found_files.append(os.path.join(root, file))
                    
        if not found_files:
            messagebox.showinfo("Пусто", "В выбранной папке не найдено подходящих документов.")
            return
            
        self._register_files(found_files)

    def _register_files(self, paths):
        added = 0
        for p in paths:
            if p not in self._files and Path(p).suffix.lower() in SUPPORTED_EXT:
                self._files.append(p)
                added += 1
        
        if added:
            self._files.sort(key=lambda x: Path(x).name.lower())
            self._refresh_file_list()
            self._update_status()

    def _remove_file(self):
        if not self._current_file:
            return
        self._files.remove(self._current_file)
        self._current_file = None
        self._refresh_file_list()
        self._reset_right_panel()
        self._update_status()

    def _clear_list(self):
        if messagebox.askyesno("Подтверждение", "Убрать все файлы из списка?"):
            self._files.clear()
            self._current_file = None
            self._refresh_file_list()
            self._reset_right_panel()
            self._update_status()

    def _refresh_file_list(self):
        for w in self._file_scroll.winfo_children():
            w.destroy()
        self._file_buttons.clear()
        
        for i, fp in enumerate(self._files):
            ext = Path(fp).suffix.lower()
            active_c, hover_c = EXT_COLORS.get(ext, ("#455A64", "#546E7A"))
            is_sel = fp == self._current_file
            
            btn = ctk.CTkButton(
                self._file_scroll,
                text=f"  {Path(fp).name}",
                anchor="w",
                height=36,
                corner_radius=6,
                fg_color=active_c if is_sel else ("gray83", "gray22"),
                hover_color=hover_c,
                text_color="white" if is_sel else ("gray20", "gray90"),
                font=ctk.CTkFont(size=12),
                command=lambda p=fp: self._on_select(p),
            )
            btn.grid(row=i, column=0, sticky="ew", pady=2, padx=4)
            self._file_buttons[fp] = btn

        if not self._files:
            self._bulk_clear_btn.configure(state="disabled")
            self._bulk_rename_btn.configure(state="disabled")
        else:
            self._bulk_clear_btn.configure(state="normal")
            self._bulk_rename_btn.configure(state="normal")

    def _on_select(self, fp: str):
        self._current_file = fp
        self._refresh_file_list()
        self._load_props(fp)

    def _load_props(self, fp: str):
        self._reset_right_panel()
        ext = Path(fp).suffix.lower()
        size_kb = os.path.getsize(fp) // 1024
        self._fname_lbl.configure(text=Path(fp).name)
        self._finfo_lbl.configure(text=f"Тип: {ext.upper().lstrip('.')}   ·   Размер: {size_kb} КБ   ·   {fp}")
        try:
            props = get_file_properties(fp)
        except Exception as exc:
            messagebox.showerror("Ошибка чтения", f"Не удалось прочитать метаданные:\n{exc}")
            return
            
        self._prop_vars = {}
        self._build_props_header()
        
        for idx, (attr, info) in enumerate(props.items()):
            self._add_prop_row(idx, attr, info["label"], info["value"])
            
        self._clear_sel_btn.configure(state="normal")
        self._clear_all_btn.configure(state="normal")

    def _build_props_header(self):
        hdr = ctk.CTkFrame(self._props_scroll, fg_color=("gray78", "gray28"), corner_radius=6)
        hdr.grid(row=0, column=0, sticky="ew", padx=4, pady=(2, 6))
        hdr.grid_columnconfigure(2, weight=1)
        ctk.CTkLabel(hdr, text="", width=50).grid(row=0, column=0)
        ctk.CTkLabel(hdr, text="Свойство", font=ctk.CTkFont(weight="bold", size=12), width=220, anchor="w").grid(row=0, column=1, padx=8, pady=7, sticky="w")
        ctk.CTkLabel(hdr, text="Текущее значение", font=ctk.CTkFont(weight="bold", size=12), anchor="w").grid(row=0, column=2, padx=8, pady=7, sticky="w")

    def _add_prop_row(self, idx: int, attr: str, label: str, value: str):
        var = ctk.BooleanVar(value=False)
        self._prop_vars[attr] = var
        bg = ("gray90", "gray21") if idx % 2 == 0 else ("gray95", "gray17")
        row = ctk.CTkFrame(self._props_scroll, fg_color=bg, corner_radius=5)
        row.grid(row=idx + 1, column=0, sticky="ew", padx=4, pady=1)
        row.grid_columnconfigure(2, weight=1)
        ctk.CTkCheckBox(row, text="", variable=var, width=50, checkbox_width=20, checkbox_height=20).grid(row=0, column=0, padx=(12, 0), pady=7)
        ctk.CTkLabel(row, text=label, font=ctk.CTkFont(size=12), width=220, anchor="w").grid(row=0, column=1, padx=8, pady=7, sticky="w")
        
        is_empty = not value or value.lower() in ("none", "")
        display = "—" if is_empty else value
        val_color = ("gray60", "gray50") if is_empty else ("gray10", "gray95")
        ctk.CTkLabel(row, text=display, font=ctk.CTkFont(size=12), text_color=val_color, anchor="w", wraplength=460).grid(row=0, column=2, padx=8, pady=7, sticky="w")

    def _reset_right_panel(self):
        for w in self._props_scroll.winfo_children():
            w.destroy()
        self._prop_vars = {}
        self._fname_lbl.configure(text="Выберите файл из списка слева")
        self._finfo_lbl.configure(text="")
        self._clear_sel_btn.configure(state="disabled")
        self._clear_all_btn.configure(state="disabled")
        self._progress_frame.grid_remove()

    def _select_all(self):
        for v in self._prop_vars.values(): v.set(True)

    def _deselect_all(self):
        for v in self._prop_vars.values(): v.set(False)

    def _clear_selected(self):
        if not self._current_file: return
        selected = [a for a, v in self._prop_vars.items() if v.get()]
        if not selected:
            messagebox.showwarning("Нет выбора", "Отметьте хотя бы одно свойство.")
            return
        self._do_clear([self._current_file], selected)

    def _clear_all(self):
        if not self._current_file: return
        self._do_clear([self._current_file], list(PROPS_MAP.keys()))

    def _bulk_clear_all(self):
        if not self._files: return
        if not messagebox.askyesno("Массовая очистка", f"ВНИМАНИЕ!\n\nВы собираетесь очистить ВСЕ метаданные у ВСЕХ файлов в списке ({len(self._files)} шт.).\n\nПродолжить?"):
            return
        
        attrs_to_clear = list(PROPS_MAP.keys())
        self._progress_frame.grid()
        self._result_lbl.configure(text="Обработка...", text_color=("black", "white"))
        self._progress_bar.set(0)
        
        success = 0
        total = len(self._files)
        
        for i, fp in enumerate(self._files):
            try:
                clear_file_properties(fp, attrs_to_clear)
                success += 1
            except Exception as e:
                print(f"Ошибка с файлом {fp}: {e}")
                
            self._progress_bar.set((i + 1) / total)
            self.update() 
            
        self._result_lbl.configure(text=f"✓ Готово! Очищено файлов: {success} из {total}", text_color="#2ECC71")
        if self._current_file:
            self._load_props(self._current_file) 

    def _bulk_rename(self):
        if not self._files: return
        
        append_text = self._rename_entry.get()
        if not append_text:
            messagebox.showwarning("Внимание", "Введите текст для добавления к имени.")
            return
            
        if not messagebox.askyesno("Переименование", f"Добавить текст '{append_text}' к именам всех загруженных файлов?"):
            return
            
        self._progress_frame.grid()
        self._result_lbl.configure(text="Переименование...", text_color=("black", "white"))
        self._progress_bar.set(0)
        
        new_file_list = []
        success = 0
        total = len(self._files)
        
        for i, fp in enumerate(self._files):
            try:
                p = Path(fp)
                new_name = p.stem + append_text + p.suffix
                new_path = p.with_name(new_name)
                
                os.rename(fp, new_path)
                new_file_list.append(str(new_path))
                
                if self._current_file == fp:
                    self._current_file = str(new_path)
                    
                success += 1
            except Exception as e:
                print(f"Ошибка переименования {fp}: {e}")
                new_file_list.append(fp)
                
            self._progress_bar.set((i + 1) / total)
            self.update()
            
        self._files = new_file_list
        self._files.sort(key=lambda x: Path(x).name.lower())
        self._refresh_file_list()
        
        if self._current_file:
            self._load_props(self._current_file)
            
        self._result_lbl.configure(text=f"✓ Переименовано файлов: {success} из {total}", text_color="#2ECC71")

    def _do_clear(self, file_paths: list, attrs: list):
        count = 0
        for fp in file_paths:
            try:
                count += clear_file_properties(fp, attrs)
            except Exception as exc:
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл {Path(fp).name}:\n{exc}")
                return
                
        self._progress_frame.grid()
        self._progress_bar.set(1.0)
        self._result_lbl.configure(text=f"✓ Успешно очищено свойств: {count}", text_color="#2ECC71")
        if self._current_file in file_paths:
            self._load_props(self._current_file)

    def _update_status(self):
        n = len(self._files)
        self._status_lbl.configure(text=f"Загружено файлов: {n}" if n else "Нет загруженных файлов")

if __name__ == "__main__":
    app = MetaCleanerApp()
    app.mainloop()