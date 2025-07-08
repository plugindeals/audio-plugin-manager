import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import ctypes
import ctypes.wintypes
import platform
import subprocess
import webbrowser
import csv

PLUGIN_DB = 'plugins.json'
PLUGIN_EXTENSIONS = ['.dll', '.vst3', '.component']

DEFAULT_FOLDERS = [
    os.path.expandvars(r'%ProgramFiles%\Steinberg\VstPlugins'),
    os.path.expandvars(r'%ProgramFiles%\Cakewalk\VstPlugins'),
    os.path.expandvars(r'%ProgramFiles(x86)%\Steinberg\VstPlugins'),
    os.path.expandvars(r'%ProgramFiles%\VSTPlugins'),
    os.path.expandvars(r'%ProgramFiles%\Native Instruments'),
    os.path.expandvars(r'%ProgramFiles(x86)%\VSTPlugins'),
    os.path.expandvars(r'%ProgramFiles%\Common Files\VST3'),
    os.path.expandvars(r'%ProgramFiles%\Common Files\VST2'),
    os.path.expandvars(r'%ProgramFiles(x86)%\Common Files\VST3'),
    '/Library/Audio/Plug-Ins/VST',
    '/Library/Audio/Plug-Ins/VST3',
    '/Library/Audio/Plug-Ins/Components',
]

EXCLUDED_DLLS = {
    "webview2loader.dll",
    "microsoft.web.webview2.core.dll",
}

def get_bitness(file_path):
    try:
        with open(file_path, 'rb') as f:
            dos_headers = f.read(64)
            if dos_headers[0:2] != b'MZ':
                return 'Unknown'

            f.seek(int.from_bytes(dos_headers[60:64], byteorder='little'))  # PE header offset
            pe_header = f.read(6)
            if pe_header[0:2] != b'PE':
                return 'Unknown'

            machine = int.from_bytes(pe_header[4:6], byteorder='little')
            if machine == 0x8664:
                return '64-bit'
            elif machine == 0x014c:
                return '32-bit'
            else:
                return 'Unknown'
    except Exception:
        return 'Unknown'


def get_file_version_info(filename):
    if not os.path.exists(filename):
        return None
    size = ctypes.windll.version.GetFileVersionInfoSizeW(filename, None)
    if size == 0:
        return None

    res = ctypes.create_string_buffer(size)
    success = ctypes.windll.version.GetFileVersionInfoW(filename, 0, size, res)
    if not success:
        return None

    lplpBuffer = ctypes.c_void_p()
    puLen = ctypes.wintypes.UINT()

    if not ctypes.windll.version.VerQueryValueW(res, r"\\VarFileInfo\\Translation", ctypes.byref(lplpBuffer), ctypes.byref(puLen)):
        return None

    lang_codepage = ctypes.cast(lplpBuffer.value, ctypes.POINTER(ctypes.c_ubyte * puLen.value)).contents
    if puLen.value < 4:
        return None
    lang = (lang_codepage[0] + (lang_codepage[1] << 8))
    codepage = (lang_codepage[2] + (lang_codepage[3] << 8))

    def query_value(name):
        sub_block = f"\\StringFileInfo\\{lang:04x}{codepage:04x}\\{name}"
        if ctypes.windll.version.VerQueryValueW(res, sub_block, ctypes.byref(lplpBuffer), ctypes.byref(puLen)):
            return ctypes.wstring_at(lplpBuffer.value, puLen.value)
        return None

    return {
        "ProductName": query_value("ProductName"),
        "CompanyName": query_value("CompanyName"),
        "FileDescription": query_value("FileDescription"),
    }

def scan_plugins(folders):
    plugins = []
    for folder in folders:
        if not os.path.exists(folder):
            continue
        for root, _, files in os.walk(folder):
            for file in files:
                ext = os.path.splitext(file)[1].lower()
                filename_lower = file.lower()
                # Exclude specific DLLs here
                if ext == '.dll' and filename_lower in EXCLUDED_DLLS:
                    continue
                if ext in PLUGIN_EXTENSIONS:
                    path = os.path.join(root, file)
                    name = os.path.splitext(file)[0]
                    company = None
                    if ext == '.dll':
                        fmt = 'VST2'
                        version_info = get_file_version_info(path)
                        if version_info and version_info["ProductName"]:
                            name = version_info["ProductName"]
                        company = version_info["CompanyName"] if version_info else None
                    elif ext == '.vst3':
                        fmt = 'VST3'
                    elif ext == '.component':
                        fmt = 'AU'
                    else:
                        fmt = 'Unknown'

                    bitness = get_bitness(path) if ext == '.dll' else '64-bit'

                    plugins.append({
                        'name': name,
                        'format': fmt,
                        'bitness': bitness,
                        'path': path,
                        'vendor': company or '',
                        'notes': '',
                        'favorite': False
                    })
    return plugins

def load_plugins():
    if not os.path.exists(PLUGIN_DB):
        return []
    with open(PLUGIN_DB, 'r', encoding='utf-8') as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []

def save_plugins(plugins):
    with open(PLUGIN_DB, 'w', encoding='utf-8') as f:
        json.dump(plugins, f, indent=2)

class PluginManagerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Audio Plugin Manager")
        self.geometry("900x500")
        self.plugins = load_plugins()
        self.create_widgets()
        self.update_list()
        self.show_disclaimer()

    def show_disclaimer(self):
        try:
            import clr
            raise ImportError  # fallback
        except ImportError:
            messagebox.showinfo("Disclaimer",
                "This program is free software licensed under the GNU General Public License v3.0.\n"
                "You may redistribute and/or modify it under the terms of the GPL-3.0.\n\n"
                "This software is provided 'as is', without any warranty of any kind.\n"
                "The author is not liable for any damages arising from its use.\n\n"
                "Source code available at:\nhttps://github.com/plugindeals")

    def create_widgets(self):
        btn_frame = ttk.Frame(self)
        btn_frame.pack(side='top', fill='x', padx=10, pady=5)

        scan_btn = ttk.Button(btn_frame, text="Scan Plugins", command=self.scan_plugins_default)
        scan_btn.pack(side='left', padx=5)

        add_dir_btn = ttk.Button(btn_frame, text="Add Directory", command=self.add_directory_manual)
        add_dir_btn.pack(side='left', padx=5)

        export_btn = ttk.Button(btn_frame, text="Export to CSV", command=self.export_csv)
        export_btn.pack(side='left', padx=5)

        about_btn = ttk.Button(btn_frame, text="About", command=self.open_about)
        about_btn.pack(side='left', padx=5)

        ttk.Label(btn_frame, text="Search:").pack(side='left', padx=(20,5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', lambda *args: self.update_list())
        search_entry = ttk.Entry(btn_frame, textvariable=self.search_var, width=25)
        search_entry.pack(side='left')

        ttk.Label(btn_frame, text="Format:").pack(side='left', padx=(15,5))
        self.format_var = tk.StringVar(value="All")
        format_menu = ttk.OptionMenu(btn_frame, self.format_var, 'All', 'All', 'VST2', 'VST3', 'AU', command=lambda e: self.update_list())
        format_menu.pack(side='left')

        ttk.Label(btn_frame, text="Bitness:").pack(side='left', padx=(15,5))
        self.bitness_var = tk.StringVar(value="All")
        bitness_menu = ttk.OptionMenu(btn_frame, self.bitness_var, 'All', 'All', '32-bit', '64-bit', command=lambda e: self.update_list())
        bitness_menu.pack(side='left')

        self.fav_only_var = tk.BooleanVar()
        fav_check = ttk.Checkbutton(btn_frame, text="Favorites Only", variable=self.fav_only_var, command=self.update_list)
        fav_check.pack(side='left', padx=10)

        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill='both', expand=True, padx=10, pady=(10,5))

        columns = ('Name', 'Vendor', 'Format', 'Bitness', 'Favorite', 'Notes', 'Path')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, stretch=True)
        self.tree.column('Name', width=150)
        self.tree.column('Vendor', width=120)
        self.tree.column('Format', width=70, anchor='center')
        self.tree.column('Bitness', width=70, anchor='center')
        self.tree.column('Favorite', width=70, anchor='center')
        self.tree.column('Notes', width=200)
        self.tree.column('Path', width=350)
        self.tree.pack(side='left', fill='both', expand=True)

        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side='right', fill='y')

        self.tree.bind('<Double-1>', self.edit_plugin)
        self.tree.bind('<Button-3>', self.show_context_menu)

        # Label to show total plugins count
        self.count_label = ttk.Label(self, text="")
        self.count_label.pack(side='bottom', anchor='w', padx=10, pady=5)

    def update_list(self):
        search = self.search_var.get().lower()
        fmt = self.format_var.get()
        bitness = self.bitness_var.get()
        fav_only = self.fav_only_var.get()

        self.tree.delete(*self.tree.get_children())
        count = 0

        for idx, plugin in enumerate(self.plugins):
            if fav_only and not plugin.get('favorite', False):
                continue
            if fmt != 'All' and plugin.get('format', '') != fmt:
                continue
            if bitness != 'All' and plugin.get('bitness', '') != bitness:
                continue
            if search:
                if not (search in plugin.get('name', '').lower() or search in plugin.get('vendor', '').lower() or search in plugin.get('notes', '').lower()):
                    continue
            fav_text = 'Yes' if plugin.get('favorite', False) else ''
            self.tree.insert('', 'end', iid=str(idx), values=(
                plugin.get('name', ''),
                plugin.get('vendor', ''),
                plugin.get('format', ''),
                plugin.get('bitness', ''),
                fav_text,
                plugin.get('notes', ''),
                plugin.get('path', '')
            ))
            count += 1

        self.count_label.config(text=f"Total Plugins: {count}")

    def scan_plugins_default(self):
        plugins_found = scan_plugins(DEFAULT_FOLDERS)
        existing_paths = {p['path'] for p in self.plugins}
        added = 0
        for p in plugins_found:
            if p['path'] not in existing_paths:
                self.plugins.append(p)
                added += 1
        save_plugins(self.plugins)
        self.update_list()
        messagebox.showinfo("Scan Complete", f"Scan completed. {added} new plugins added.")

    def add_directory_manual(self):
        folder = filedialog.askdirectory(title="Select Folder to Scan")
        if folder:
            plugins_found = scan_plugins([folder])
            existing_paths = {p['path'] for p in self.plugins}
            added = 0
            for p in plugins_found:
                if p['path'] not in existing_paths:
                    self.plugins.append(p)
                    added += 1
            save_plugins(self.plugins)
            self.update_list()
            messagebox.showinfo("Scan Complete", f"Scan completed. {added} new plugins added.")

    def edit_plugin(self, event):
        item_id = self.tree.focus()
        if not item_id:
            return
        plugin = self.plugins[int(item_id)]

        edit_dialog = PluginEditDialog(self, plugin)
        self.wait_window(edit_dialog)

        if edit_dialog.updated:
            save_plugins(self.plugins)
            self.update_list()

    def show_context_menu(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        self.tree.selection_set(item_id)
        plugin = self.plugins[int(item_id)]

        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Open Folder", command=lambda: self.open_folder(plugin['path']))
        menu.add_command(label="Edit Plugin Info", command=lambda: self.edit_plugin_manual(int(item_id)))
        menu.add_command(label="Mark as Favorite" if not plugin.get('favorite') else "Unmark Favorite",
                         command=lambda: self.toggle_favorite(int(item_id)))
        menu.post(event.x_root, event.y_root)

    def open_folder(self, path):
        folder = os.path.dirname(path)
        if platform.system() == "Windows":
            os.startfile(folder)
        elif platform.system() == "Darwin":
            subprocess.run(["open", folder])
        else:
            subprocess.run(["xdg-open", folder])

    def open_file(self, path):
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])

    def edit_plugin_manual(self, idx):
        plugin = self.plugins[idx]
        edit_dialog = PluginEditDialog(self, plugin)
        self.wait_window(edit_dialog)
        if edit_dialog.updated:
            save_plugins(self.plugins)
            self.update_list()

    def toggle_favorite(self, idx):
        plugin = self.plugins[idx]
        plugin['favorite'] = not plugin.get('favorite', False)
        save_plugins(self.plugins)
        self.update_list()

    def export_csv(self):
        filepath = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save CSV"
        )
        if not filepath:
            return

        with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Name', 'Vendor', 'Format', 'Bitness', 'Favorite', 'Notes', 'Path'])
            for p in self.plugins:
                writer.writerow([
                    p.get('name', ''),
                    p.get('vendor', ''),
                    p.get('format', ''),
                    p.get('bitness', ''),
                    'Yes' if p.get('favorite', False) else 'No',
                    p.get('notes', ''),
                    p.get('path', '')
                ])
        messagebox.showinfo("Export Complete", f"Plugin list exported to {filepath}")

    def open_about(self):
        webbrowser.open("https://github.com/plugindeals/audio-plugin-manager")

class PluginEditDialog(tk.Toplevel):
    def __init__(self, parent, plugin):
        super().__init__(parent)
        self.title("Edit Plugin")
        self.plugin = plugin
        self.updated = False

        self.create_widgets()
        self.grab_set()
        self.focus_force()

    def create_widgets(self):
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill='both', expand=True)

        ttk.Label(frm, text="Name:").grid(row=0, column=0, sticky='e')
        self.name_var = tk.StringVar(value=self.plugin.get('name', ''))
        ttk.Entry(frm, textvariable=self.name_var, width=40).grid(row=0, column=1, sticky='w')

        ttk.Label(frm, text="Vendor:").grid(row=1, column=0, sticky='e')
        self.vendor_var = tk.StringVar(value=self.plugin.get('vendor', ''))
        ttk.Entry(frm, textvariable=self.vendor_var, width=40).grid(row=1, column=1, sticky='w')

        ttk.Label(frm, text="Format:").grid(row=2, column=0, sticky='e')
        self.format_var = tk.StringVar(value=self.plugin.get('format', ''))
        ttk.Entry(frm, textvariable=self.format_var, width=40).grid(row=2, column=1, sticky='w')

        ttk.Label(frm, text="Bitness:").grid(row=3, column=0, sticky='e')
        self.bitness_var = tk.StringVar(value=self.plugin.get('bitness', ''))
        ttk.Entry(frm, textvariable=self.bitness_var, width=40).grid(row=3, column=1, sticky='w')

        ttk.Label(frm, text="Notes:").grid(row=4, column=0, sticky='ne')
        self.notes_text = tk.Text(frm, width=40, height=5)
        self.notes_text.grid(row=4, column=1, sticky='w')
        self.notes_text.insert('1.0', self.plugin.get('notes', ''))

        self.favorite_var = tk.BooleanVar(value=self.plugin.get('favorite', False))
        ttk.Checkbutton(frm, text="Favorite", variable=self.favorite_var).grid(row=5, column=1, sticky='w')

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=10)

        ttk.Button(btn_frame, text="Save", command=self.save).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side='left', padx=5)

    def save(self):
        self.plugin['name'] = self.name_var.get()
        self.plugin['vendor'] = self.vendor_var.get()
        self.plugin['format'] = self.format_var.get()
        self.plugin['bitness'] = self.bitness_var.get()
        self.plugin['notes'] = self.notes_text.get('1.0', 'end').strip()
        self.plugin['favorite'] = self.favorite_var.get()
        self.updated = True
        self.destroy()

    def cancel(self):
        self.destroy()

if __name__ == "__main__":
    app = PluginManagerApp()
    app.mainloop()
