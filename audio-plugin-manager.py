import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import ctypes
import ctypes.wintypes
import platform
import subprocess
import webbrowser

PLUGIN_DB = 'plugins.json'
PLUGIN_EXTENSIONS = ['.dll', '.vst3', '.component']

DEFAULT_FOLDERS = [
    os.path.expandvars(r'%ProgramFiles%\Steinberg\VstPlugins'),
    os.path.expandvars(r'%ProgramFiles(x86)%\Steinberg\VstPlugins'),
    os.path.expandvars(r'%ProgramFiles%\VSTPlugins'),
    os.path.expandvars(r'%ProgramFiles(x86)%\VSTPlugins'),
    '/Library/Audio/Plug-Ins/VST',
    '/Library/Audio/Plug-Ins/VST3',
    '/Library/Audio/Plug-Ins/Components',
]

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

    if not ctypes.windll.version.VerQueryValueW(res, r"\VarFileInfo\Translation", ctypes.byref(lplpBuffer), ctypes.byref(puLen)):
        return None

    lang_codepage = ctypes.cast(lplpBuffer.value, ctypes.POINTER(ctypes.c_ubyte * puLen.value)).contents
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
                if ext in PLUGIN_EXTENSIONS:
                    path = os.path.join(root, file)
                    name = os.path.splitext(file)[0]
                    if ext == '.dll':
                        fmt = 'VST2'
                        version_info = get_file_version_info(path)
                        if version_info and version_info["ProductName"]:
                            name = version_info["ProductName"]
                        company = version_info["CompanyName"] if version_info else None
                    elif ext == '.vst3':
                        fmt = 'VST3'
                        company = None
                    elif ext == '.component':
                        fmt = 'AU'
                        company = None
                    else:
                        fmt = 'Unknown'
                        company = None
                    bitness = '64-bit' if 'x64' in path.lower() or '64' in path else '32-bit'
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
        return json.load(f)

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

        about_btn = ttk.Button(btn_frame, text="About", command=lambda: webbrowser.open("https://github.com/plugindeals/audio-plugin-manager"))
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
        tree_frame.pack(fill='both', expand=True, padx=10, pady=10)

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
        scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind('<Double-1>', self.edit_notes)
        self.tree.bind('<Button-3>', self.show_context_menu)

        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Toggle Favorite", command=self.toggle_favorite)
        self.context_menu.add_command(label="Open File Location", command=self.open_file_location)

        # Footer status bar
        self.status_bar = ttk.Label(self, text="Total plugins: 0", anchor='w')
        self.status_bar.pack(side='bottom', fill='x', padx=10, pady=2)

    def update_list(self):
        self.tree.delete(*self.tree.get_children())
        count = 0
        for plugin in self.plugins:
            if self.search_var.get().lower() not in plugin['name'].lower():
                continue
            if self.format_var.get() != 'All' and plugin['format'] != self.format_var.get():
                continue
            if self.bitness_var.get() != 'All' and plugin['bitness'] != self.bitness_var.get():
                continue
            if self.fav_only_var.get() and not plugin['favorite']:
                continue
            self.tree.insert('', 'end', values=(
                plugin['name'], plugin['vendor'], plugin['format'], plugin['bitness'],
                'Yes' if plugin['favorite'] else 'No', plugin['notes'], plugin['path']
            ))
            count += 1
        self.status_bar.config(text=f"Total plugins: {count}")

    def scan_plugins_default(self):
        new_plugins = scan_plugins(DEFAULT_FOLDERS)
        if not new_plugins:
            messagebox.showinfo("No Plugins Found", "No plugins found in default folders.")
            return
        existing_paths = {p['path']: p for p in self.plugins}
        for np in new_plugins:
            if np['path'] in existing_paths:
                np['notes'] = existing_paths[np['path']].get('notes', '')
                np['favorite'] = existing_paths[np['path']].get('favorite', False)
        merged = {p['path']: p for p in (self.plugins + new_plugins)}
        self.plugins = list(merged.values())
        save_plugins(self.plugins)
        self.update_list()
        messagebox.showinfo("Scan Complete", f"Found and added {len(new_plugins)} plugins.")

    def add_directory_manual(self):
        folder = filedialog.askdirectory(title="Select Plugin Folder to Add")
        if not folder:
            return
        plugins_in_folder = scan_plugins([folder])
        if not plugins_in_folder:
            messagebox.showinfo("No Plugins Found", "No plugin files found in the selected directory.")
            return
        existing_paths = {p['path'] for p in self.plugins}
        new_plugins = [p for p in plugins_in_folder if p['path'] not in existing_paths]
        self.plugins.extend(new_plugins)
        save_plugins(self.plugins)
        self.update_list()
        messagebox.showinfo("Directory Added", f"Added {len(new_plugins)} new plugins.")

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[('CSV files', '*.csv')])
        if not path:
            return
        with open(path, 'w', encoding='utf-8') as f:
            f.write("Name,Vendor,Format,Bitness,Favorite,Notes,Path\n")
            for plugin in self.plugins:
                f.write(f'"{plugin["name"]}","{plugin["vendor"]}","{plugin["format"]}","{plugin["bitness"]}",'
                        f'"{"Yes" if plugin["favorite"] else "No"}","{plugin["notes"]}","{plugin["path"]}"\n')
        messagebox.showinfo("Exported", f"Plugin list exported to {path}")

    def show_context_menu(self, event):
        selected = self.tree.identify_row(event.y)
        if selected:
            self.tree.selection_set(selected)
            self.context_menu.tk_popup(event.x_root, event.y_root)

    def toggle_favorite(self):
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        path = item['values'][6]
        for plugin in self.plugins:
            if plugin['path'] == path:
                plugin['favorite'] = not plugin['favorite']
                break
        save_plugins(self.plugins)
        self.update_list()

    def edit_notes(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        path = item['values'][6]
        current_notes = item['values'][5]
        new_notes = simpledialog.askstring("Edit Notes", "Enter notes for this plugin:", initialvalue=current_notes)
        if new_notes is not None:
            for plugin in self.plugins:
                if plugin['path'] == path:
                    plugin['notes'] = new_notes
                    break
            save_plugins(self.plugins)
            self.update_list()

    def open_file_location(self):
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        plugin_path = item['values'][6]
        if not os.path.exists(plugin_path):
            messagebox.showwarning("Not Found", f"File not found: {plugin_path}")
            return
        folder = os.path.dirname(plugin_path)
        if platform.system() == 'Windows':
            subprocess.run(['explorer', '/select,', plugin_path])
        elif platform.system() == 'Darwin':
            subprocess.run(['open', '--reveal', plugin_path])
        else:
            subprocess.run(['xdg-open', folder])

if __name__ == "__main__":
    app = PluginManagerApp()
    app.mainloop()
