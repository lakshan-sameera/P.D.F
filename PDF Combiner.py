import customtkinter as ctk
from tkinter import filedialog, messagebox
from pypdf import PdfReader, PdfWriter
import os
import configparser
import sys
import subprocess
import json
import datetime

HISTORY_FILE = 'combined_history.json'

# A custom CTkInputDialog that can be given a parent
class CustomInputDialog(ctk.CTkInputDialog):
    def __init__(self, *args, **kwargs):
        parent = kwargs.pop("parent", None)
        super().__init__(*args, **kwargs)
        if parent:
            # Position the dialog in the center of the parent window
            self.geometry(f"+{parent.winfo_x()+parent.winfo_width()//2-self.winfo_width()//2}+{parent.winfo_y()+parent.winfo_height()//2-self.winfo_height()//2}")

class RotationDialog(ctk.CTkToplevel):
    def __init__(self, parent, file_name, max_pages):
        super().__init__(parent)
        self.transient(parent)
        self.title("Rotate Pages")
        self.geometry("350x200")
        
        self.result = None

        ctk.CTkLabel(self, text=f"Rotate: {file_name}").pack(pady=5)

        ctk.CTkLabel(self, text="Rotation Angle:").pack()
        self.angle_var = ctk.StringVar(value="90")
        ctk.CTkOptionMenu(self, variable=self.angle_var, values=["90", "180", "270"]).pack()

        ctk.CTkLabel(self, text=f"Pages (e.g., '1-3, 5', or 'all'):").pack()
        self.pages_var = ctk.StringVar(value="all")
        ctk.CTkEntry(self, textvariable=self.pages_var).pack(fill="x", padx=10)

        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=10)
        ctk.CTkButton(button_frame, text="Apply", command=self.apply).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Cancel", command=self.cancel).pack(side="left", padx=5)
        
        self.grab_set()
        self.wait_window()

    def apply(self):
        self.result = {
            "angle": int(self.angle_var.get()),
            "pages_str": self.pages_var.get()
        }
        self.destroy()

    def cancel(self):
        self.destroy()

class ScrollableFileList(ctk.CTkScrollableFrame):
    def __init__(self, master, app_instance, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app_instance
        self.labels = []
        
        self.placeholder_label = ctk.CTkLabel(self, text="Add PDFs using the button above", text_color="gray50")
        self.placeholder_label.pack(pady=20)

    def update_list(self):
        for label in self.labels:
            label.destroy()
        self.labels = []
        
        if not self.app.file_list:
            self.placeholder_label.pack(pady=20)
        else:
            self.placeholder_label.pack_forget()

        for i, item in enumerate(self.app.file_list):
            display_text = os.path.basename(item['path'])
            if item.get('pages'):
                display_text += f" (Pages: {item['pages']})"
            if item.get('rotation'):
                display_text += f" (Rotated)"
            
            label_frame = ctk.CTkFrame(self, corner_radius=6)
            if i == self.app.selected_index:
                label_frame.configure(fg_color=("gray75", "gray25"), border_width=2, border_color=("gray60", "gray40"))
            else:
                label_frame.configure(fg_color="transparent")

            label_frame.pack(fill="x", padx=5, pady=(2,3), ipady=5)

            label = ctk.CTkLabel(label_frame, text=display_text, fg_color="transparent", anchor="w")
            label.pack(side="left", fill="x", expand=True, padx=10)

            label.bind("<Button-1>", lambda e, index=i: self.app.select_file(index))
            label_frame.bind("<Button-1>", lambda e, index=i: self.app.select_file(index))

            self.labels.append(label_frame)

class PDFCombinerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Combiner")
        self.root.geometry("600x750")

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.file_list = []
        self.selected_index = -1
        self.config_file = 'config.ini'
        self.last_directory = self.load_last_directory()
        self.last_removed_item = None

        self.main_frame = ctk.CTkFrame(root)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.top_frame = ctk.CTkFrame(self.main_frame)
        self.top_frame.pack(fill="x", pady=(0, 5))

        self.add_button = ctk.CTkButton(self.top_frame, text="Add PDFs", command=self.add_pdfs)
        self.add_button.pack(side="left", padx=5, pady=5)
        
        self.clear_all_button = ctk.CTkButton(self.top_frame, text="Clear All", command=self.reset)
        self.clear_all_button.pack(side="left", padx=5, pady=5)
        
        self.theme_modes = ["system", "dark", "light"]
        self.theme_index = self.theme_modes.index(ctk.get_appearance_mode() if ctk.get_appearance_mode() in self.theme_modes else "system")
        self.theme_button = ctk.CTkButton(self.top_frame, text=f"Theme: {self.theme_modes[self.theme_index].capitalize()} (Click to change)", command=self.toggle_theme)
        self.theme_button.pack(side="right", padx=5, pady=5)

        self.file_list_frame = ScrollableFileList(self.main_frame, self, label_text="Files to Combine")
        self.file_list_frame.pack(fill="both", expand=True, pady=5)

        self.list_mgmt_frame = ctk.CTkFrame(self.main_frame)
        self.list_mgmt_frame.pack(fill="x", pady=5)

        self.up_button = ctk.CTkButton(self.list_mgmt_frame, text="Move Up", command=self.move_up)
        self.up_button.pack(side="left", padx=5, pady=5)

        self.down_button = ctk.CTkButton(self.list_mgmt_frame, text="Move Down", command=self.move_down)
        self.down_button.pack(side="left", padx=5, pady=5)
        
        self.remove_button = ctk.CTkButton(self.list_mgmt_frame, text="Remove", command=self.remove_selected)
        self.remove_button.pack(side="left", padx=5, pady=5)

        self.undo_button = ctk.CTkButton(self.list_mgmt_frame, text="Undo Remove", command=self.undo_remove, state="disabled")
        self.undo_button.pack(side="left", padx=5, pady=5)

        self.page_range_button = ctk.CTkButton(self.list_mgmt_frame, text="Set Page Range", command=self.set_page_range)
        self.page_range_button.pack(side="left", padx=5, pady=5)

        self.rotate_button = ctk.CTkButton(self.list_mgmt_frame, text="Rotate Pages", command=self.rotate_pages)
        self.rotate_button.pack(side="left", padx=5, pady=5)
        
        self.tab_view = ctk.CTkTabview(self.main_frame, height=230)
        self.tab_view.pack(fill="x", pady=5)
        self.tab_view.add("Metadata")
        self.tab_view.add("Options")
        self.tab_view.add("History")
        self.tab_view.tab("Metadata").grid_columnconfigure(1, weight=1)
        self.tab_view.tab("Options").grid_columnconfigure(1, weight=1)
        self.tab_view.tab("History").grid_columnconfigure(0, weight=1)

        self.meta_frame = self.tab_view.tab("Metadata")
        self.options_frame = self.tab_view.tab("Options")
        self.history_frame = self.tab_view.tab("History")
        
        self.title_var = ctk.StringVar()
        self.author_var = ctk.StringVar()
        self.subject_var = ctk.StringVar()
        self.creator_var = ctk.StringVar()
        self.producer_var = ctk.StringVar()
        self.keywords_var = ctk.StringVar()
        self.creation_date_var = ctk.StringVar()
        self.mod_date_var = ctk.StringVar()

        meta_fields = [
            ("Title:", self.title_var), ("Author:", self.author_var),
            ("Subject:", self.subject_var), ("Creator:", self.creator_var),
            ("Producer:", self.producer_var), ("Keywords:", self.keywords_var),
            ("Creation Date (YYYYMMDDHHmmSS):", self.creation_date_var),
            ("Modification Date (YYYYMMDDHHmmSS):", self.mod_date_var)
        ]

        for i, (text, var) in enumerate(meta_fields):
            ctk.CTkLabel(self.meta_frame, text=text).grid(row=i, column=0, sticky="e", padx=5, pady=2)
            ctk.CTkEntry(self.meta_frame, textvariable=var).grid(row=i, column=1, sticky="ew", padx=5, pady=2)
        
        self.password_var = ctk.StringVar()
        ctk.CTkLabel(self.options_frame, text="Password (optional):").grid(row=0, column=0, sticky="e", padx=5, pady=2)
        ctk.CTkEntry(self.options_frame, textvariable=self.password_var, show="*").grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        
        self.auto_open_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(self.options_frame, text="Open file after saving", variable=self.auto_open_var).grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        self.combine_button = ctk.CTkButton(self.main_frame, text="Combine PDFs", command=self.combine_pdfs, height=30)
        self.combine_button.pack(fill="x", pady=5)

        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.set(0)

        self.status_var = ctk.StringVar(value="Ready")
        self.status_bar = ctk.CTkLabel(self.main_frame, textvariable=self.status_var, anchor='w')
        self.status_bar.pack(fill='x', pady=(5,0))

        # History UI
        self.history_frame.grid_rowconfigure(0, weight=1)
        self.history_frame.grid_columnconfigure(0, weight=1)
        self.history_listbox = ctk.CTkScrollableFrame(self.history_frame)
        self.history_listbox.grid(row=0, column=0, sticky="nsew", padx=5, pady=5, columnspan=2)
        self.history_labels = []

        self.history_detail = ctk.CTkTextbox(self.history_frame, state="disabled")
        self.history_detail.grid(row=1, column=0, sticky="ew", padx=5, pady=5, columnspan=2)
        self.history_frame.grid_rowconfigure(1, weight=1)

        self.clear_history_button = ctk.CTkButton(self.history_frame, text="Clear All History", command=self.clear_history)
        self.clear_history_button.grid(row=2, column=0, columnspan=2, pady=5)

        self.load_history()
        self.refresh_history_ui()

    def _disable_undo(self):
        self.last_removed_item = None
        self.undo_button.configure(state="disabled")

    def undo_remove(self):
        if self.last_removed_item:
            item = self.last_removed_item['item']
            index = self.last_removed_item['index']
            self.file_list.insert(index, item)
            self.update_status(f"Restored: {os.path.basename(item['path'])}")
            self.file_list_frame.update_list()
            self._disable_undo()

    def select_file(self, index):
        if self.selected_index == index:
            self.selected_index = -1
            self.clear_metadata_fields()
            self.update_status("File deselected.")
        else:
            self.selected_index = index
            self.preview_metadata()
        self.file_list_frame.update_list()
    
    def toggle_theme(self):
        self.theme_index = (self.theme_index + 1) % len(self.theme_modes)
        new_mode = self.theme_modes[self.theme_index]
        ctk.set_appearance_mode(new_mode)
        self.theme_button.configure(text=f"Theme: {new_mode.capitalize()} (Click to change)")
        self.update_status(f"Theme changed to {new_mode} mode.")

    def load_last_directory(self):
        config = configparser.ConfigParser()
        if os.path.exists(self.config_file):
            config.read(self.config_file)
            return config.get('Settings', 'LastDirectory', fallback=os.path.expanduser('~'))
        return os.path.expanduser('~')

    def save_last_directory(self, directory):
        config = configparser.ConfigParser()
        config['Settings'] = {'LastDirectory': directory}
        with open(self.config_file, 'w') as configfile:
            config.write(configfile)
        self.last_directory = directory

    def update_status(self, text):
        self.status_var.set(text)

    def open_file(self, path):
        try:
            if sys.platform == "win32":
                os.startfile(path)
            else:
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.call([opener, path])
            self.update_status(f"Opening {os.path.basename(path)}...")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file.\n\n{e}")
            self.update_status("Error opening file.")

    def parse_page_range(self, range_str, max_pages):
        if not range_str:
            return list(range(max_pages))
        indices = set()
        parts = range_str.replace(" ", "").split(',')
        for part in parts:
            if '-' in part:
                try:
                    start, end = map(int, part.split('-'))
                    if not (1 <= start <= end <= max_pages):
                        raise ValueError(f"Invalid range '{part}': values out of bounds (1-{max_pages}).")
                    indices.update(range(start - 1, end))
                except (ValueError, TypeError):
                    raise ValueError(f"Invalid range format: '{part}'")
            else:
                try:
                    page_num = int(part)
                    if not (1 <= page_num <= max_pages):
                        raise ValueError(f"Page number {page_num} out of bounds (1-{max_pages}).")
                    indices.add(page_num - 1)
                except (ValueError, TypeError):
                    raise ValueError(f"Invalid page number: '{part}'")
        return sorted(list(indices))
    
    def add_pdfs(self):
        self._disable_undo()
        files = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF Files", "*.pdf")],
            initialdir=self.last_directory
        )
        if files:
            added_count = 0
            for f in files:
                if f.lower().endswith('.pdf') and not any(d['path'] == f for d in self.file_list):
                    self.file_list.append({'path': f, 'pages': None})
                    added_count += 1
            if added_count > 0:
                self.file_list_frame.update_list()
                self.update_status(f"Added {added_count} file(s).")
            self.save_last_directory(os.path.dirname(files[0]))

    def move_up(self):
        if self.selected_index > 0:
            self._disable_undo()
            self.file_list[self.selected_index], self.file_list[self.selected_index - 1] = \
                self.file_list[self.selected_index - 1], self.file_list[self.selected_index]
            self.select_file(self.selected_index - 1)
            self.update_status("File moved up.")
    
    def move_down(self):
        if 0 <= self.selected_index < len(self.file_list) - 1:
            self._disable_undo()
            self.file_list[self.selected_index], self.file_list[self.selected_index + 1] = \
                self.file_list[self.selected_index + 1], self.file_list[self.selected_index]
            self.select_file(self.selected_index + 1)
            self.update_status("File moved down.")
            
    def remove_selected(self):
        if 0 <= self.selected_index < len(self.file_list):
            if messagebox.askyesno("Confirm Remove", f"Are you sure you want to remove '{os.path.basename(self.file_list[self.selected_index]['path'])}'?", icon='warning'):
                item_to_remove = self.file_list[self.selected_index]
                self.last_removed_item = {'item': item_to_remove, 'index': self.selected_index}

                removed_file_name = os.path.basename(self.file_list.pop(self.selected_index)['path'])
                
                if self.selected_index >= len(self.file_list) and len(self.file_list) > 0:
                    self.selected_index = len(self.file_list) - 1
                elif not self.file_list:
                    self.selected_index = -1
                
                self.file_list_frame.update_list()
                self.undo_button.configure(state="normal")
                self.update_status(f"Removed: {removed_file_name}. Click Undo to restore.")

    def rotate_pages(self):
        if not (0 <= self.selected_index < len(self.file_list)):
            messagebox.showinfo("Info", "Select a PDF to rotate its pages.", parent=self.root)
            return

        file_item = self.file_list[self.selected_index]
        pdf_path = file_item['path']
        
        try:
            reader = self.get_pdf_reader_with_password(pdf_path)
            if not reader: return
            max_pages = len(reader.pages)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open {os.path.basename(pdf_path)}.\n\n{e}", parent=self.root)
            return

        dialog = RotationDialog(self.root, os.path.basename(pdf_path), max_pages)
        result = dialog.result
        
        if result:
            try:
                # Validate pages string
                if result['pages_str'].lower() != 'all':
                    self.parse_page_range(result['pages_str'], max_pages) # Use for validation
                
                file_item['rotation'] = result
                self.update_status(f"Rotation set for {os.path.basename(pdf_path)}.")
                self.file_list_frame.update_list()
            except ValueError as e:
                messagebox.showerror("Invalid Page Range", str(e), parent=self.root)

    def set_page_range(self):
        if not (0 <= self.selected_index < len(self.file_list)):
            messagebox.showinfo("Info", "Select a PDF to set a page range.", parent=self.root)
            return
        
        file_item = self.file_list[self.selected_index]
        pdf_path = file_item['path']

        try:
            reader = self.get_pdf_reader_with_password(pdf_path)
            if not reader: return
            max_pages = len(reader.pages)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open {os.path.basename(pdf_path)}.\n\n{e}", parent=self.root)
            return

        current_range = file_item.get('pages') or f"1-{max_pages}"

        dialog = CustomInputDialog(
            text=f"Enter page range for {os.path.basename(pdf_path)}\n(e.g., '1-5, 8, 10-12'). Total pages: {max_pages}",
            title="Set Page Range",
            parent=self.root
        )
        new_range = dialog.get_input()

        if new_range is not None:
            try:
                self.parse_page_range(new_range, max_pages)
                file_item['pages'] = new_range
                self.update_status(f"Set page range for {os.path.basename(pdf_path)}")
                self.file_list_frame.update_list()
            except ValueError as e:
                messagebox.showerror("Invalid Range", str(e), parent=self.root)

    def clear_metadata_fields(self):
        self.title_var.set("")
        self.author_var.set("")
        self.subject_var.set("")
        self.creator_var.set("")
        self.producer_var.set("")
        self.keywords_var.set("")
        self.creation_date_var.set("")
        self.mod_date_var.set("")
        self.password_var.set("")

    def preview_metadata(self):
        if not (0 <= self.selected_index < len(self.file_list)):
            messagebox.showinfo("Info", "Select a PDF to preview its metadata.", parent=self.root)
            return
        
        pdf_path = self.file_list[self.selected_index]['path']
        self.update_status(f"Previewing metadata for {os.path.basename(pdf_path)}...")
        try:
            reader = self.get_pdf_reader_with_password(pdf_path)
            if reader is None:
                self.update_status("Metadata preview cancelled.")
                return
            
            meta = reader.metadata
            self.title_var.set(meta.get('/Title', ''))
            self.author_var.set(meta.get('/Author', ''))
            self.subject_var.set(meta.get('/Subject', ''))
            self.creator_var.set(meta.get('/Creator', ''))
            self.producer_var.set(meta.get('/Producer', ''))
            self.keywords_var.set(meta.get('/Keywords', ''))
            
            creation_date = meta.get('/CreationDate', '')
            self.creation_date_var.set(creation_date[2:] if creation_date and creation_date.startswith("D:") else creation_date)
            
            mod_date = meta.get('/ModDate', '')
            self.mod_date_var.set(mod_date[2:] if mod_date and mod_date.startswith("D:") else mod_date)

            self.update_status(f"Metadata for {os.path.basename(pdf_path)} loaded.")
            self.tab_view.set("Metadata")
        except Exception as e:
            messagebox.showerror("Error", f"Could not read metadata from {os.path.basename(pdf_path)}.\n\n{e}", parent=self.root)
            self.update_status(f"Error reading metadata for {os.path.basename(pdf_path)}.")

    def get_pdf_reader_with_password(self, pdf_path):
        try:
            reader = PdfReader(pdf_path)
            if reader.is_encrypted:
                for _ in range(3):
                    dialog = CustomInputDialog(text=f"Enter password for {os.path.basename(pdf_path)}:", title="Password Required", parent=self.root)
                    password = dialog.get_input()
                    
                    if password is None: return None
                    try:
                        if reader.decrypt(password):
                            _ = reader.pages[0]
                            return reader
                    except Exception:
                        pass # Let the error be shown after loop
                messagebox.showerror("Error", "Incorrect password or failed to open PDF.", parent=self.root)
                return None
            return reader
        except Exception as e:
            messagebox.showerror("Error", f"Could not open PDF: {os.path.basename(pdf_path)}\n\n{e}", parent=self.root)
            return None

    def save_to_history(self, file_path, metadata):
        entry = {
            "file_path": file_path,
            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "metadata": metadata
        }
        try:
            if os.path.exists(HISTORY_FILE):
                with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                    history = json.load(f)
            else:
                history = []
            history.insert(0, entry)  # newest first
            with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
                json.dump(history, f, indent=2)
        except Exception as e:
            print(f"Failed to save history: {e}")
        self.load_history()
        self.refresh_history_ui()

    def load_history(self):
        self.history = []
        if os.path.exists(HISTORY_FILE):
            try:
                with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
                    self.history = json.load(f)
            except Exception as e:
                print(f"Failed to load history: {e}")
                self.history = []

    def refresh_history_ui(self):
        for label in self.history_labels:
            label.destroy()
        self.history_labels = []
        if not self.history:
            label = ctk.CTkLabel(self.history_listbox, text="No history yet.", text_color="gray50")
            label.pack(pady=10)
            self.history_labels.append(label)
            self.history_detail.configure(state="normal")
            self.history_detail.delete("1.0", "end")
            self.history_detail.insert("end", "Select a file to view its metadata.")
            self.history_detail.configure(state="disabled")
            return
        for i, entry in enumerate(self.history):
            entry_frame = ctk.CTkFrame(self.history_listbox)
            entry_frame.pack(fill="x", padx=5, pady=2)
            entry_frame.grid_columnconfigure(0, weight=1)

            display = f"{os.path.basename(entry['file_path'])} ({entry['timestamp']})"
            label = ctk.CTkLabel(entry_frame, text=display, anchor="w")
            label.grid(row=0, column=0, sticky="ew", padx=5)
            label.bind("<Button-1>", lambda e, idx=i: self.show_history_detail(idx))
            
            open_button = ctk.CTkButton(entry_frame, text="Open", width=60, command=lambda idx=i: self.open_history_file(idx))
            open_button.grid(row=0, column=1, padx=(0,5))
            
            delete_button = ctk.CTkButton(entry_frame, text="Delete", width=60, command=lambda idx=i: self.delete_history_entry(idx))
            delete_button.grid(row=0, column=2, padx=(0,5))
            
            self.history_labels.append(entry_frame)

        self.history_detail.configure(state="normal")
        self.history_detail.delete("1.0", "end")
        self.history_detail.insert("end", "Select a file to view its metadata.")
        self.history_detail.configure(state="disabled")

    def show_history_detail(self, idx):
        entry = self.history[idx]
        meta_str = "\n".join([f"{k.replace('/', '')}: {v}" for k, v in entry["metadata"].items()])
        detail_text = (
            f"File: {entry['file_path']}\n"
            f"Combined: {entry['timestamp']}\n\n"
            f"--- Metadata ---\n{meta_str}"
        )
        self.history_detail.configure(state="normal")
        self.history_detail.delete("1.0", "end")
        self.history_detail.insert("end", detail_text)
        self.history_detail.configure(state="disabled")

    def open_history_file(self, idx):
        if 0 <= idx < len(self.history):
            file_path = self.history[idx]['file_path']
            if os.path.exists(file_path):
                self.open_file(file_path)
            else:
                messagebox.showerror("Error", f"File not found:\n{file_path}", parent=self.root)

    def delete_history_entry(self, idx):
        if 0 <= idx < len(self.history):
            if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this history entry?"):
                self.history.pop(idx)
                self.save_history_data()
                self.refresh_history_ui()
                self.update_status("History entry deleted.")
    
    def clear_history(self):
        if self.history:
            if messagebox.askyesno("Confirm Clear History", "Are you sure you want to delete ALL history entries?\nThis action cannot be undone.", icon='warning'):
                self.history = []
                self.save_history_data()
                self.refresh_history_ui()
                self.update_status("History cleared.")
    
    def save_history_data(self):
        try:
            with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.history, f, indent=2)
        except Exception as e:
            print(f"Failed to save history data: {e}")

    def combine_pdfs(self):
        if not self.file_list:
            messagebox.showerror("Error", "No PDFs selected.", parent=self.root)
            self.update_status("Combine failed: No PDFs selected.")
            return

        self._disable_undo()
        self.progress_bar.pack(fill="x", padx=10, pady=(5,0))
        self.progress_bar.set(0)

        save_path = filedialog.asksaveasfilename(
            initialdir=self.last_directory,
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")]
        )
        if not save_path:
            self.update_status("Save cancelled.")
            return

        self.save_last_directory(os.path.dirname(save_path))
        self.update_status(f"Combining {len(self.file_list)} files...")
        writer = PdfWriter()
        
        try:
            for i, item in enumerate(self.file_list):
                pdf_path = item['path']
                page_range_str = item.get('pages')
                rotation_info = item.get('rotation')
                
                reader = self.get_pdf_reader_with_password(pdf_path)
                if reader is None:
                    raise Exception(f"Skipping file due to password failure: {os.path.basename(pdf_path)}")
                
                page_indices = self.parse_page_range(page_range_str, len(reader.pages))

                # Apply rotation if specified
                if rotation_info:
                    pages_to_rotate_str = rotation_info['pages_str']
                    angle = rotation_info['angle']
                    if pages_to_rotate_str.lower() == 'all':
                        pages_to_rotate_indices = range(len(reader.pages))
                    else:
                        pages_to_rotate_indices = self.parse_page_range(pages_to_rotate_str, len(reader.pages))
                    
                    for page_num in pages_to_rotate_indices:
                        reader.pages[page_num].rotate(angle)

                for page_num in page_indices:
                    writer.add_page(reader.pages[page_num])
                
                self.progress_bar.set((i + 1) / len(self.file_list))
                self.root.update_idletasks()

            metadata = {
                "/Title": self.title_var.get(), "/Author": self.author_var.get(),
                "/Subject": self.subject_var.get(), "/Creator": self.creator_var.get(),
                "/Producer": self.producer_var.get(), "/Keywords": self.keywords_var.get(),
                "/CreationDate": "D:" + self.creation_date_var.get() if self.creation_date_var.get() else "",
                "/ModDate": "D:" + self.mod_date_var.get() if self.mod_date_var.get() else ""
            }
            writer.add_metadata({k: v for k, v in metadata.items() if v})
            
            password = self.password_var.get()
            if password:
                writer.encrypt(password)

            with open(save_path, "wb") as f:
                writer.write(f)
            self.update_status("Successfully combined PDF saved.")
            messagebox.showinfo("Success", f"Combined PDF saved to:\n{save_path}", parent=self.root)
            
            if self.auto_open_var.get():
                self.open_file(save_path)
            
            # Save to history
            self.save_to_history(save_path, {k: v for k, v in metadata.items() if v})
            self.reset()
        except Exception as e:
            self.update_status(f"Error: {e}")
            messagebox.showerror("Error", f"Failed to combine PDFs.\n\n{e}", parent=self.root)
        finally:
            self.progress_bar.pack_forget()

    def reset(self):
        if self.file_list:
            if not messagebox.askyesno("Confirm Clear All", "Are you sure you want to clear the current file list?"):
                return
        self._disable_undo()
        self.file_list.clear()
        self.selected_index = -1
        self.file_list_frame.update_list()
        self.clear_metadata_fields()
        self.update_status("Ready")

if __name__ == "__main__":
    root = ctk.CTk()
    app = PDFCombinerApp(root)
    root.mainloop()