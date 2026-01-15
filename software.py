import os
import threading
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import tkinter.font as tkfont
import urllib.request
import tempfile
import importlib.util
from pandas import ExcelFile


def check_keywords(text, group):
    main_keywords = group['MainKeywords']
    group_keywords = [group[key] for key in group if key.startswith('GroupKeywords')]
    exclusion_patterns = group.get('ExclusionPatterns', [])
    text_lower = text.lower()
    for pattern in exclusion_patterns:
        if re.search(pattern, text_lower):
            return None
    matched_keywords = []

    def keyword_pattern(keyword):
        return re.escape(keyword.lower())

    for keyword in main_keywords:
        if re.search(keyword_pattern(keyword), text_lower):
            matched_keywords.append(keyword)
        else:
            return None
    for keyword_set in group_keywords:
        group_matched = False
        for keyword in keyword_set:
            if re.search(keyword_pattern(keyword), text_lower):
                matched_keywords.append(keyword)
                group_matched = True
                break
        if not group_matched:
            return None
    return {
        'GroupDescription': group['GroupDescription'],
        'SubGroupDescription': group['SubGroupDescription'],
        'Addressable': group['Addressable'],
        'Priority-Type': group.get('Priority-Type', [''])[0],
        'Matched Keywords': ', '.join(matched_keywords)
    }

def ensure_keywords_file():
    """Download keyword_definitions.py into %TEMP%/la and return its absolute path."""
    temp_dir = os.path.join(tempfile.gettempdir(), 'la')
    os.makedirs(temp_dir, exist_ok=True)
    dest_path = os.path.join(temp_dir, 'keyword_definitions.py')

    # Try main then master branch raw URLs
    base_raw = 'https://raw.githubusercontent.com/VaibhavAnakage/keywords'
    candidates = [
        f"{base_raw}/main/keyword_definitions.py",
        f"{base_raw}/master/keyword_definitions.py",
    ]
    last_error = None
    for url in candidates:
        try:
            with urllib.request.urlopen(url, timeout=20) as resp:
                content = resp.read()
            with open(dest_path, 'wb') as f:
                f.write(content)
            return dest_path
        except Exception as e:
            last_error = e
            continue
    raise RuntimeError(f"Failed to download keyword_definitions.py: {last_error}")


def load_groups_from_file(py_path):
    """Dynamically import keyword_definitions.py from a path and return groups."""
    spec = importlib.util.spec_from_file_location('keyword_definitions_dynamic', py_path)
    if spec is None or spec.loader is None:
        raise RuntimeError('Unable to prepare loader for keyword_definitions.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    if not hasattr(module, 'groups'):
        raise RuntimeError('Downloaded keyword_definitions.py does not define "groups"')
    return module.groups


def process_excel(input_path, progress_callback=None):
    # Ensure latest keyword_definitions and load groups
    keywords_path = ensure_keywords_file()
    groups = load_groups_from_file(keywords_path)
    sheet_name = find_data_sheet_name(input_path)
    if not sheet_name:
        raise RuntimeError("Sheet named 'data' (any case) not found.")
    df = pd.read_excel(input_path, sheet_name=sheet_name)
    df['Description'] = df['Description'].apply(lambda x: f' {x} ' if pd.notna(x) else x)
    df['Resolution'] = df['Resolution'].apply(lambda x: f' {x} ' if pd.notna(x) else x)
    pro_priority_groups = [g for g in groups if g.get('Priority-Type', [''])[0] == 'Pro']
    high_priority_groups = [g for g in groups if g.get('Priority-Type', [''])[0] == 'High']
    low_priority_groups = [g for g in groups if g not in pro_priority_groups + high_priority_groups]
    # Track which columns were present originally to decide processing scope
    original_columns = set(df.columns)
    for col in ['GroupDescription', 'SubGroupDescription', 'Addressable', 'Priority-Type', 'Matched Keywords']:
        if col not in df.columns:
            df[col] = ''
    # Decide rows to consider
    def _is_empty(v):
        return (pd.isna(v)) or (isinstance(v, str) and v.strip() == '')
    if 'Addressable' in original_columns:
        to_consider = [idx for idx, v in df['Addressable'].items() if _is_empty(v)]
    else:
        to_consider = [idx for idx, v in df['GroupDescription'].items() if _is_empty(v)]
    total_considered = len(to_consider)
    if progress_callback:
        progress_callback(0, total_considered)
    unmatched_indices = []
    matched_count = 0
    for i, index in enumerate(to_consider, start=1):
        row = df.loc[index]
        description = str(row['Description'])
        resolution = str(row['Resolution'])
        matched = False
        for group in pro_priority_groups:
            result = check_keywords(description, group) or check_keywords(resolution, group)
            if result:
                for key in result:
                    df.at[index, key] = result[key]
                matched = True
                break
        if not matched:
            for group in high_priority_groups:
                result = check_keywords(description, group) or check_keywords(resolution, group)
                if result:
                    for key in result:
                        df.at[index, key] = result[key]
                    matched = True
                    break
        if not matched:
            for group in low_priority_groups:
                result = check_keywords(description, group) or check_keywords(resolution, group)
                if result:
                    for key in result:
                        df.at[index, key] = result[key]
                    matched = True
                    break
        if not matched:
            unmatched_indices.append(index)
        else:
            matched_count += 1
        if progress_callback and (i % 10 == 0 or i == total_considered):
            progress_callback(i, total_considered)
    base, ext = os.path.splitext(input_path)
    output_path = f"{base}_processed.xlsx"
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

    unmatched_rows = df.loc[unmatched_indices, ['Description', 'Resolution']].copy() if unmatched_indices else pd.DataFrame(columns=['Description', 'Resolution'])
    return {
        'output_path': output_path,
        'considered': total_considered,
        'matched': matched_count,
        'unmatched_rows': unmatched_rows.to_dict(orient='records'),
    }

def find_data_sheet_name(path):
    """Return the actual sheet name that matches 'data' case-insensitively, or None."""
    with ExcelFile(path) as xf:
        for name in xf.sheet_names:
            if str(name).lower() == 'data':
                return name
    return None


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Log Analyzer')
        self.geometry('720x360')
        self.resizable(False, False)
        self.file_path = tk.StringVar()
        self.status_text = tk.StringVar(value='Select an Excel file to process.')
        self._apply_style()
        self._build_menu()
        self._build_ui()
        self._center_window()
        self.bind('<Return>', lambda e: self.start_processing())

    def _build_ui(self):
        frm = ttk.Frame(self, padding=16, style='Main.TFrame')
        frm.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(frm, text='Log Analyzer', font=('Calibri', 16, 'bold'), style='Title.TLabel')
        title.grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 8))

        instr = (
            "Instructions:\n"
            "- Make sure the file extension is .xlsx.\n"
            "- The sheet name should be 'data' (any case: data/Data/DATA).\n"
            "- Ensure that the 'data' sheet contains two columns: 'Description' and 'Resolution'."
        )
        ttk.Label(frm, text=instr, justify='left', style='Main.TLabel').grid(row=1, column=0, columnspan=3, sticky='w')

        ttk.Label(frm, text='Selected file:', style='Main.TLabel').grid(row=2, column=0, sticky='w', pady=(8, 0))
        self.path_entry = ttk.Entry(frm, textvariable=self.file_path, width=70)
        self.path_entry.grid(row=3, column=0, columnspan=2, sticky='we', pady=(4, 8))
        ttk.Button(frm, text='Browse...', command=self.browse).grid(row=3, column=2, padx=(8, 0))

        btns = ttk.Frame(frm, style='Main.TFrame')
        btns.grid(row=4, column=0, columnspan=3, sticky='we')
        self.process_btn = ttk.Button(btns, text='Process', command=self.start_processing, style='Accent.TButton')
        self.process_btn.pack(side=tk.LEFT)
        self.clear_btn = ttk.Button(btns, text='Clear', command=self.clear_selection)
        self.clear_btn.pack(side=tk.LEFT, padx=(8, 0))

        self.progress = ttk.Progressbar(frm, orient='horizontal', mode='determinate', length=560, style='Slim.Horizontal.TProgressbar')
        self.progress.grid(row=5, column=0, columnspan=3, sticky='we', pady=(12, 0))
        # Centered percentage text overlay on the progress bar
        self.progress_text = ttk.Label(frm, text='0%', font=('Calibri', 10, 'bold'))
        self.progress_text.place(in_=self.progress, relx=0.5, rely=0.5, anchor='center')
        self.status_lbl = ttk.Label(frm, textvariable=self.status_text, style='Main.TLabel')
        self.status_lbl.grid(row=6, column=0, columnspan=3, sticky='w', pady=(8, 0))

        for i in range(3):
            frm.columnconfigure(i, weight=1)

        # Status bar
        self.statusbar = ttk.Frame(self, style='Main.TFrame')
        self.statusbar.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Separator(self.statusbar, orient='horizontal').pack(fill=tk.X, side=tk.TOP)
        self.sb_label = ttk.Label(self.statusbar, text='Ready', anchor='w', style='Main.TLabel')
        self.sb_label.pack(fill=tk.X, padx=8, pady=4)

    def _apply_style(self):
        style = ttk.Style()
        try:
            # Prefer clam on Windows for modern look
            style.theme_use('clam')
        except Exception:
            pass
        # App background color (white)
        base_color = '#FFFFFF'
        try:
            self.configure(bg=base_color)
        except Exception:
            pass
        # Set default fonts to Calibri
        try:
            default_font = tkfont.nametofont('TkDefaultFont')
            default_font.configure(family='Calibri', size=11)
            text_font = tkfont.nametofont('TkTextFont')
            text_font.configure(family='Calibri', size=11)
            fixed_font = tkfont.nametofont('TkFixedFont')
            fixed_font.configure(family='Consolas', size=11)
            heading_font = tkfont.nametofont('TkHeadingFont')
            heading_font.configure(family='Calibri', size=12, weight='bold')
        except Exception:
            pass
        # Themed backgrounds
        style.configure('Main.TFrame', background=base_color)
        style.configure('Main.TLabel', background=base_color, font=('Calibri', 11))
        style.configure('Title.TLabel', background=base_color, font=('Calibri', 16, 'bold'))
        style.configure('Accent.TButton', font=('Calibri', 11, 'bold'))
        # Improved progress bar styling
        style.configure('Slim.Horizontal.TProgressbar', thickness=14, background='#4F8EF7', troughcolor='#E6E6E6')

    def _build_menu(self):
        menubar = tk.Menu(self)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label='Exit', command=self.destroy)
        menubar.add_cascade(label='File', menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label='Instructions', command=self._show_instructions)
        help_menu.add_command(label='About', command=lambda: messagebox.showinfo('About', 'Log Analyzer'))
        menubar.add_cascade(label='Help', menu=help_menu)
        self.config(menu=menubar)

    def _show_instructions(self):
        message = (
            'Instructions:\n'
            '- Make sure the file extension is .xlsx.\n'
            "- The sheet name should be 'data' (any case: data/Data/DATA).\n"
            "- Ensure the 'data' sheet contains columns: 'Description' and 'Resolution'."
        )
        messagebox.showinfo('Instructions', message)

    def _center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw // 2) - (w // 2)
        y = (sh // 2) - (h // 2)
        self.geometry(f'+{x}+{y}')

    def browse(self):
        path = filedialog.askopenfilename(
            title='Select Excel file',
            filetypes=[('Excel .xlsx', '*.xlsx')]
        )
        if path:
            self.file_path.set(path)
            self.status_text.set('Ready to validate and process.')

    def clear_selection(self):
        self.file_path.set('')
        self.status_text.set('Select an Excel file to process.')

    def validate_input_file(self, path):
        if not path.lower().endswith('.xlsx'):
            return False, "File extension must be .xlsx"
        if not os.path.exists(path):
            return False, "The selected file does not exist."
        try:
            sheet_name = find_data_sheet_name(path)
            if not sheet_name:
                return False, "Sheet named 'data' (any case) not found."
            header_df = pd.read_excel(path, sheet_name=sheet_name, nrows=0)
            cols = set(map(str, header_df.columns))
            required = {'Description', 'Resolution'}
            missing = required - cols
            if missing:
                return False, f"Missing required columns in 'data' sheet: {', '.join(sorted(missing))}"
        except Exception as e:
            return False, f"Failed to open or validate Excel file: {e}"
        return True, None

    def start_processing(self):
        path = self.file_path.get().strip()
        if not path:
            messagebox.showwarning('No file', 'Please select an Excel file to process.')
            return

        ok, err = self.validate_input_file(path)
        if not ok:
            messagebox.showerror('Validation error', err)
            return
        self._set_busy(True)
        self.status_text.set('Downloading keywords and processing...')
        self.progress['value'] = 0
        self.progress['maximum'] = 100
        def run():
            try:
                def progress_callback(done, total):
                    pct = int((done / total) * 100) if total else 0
                    self.after(0, lambda: self._update_progress(pct, done, total))
                result = process_excel(path, progress_callback)
                self.after(0, lambda: self._on_done(success=True, output=result))
            except Exception as e:
                self.after(0, lambda e=e: self._on_done(success=False, error=str(e)))
        threading.Thread(target=run, daemon=True).start()

    def _update_progress(self, pct, done=None, total=None):
        self.progress['value'] = pct
        # Keep percent only on the bar overlay
        try:
            self.progress_text.config(text=f'{pct}%')
        except Exception:
            pass
        if done is not None and total is not None:
            self.status_text.set(f'Ticket {done} of {total}')

    def _on_done(self, success, output=None, error=None):
        self._set_busy(False)
        if success:
            self.progress['value'] = 100
            self.status_text.set('Completed.')
            # Show summary popup and update status
            if isinstance(output, dict):
                total = output.get('considered', 0)
                matched = output.get('matched', 0)
                unmatched = total - matched
                pct = round((matched / total) * 100, 2) if total else 0
                summary = (
                    f"Output saved to:\n{output.get('output_path','')}\n\n"
                    f"Matched: {matched}/{total} ({pct}%)\nUnmatched: {unmatched}"
                )
                messagebox.showinfo('Done', summary)
                self.status_text.set(f"Completed. Matched: {matched}/{total} ({pct}%). Unmatched: {unmatched}. Output: {output.get('output_path','')}")
            else:
                messagebox.showinfo('Done', f'Output saved to:\n{output}')
                self.status_text.set(f'Completed. Output: {output}')
        else:
            self.status_text.set('Failed.')
            messagebox.showerror('Error', error)

    def _set_busy(self, busy: bool):
        state = tk.DISABLED if busy else tk.NORMAL
        self.process_btn.config(state=state)
        self.clear_btn.config(state=state)
        # Disable browse by overlaying state on the entry and controlling button via grid_slaves lookup
        for child in self.children.values():
            pass
        # Find the browse button in UI
        # Since we did not store it, iterate widgets in the toplevel frame to disable
        try:
            for w in self.winfo_children():
                for sub in w.winfo_children():
                    if isinstance(sub, ttk.Button) and sub['text'] == 'Browse...':
                        sub.config(state=state)
        except Exception:
            pass
        self.path_entry.config(state=tk.NORMAL if not busy else tk.DISABLED)
        # Status bar shows general state only
        self.sb_label.config(text='Processing...' if busy else 'Ready')


if __name__ == '__main__':
    app = App()
    app.mainloop()
