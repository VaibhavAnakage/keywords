import os
import threading
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from keyword_definitions import groups


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


def process_excel(input_path, progress_callback=None):
    df = pd.read_excel(input_path, sheet_name='Data')
    df['Description'] = df['Description'].apply(lambda x: f' {x} ' if pd.notna(x) else x)
    df['Resolution'] = df['Resolution'].apply(lambda x: f' {x} ' if pd.notna(x) else x)
    pro_priority_groups = [g for g in groups if g.get('Priority-Type', [''])[0] == 'Pro']
    high_priority_groups = [g for g in groups if g.get('Priority-Type', [''])[0] == 'High']
    low_priority_groups = [g for g in groups if g not in pro_priority_groups + high_priority_groups]
    for col in ['GroupDescription', 'SubGroupDescription', 'Addressable', 'Priority-Type', 'Matched Keywords']:
        if col not in df.columns:
            df[col] = ''
    total = len(df)
    for i, (index, row) in enumerate(df.iterrows(), start=1):
        if df.at[index, 'GroupDescription'] == '':
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
                        break
        if progress_callback and (i % 10 == 0 or i == total):
            progress_callback(i, total)
    base, ext = os.path.splitext(input_path)
    output_path = f"{base}_processed.xlsx"
    df.to_excel(output_path, index=False)
    return output_path


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Logs Analysis Processor')
        self.geometry('560x220')
        self.resizable(False, False)
        self.file_path = tk.StringVar()
        self.status_text = tk.StringVar(value='Select an Excel file to process.')
        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self, padding=16)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text='Selected file:').grid(row=0, column=0, sticky='w')
        self.path_entry = ttk.Entry(frm, textvariable=self.file_path, width=60)
        self.path_entry.grid(row=1, column=0, columnspan=2, sticky='we', pady=(4, 8))
        ttk.Button(frm, text='Browse...', command=self.browse).grid(row=1, column=2, padx=(8, 0))
        self.process_btn = ttk.Button(frm, text='Process', command=self.start_processing)
        self.process_btn.grid(row=2, column=0, sticky='w')
        self.progress = ttk.Progressbar(frm, orient='horizontal', mode='determinate', length=400)
        self.progress.grid(row=3, column=0, columnspan=3, sticky='we', pady=(12, 0))
        self.status_lbl = ttk.Label(frm, textvariable=self.status_text)
        self.status_lbl.grid(row=4, column=0, columnspan=3, sticky='w', pady=(8, 0))
        for i in range(3):
            frm.columnconfigure(i, weight=1)

    def browse(self):
        path = filedialog.askopenfilename(
            title='Select Excel file',
            filetypes=[('Excel files', '*.xlsx *.xls')]
        )
        if path:
            self.file_path.set(path)
            self.status_text.set('Ready to process.')

    def start_processing(self):
        path = self.file_path.get().strip()
        if not path:
            messagebox.showwarning('No file', 'Please select an Excel file to process.')
            return
        if not os.path.exists(path):
            messagebox.showerror('File not found', 'The selected file does not exist.')
            return
        self.process_btn.config(state=tk.DISABLED)
        self.status_text.set('Processing...')
        self.progress['value'] = 0
        self.progress['maximum'] = 100
        def run():
            try:
                def progress_callback(done, total):
                    pct = int((done / total) * 100) if total else 0
                    self.after(0, lambda: self._update_progress(pct))
                output = process_excel(path, progress_callback)
                self.after(0, lambda: self._on_done(success=True, output=output))
            except Exception as e:
                self.after(0, lambda: self._on_done(success=False, error=str(e)))
        threading.Thread(target=run, daemon=True).start()

    def _update_progress(self, pct):
        self.progress['value'] = pct
        self.status_text.set(f'Processing... {pct}%')

    def _on_done(self, success, output=None, error=None):
        self.process_btn.config(state=tk.NORMAL)
        if success:
            self.progress['value'] = 100
            self.status_text.set('Completed.')
            messagebox.showinfo('Done', f'Output saved to:\n{output}')
        else:
            self.status_text.set('Failed.')
            messagebox.showerror('Error', error)


if __name__ == '__main__':
    app = App()
    app.mainloop()
