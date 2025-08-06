#Copyright 2025-present Linuxlawyer.com
#Licensed under the GPL v3.0 license

import os
import mailbox
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from pathlib import Path
from fpdf import FPDF
import threading
import time
import gc

# Optimized PDF creation function
def email_to_pdf(email_content, output_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=10)

    lines = email_content.splitlines()
    for line in lines:
        if pdf.get_y() > 270:
            pdf.add_page()
        pdf.multi_cell(0, 5, line)

    pdf.output(output_path)
    del pdf
    gc.collect()

# Main Application Class
class EDiscoveryApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Enhanced eDiscovery Tool')
        self.root.geometry('900x600')
        self.directory = ''
        self.setup_ui()

    def setup_ui(self):
        content_frame = ttk.Frame(self.root)
        content_frame.pack(fill='both', expand=True, padx=10, pady=10)

        frame_left = ttk.Frame(content_frame)
        frame_left.pack(side='left', fill='both', expand=True, padx=(0, 5))

        ttk.Button(frame_left, text='Select Directory', command=self.select_directory).pack(pady=10)

        self.tree = ttk.Treeview(frame_left)
        self.tree.pack(fill='both', expand=True)

        frame_right = ttk.Frame(content_frame)
        frame_right.pack(side='right', fill='y', padx=(5, 0))

        ttk.Label(frame_right, text='Search Terms').pack(pady=5)
        self.search_entries = [ttk.Entry(frame_right, width=50) for _ in range(3)]
        for entry in self.search_entries:
            entry.pack(pady=2)

        ttk.Label(frame_right, text='Output Format').pack(pady=5)
        self.format_var = tk.StringVar(value='original')
        ttk.Combobox(frame_right, textvariable=self.format_var,
                     values=['original', 'pdf', 'text']).pack(pady=5)

        ttk.Button(frame_right, text='Generate', command=self.start_processing).pack(pady=20)

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor='w')
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def select_directory(self):
        self.directory = filedialog.askdirectory()
        if self.directory:
            self.populate_tree()

    def populate_tree(self):
        self.tree.delete(*self.tree.get_children())
        path = Path(self.directory)

        def insert_nodes(parent, path):
            for p in path.iterdir():
                node = self.tree.insert(parent, 'end', text=p.name, open=False)
                if p.is_dir():
                    insert_nodes(node, p)

        insert_nodes('', path)

    def update_status(self, message):
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.status_var.set(f"[{timestamp}] {message}")
        self.root.update_idletasks()

    def start_processing(self):
        threading.Thread(target=self.generate, daemon=True).start()

    def generate(self):
        start_time = time.time()

        if not self.directory:
            messagebox.showerror('Error', 'Please select a directory first.')
            return

        terms = [entry.get().strip().lower() for entry in self.search_entries if entry.get().strip()]
        if not terms:
            messagebox.showerror('Error', 'Please enter at least one search term.')
            return

        output_format = self.format_var.get()
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_dir = Path(self.directory) / f'ediscovery_output_{timestamp}'
        output_dir.mkdir(parents=True, exist_ok=True)

        email_count = 0
        errors = []

        mbox_files = list(Path(self.directory).rglob('*.mbox'))

        for mbox_path in mbox_files:
            file = mbox_path.name
            self.update_status(f"Processing {file}")
            try:
                mbox = mailbox.mbox(str(mbox_path))
                for idx, message in enumerate(mbox):
                    self.update_status(f"Processing email {idx+1} in {file}")
                    try:
                        email_body = message.get_payload(decode=True)
                        email_body = email_body.decode('utf-8', errors='ignore').lower() if email_body else ""

                        sender = str(message.get('from', '')).lower()
                        receiver = str(message.get('to', '')).lower()
                        subject = str(message.get('subject', '')).lower()

                        email_content_combined = f"{sender}\n{receiver}\n{subject}\n{email_body}"

                        if any(term in email_content_combined for term in terms):
                            email_count += 1
                            email_filename = f'email_{email_count}'
                            out_path = output_dir / f'{email_filename}.{output_format if output_format != "original" else "eml"}'

                            content = message.as_string()
                            if output_format == 'original':
                                with open(out_path, 'w', encoding='utf-8') as f:
                                    f.write(content)
                            elif output_format == 'pdf':
                                email_to_pdf(content, out_path)
                            elif output_format == 'text':
                                with open(out_path, 'w', encoding='utf-8') as f:
                                    f.write(content)

                            del content
                        del email_body, sender, receiver, subject, email_content_combined
                        gc.collect()
                    except Exception as e:
                        errors.append(f"Error processing email in {file}: {e}")
                del mbox
                gc.collect()
            except Exception as e:
                errors.append(f"Error processing file {file}: {e}")

        if errors:
            with open(output_dir / "errors.txt", 'w') as error_file:
                error_file.write("\n".join(errors))

        elapsed_time = time.time() - start_time
        self.update_status(f"Completed {email_count} emails in {elapsed_time:.2f} seconds.")
        messagebox.showinfo('Completed', f'Found and extracted {email_count} matching emails in {elapsed_time:.2f} seconds.')

if __name__ == '__main__':
    root = tk.Tk()
    app = EDiscoveryApp(root)
    root.mainloop()
