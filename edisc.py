import os
import mailbox
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from pathlib import Path
from fpdf import FPDF

# Utility function for creating PDFs from emails
def email_to_pdf(email_content, output_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.multi_cell(0, 10, email_content)
    pdf.output(output_path)

# Main Application Class
class EDiscoveryApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Simple eDiscovery Tool')
        self.root.geometry('800x500')

        self.directory = ''

        # UI Layout
        self.setup_ui()

    def setup_ui(self):
        # Left side - Directory and Hierarchy
        frame_left = ttk.Frame(self.root)
        frame_left.pack(side='left', fill='y', padx=10, pady=10)

        ttk.Button(frame_left, text='Select Directory', command=self.select_directory).pack(pady=10)

        self.tree = ttk.Treeview(frame_left)
        self.tree.pack(fill='y', expand=True)

        # Right side - Search Terms and Options
        frame_right = ttk.Frame(self.root)
        frame_right.pack(side='right', fill='both', expand=True, padx=10, pady=10)

        ttk.Label(frame_right, text='Search Terms').pack(pady=5)
        self.search_entries = [ttk.Entry(frame_right, width=40) for _ in range(3)]
        for entry in self.search_entries:
            entry.pack(pady=2)

        ttk.Label(frame_right, text='Output Format').pack(pady=5)
        self.format_var = tk.StringVar(value='original')
        ttk.Combobox(frame_right, textvariable=self.format_var,
                     values=['original', 'pdf', 'text']).pack(pady=5)

        ttk.Button(frame_right, text='Generate', command=self.generate).pack(pady=20)

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

    def generate(self):
        if not self.directory:
            messagebox.showerror('Error', 'Please select a directory first.')
            return

        terms = [entry.get().strip() for entry in self.search_entries if entry.get().strip()]
        if not terms:
            messagebox.showerror('Error', 'Please enter at least one search term.')
            return

        output_format = self.format_var.get()

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_dir = Path(self.directory) / f'ediscovery_output_{timestamp}'
        output_dir.mkdir(parents=True, exist_ok=True)

        email_count = 0

        for root, _, files in os.walk(self.directory):
            for file in files:
                if file.lower().endswith('.mbox'):
                    mbox_path = os.path.join(root, file)
                    mbox = mailbox.mbox(mbox_path)
                    for message in mbox:
                        content = message.as_string()
                        if any(term.lower() in content.lower() for term in terms):
                            email_count += 1
                            email_filename = f'email_{email_count}'
                            if output_format == 'original':
                                out_path = output_dir / f'{email_filename}.eml'
                                with open(out_path, 'w', encoding='utf-8') as f:
                                    f.write(content)
                            elif output_format == 'pdf':
                                out_path = output_dir / f'{email_filename}.pdf'
                                email_to_pdf(content, out_path)
                            elif output_format == 'text':
                                out_path = output_dir / f'{email_filename}.txt'
                                with open(out_path, 'w', encoding='utf-8') as f:
                                    f.write(content)

        messagebox.showinfo('Completed', f'Found and extracted {email_count} matching emails.')

# Run the application
if __name__ == '__main__':
    root = tk.Tk()
    app = EDiscoveryApp(root)
    root.mainloop()
