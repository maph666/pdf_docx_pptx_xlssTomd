import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path
from converter import convert_file


class App:
    def __init__(self, root):
        self.root = root
        root.title('Document → Markdown')

        frm = tk.Frame(root)
        frm.pack(padx=10, pady=10, fill='x')

        tk.Label(frm, text='Input file').grid(row=0, column=0, sticky='w')
        self.in_entry = tk.Entry(frm, width=60)
        self.in_entry.grid(row=0, column=1, padx=6)
        tk.Button(frm, text='Browse...', command=self.browse_in).grid(row=0, column=2)

        tk.Label(frm, text='Output .md').grid(row=1, column=0, sticky='w')
        self.out_entry = tk.Entry(frm, width=60)
        self.out_entry.grid(row=1, column=1, padx=6)
        tk.Button(frm, text='Browse...', command=self.browse_out).grid(row=1, column=2)



        self.convert_btn = tk.Button(frm, text='Convert', command=self.on_convert)
        self.convert_btn.grid(row=2, column=1, sticky='w', pady=8)

        self.status_var = tk.StringVar(value='Idle')
        tk.Label(frm, text='Status:').grid(row=2, column=0, sticky='w')
        tk.Label(frm, textvariable=self.status_var).grid(row=2, column=1, sticky='e')

        self.log = scrolledtext.ScrolledText(root, width=80, height=20, state='disabled')
        self.log.pack(padx=10, pady=(0,10))

    def browse_in(self):
        path = filedialog.askopenfilename(filetypes=[('Supported files','*.pdf *.docx *.pptx *.xlsx'),
                                                     ('PDF Files','*.pdf'),
                                                     ('Word Documents','*.docx'),
                                                     ('PowerPoint','*.pptx'),
                                                     ('Excel Files','*.xlsx'),
                                                     ('All Files','*.*')])
        if path:
            self.in_entry.delete(0, tk.END)
            self.in_entry.insert(0, path)
            # default out path: same name with .md
            out = Path(path).with_suffix('.md')
            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, str(out))

    def browse_out(self):
        path = filedialog.asksaveasfilename(defaultextension='.md', filetypes=[('Markdown','*.md')])
        if path:
            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, path)

    def append_log(self, text):
        self.log.configure(state='normal')
        self.log.insert(tk.END, text + '\n')
        self.log.configure(state='disabled')
        self.log.yview(tk.END)

    def on_convert(self):
        input_path = self.in_entry.get().strip()
        output_path = self.out_entry.get().strip()
        if not input_path:
            messagebox.showerror('Error', 'Selecciona un archivo de entrada (PDF/Word/PowerPoint/Excel)')
            return
        if not output_path:
            messagebox.showerror('Error', 'Selecciona un archivo de salida')
            return
        # disable UI
        self.convert_btn.config(state='disabled')
        self.status_var.set('Convirtiendo...')
        threading.Thread(target=self._worker, args=(input_path, output_path), daemon=True).start()

    def _worker(self, input_path, output_path):
        try:
            self.root.after(0, lambda: self.append_log(f'Converting {input_path} → {output_path}...'))
            out = convert_file(input_path, output_path)
        except Exception as e:
            self.root.after(0, lambda: self._on_done(False, str(e)))
        else:
            self.root.after(0, lambda: self._on_done(True, out))

    def _on_done(self, success, payload):
        if success:
            self.status_var.set(f'Hecho: {payload}')
            self.append_log(f'Wrote Markdown to {payload}')
        else:
            self.status_var.set('Error')
            self.append_log(f'Error: {payload}')
            messagebox.showerror('Error', payload)
        self.convert_btn.config(state='normal')


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == '__main__':
    main()