import os
from tkinter import Tk, Label, Button, Listbox, filedialog, END, messagebox, SINGLE
from tkinterdnd2 import DND_FILES, TkinterDnD
from docx import Document
from docxcompose.composer import Composer

try:
    import comtypes.client  # for PDF conversion (Windows only)
    has_comtypes = True
except ImportError:
    has_comtypes = False

class WordMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word File Merger")
        self.root.geometry("600x450")

        Label(root, text="Drag and drop .docx files below:", font=("Arial", 14)).pack(pady=10)

        self.file_listbox = Listbox(root, selectmode=SINGLE, width=70, height=15)
        self.file_listbox.pack(padx=10, pady=10)
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.drop_files)

        Button(root, text="Combine to Word", command=self.combine_documents).pack(pady=5)
        Button(root, text="Convert Combined to PDF", command=self.convert_to_pdf).pack(pady=5)

        Button(root, text="Remove Selected File", command=self.remove_selected_file).pack(pady=5)
        Button(root, text="Clear All Files", command=self.clear_all_files).pack(pady=5)

    def drop_files(self, event):
        files = root.tk.splitlist(event.data)
        for file_path in files:
            if file_path.endswith(".docx") and file_path not in self.file_listbox.get(0, END):
                self.file_listbox.insert(END, file_path)

    def remove_selected_file(self):
        selected = self.file_listbox.curselection()
        if selected:
            self.file_listbox.delete(selected[0])

    def clear_all_files(self):
        self.file_listbox.delete(0, END)

    def combine_documents(self):
        files = self.file_listbox.get(0, END)
        if not files:
            messagebox.showerror("Error", "No files to combine.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if not save_path:
            return

        # Start with the first document
        master = Document(files[0])
        composer = Composer(master)

        for file in files[1:]:
            doc = Document(file)
            composer.append(doc)

        composer.save(save_path)
        messagebox.showinfo("Success", f"Combined file saved to:\n{save_path}")
        self.last_combined_path = save_path

    def convert_to_pdf(self):
        if not has_comtypes:
            messagebox.showerror("Missing Dependency", "PDF conversion requires 'comtypes' and Microsoft Word (Windows only).")
            return

        if not hasattr(self, 'last_combined_path'):
            messagebox.showwarning("No File", "Please combine and save a Word file first.")
            return

        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(self.last_combined_path)

        pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not pdf_path:
            doc.Close()
            word.Quit()
            return

        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        word.Quit()
        messagebox.showinfo("Success", f"PDF saved to:\n{pdf_path}")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = WordMergerApp(root)
    root.mainloop()
