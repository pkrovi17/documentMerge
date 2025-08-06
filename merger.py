import os
from tkinter import Tk, Label, Button, Listbox, filedialog, END, messagebox, SINGLE, Frame, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from docx import Document
from docxcompose.composer import Composer

class WordMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Merger - Dark Theme")
        self.root.geometry("700x500")
        self.root.configure(bg='#1e1e1e')
        
        # Configure dark theme colors
        self.colors = {
            'bg_dark': '#1e1e1e',
            'bg_medium': '#2d2d2d',
            'bg_light': '#3d3d3d',
            'text_primary': '#ffffff',
            'text_secondary': '#cccccc',
            'accent_amber': '#ffb74d',
            'accent_amber_hover': '#ffa726',
            'border': '#404040',
            'success': '#4caf50',
            'error': '#f44336'
        }
        
        # Initialize file paths list
        self.file_paths = []
        
        # Configure the main window
        self.root.configure(bg=self.colors['bg_dark'])
        self.setup_styles()
        self.create_widgets()
        
    def setup_styles(self):
        """Configure custom styles for the application"""
        # Configure ttk styles if available
        try:
            style = ttk.Style()
            style.theme_use('clam')
            
            # Configure button styles
            style.configure('Amber.TButton',
                          background=self.colors['accent_amber'],
                          foreground=self.colors['bg_dark'],
                          borderwidth=0,
                          focuscolor='none',
                          font=('Segoe UI', 10, 'bold'))
            
            style.map('Amber.TButton',
                     background=[('active', self.colors['accent_amber_hover']),
                                ('pressed', self.colors['accent_amber_hover'])])
            
        except:
            pass  # Fallback to regular tkinter styling
    
    def create_widgets(self):
        """Create and configure all UI widgets"""
        # Main container
        main_frame = Frame(self.root, bg=self.colors['bg_dark'])
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Title with gradient-like effect
        title_frame = Frame(main_frame, bg=self.colors['bg_dark'])
        title_frame.pack(fill='x', pady=(0, 10))
        
        title_label = Label(title_frame, 
                           text="Document Merger", 
                           font=('Segoe UI', 28, 'bold'),
                           fg=self.colors['accent_amber'],
                           bg=self.colors['bg_dark'])
        title_label.pack()
        
        # Subtitle
        subtitle_label = Label(main_frame,
                              text="Drag and drop .docx files below to merge them",
                              font=('Segoe UI', 12),
                              fg=self.colors['text_secondary'],
                              bg=self.colors['bg_dark'])
        subtitle_label.pack(pady=(0, 25))
        
        # File list container with border
        list_frame = Frame(main_frame, bg=self.colors['bg_medium'], relief='flat', bd=1)
        list_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # File listbox with custom styling
        self.file_listbox = Listbox(list_frame, 
                                   selectmode=SINGLE, 
                                   width=70, 
                                   height=12,
                                   font=('Segoe UI', 10),
                                   bg=self.colors['bg_light'],
                                   fg=self.colors['text_primary'],
                                   selectbackground=self.colors['accent_amber'],
                                   selectforeground=self.colors['bg_dark'],
                                   relief='flat',
                                   bd=0,
                                   highlightthickness=1,
                                   highlightcolor=self.colors['accent_amber'],
                                   highlightbackground=self.colors['border'])
        self.file_listbox.pack(fill='both', expand=True, padx=10, pady=10)
        self.file_listbox.drop_target_register(DND_FILES)
        self.file_listbox.dnd_bind('<<Drop>>', self.drop_files)
        
        # Buttons container
        button_frame = Frame(main_frame, bg=self.colors['bg_dark'])
        button_frame.pack(pady=25)
        
        # Styled buttons
        self.create_styled_button(button_frame, "Combine Documents", self.combine_documents, 0)
        self.create_styled_button(button_frame, "Remove Selected", self.remove_selected_file, 1)
        self.create_styled_button(button_frame, "Clear All Files", self.clear_all_files, 2)
    
    def create_styled_button(self, parent, text, command, column):
        """Create a styled button with amber theme"""
        button = Button(parent, 
                       text=text,
                       command=command,
                       font=('Segoe UI', 11, 'bold'),
                       bg=self.colors['accent_amber'],
                       fg=self.colors['bg_dark'],
                       activebackground=self.colors['accent_amber_hover'],
                       activeforeground=self.colors['bg_dark'],
                       relief='flat',
                       bd=0,
                       padx=25,
                       pady=12,
                       cursor='hand2')
        button.pack(side='left', padx=12)
        
        # Add hover effects
        button.bind('<Enter>', lambda e: button.configure(bg=self.colors['accent_amber_hover']))
        button.bind('<Leave>', lambda e: button.configure(bg=self.colors['accent_amber']))
        
        return button

    def drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            if file_path.endswith(".docx") and file_path not in self.file_paths:
                # Add to file paths list
                self.file_paths.append(file_path)
                # Display filename with icon
                filename = os.path.basename(file_path)
                self.file_listbox.insert(END, f"📄 {filename}")

    def remove_selected_file(self):
        selected = self.file_listbox.curselection()
        if selected:
            index = selected[0]
            # Remove from both listbox and file_paths
            self.file_listbox.delete(index)
            if index < len(self.file_paths):
                self.file_paths.pop(index)

    def clear_all_files(self):
        self.file_listbox.delete(0, END)
        self.file_paths.clear()

    def combine_documents(self):
        if not self.file_paths:
            messagebox.showerror("Error", "No files to combine.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".docx", 
            filetypes=[("Word files", "*.docx")],
            title="Save Combined Document"
        )
        if not save_path:
            return

        try:
            # Start with the first document
            master = Document(self.file_paths[0])
            composer = Composer(master)

            for file_path in self.file_paths[1:]:
                doc = Document(file_path)
                composer.append(doc)

            composer.save(save_path)
            messagebox.showinfo("Success", f"Combined file saved to:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to combine documents:\n{str(e)}")

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = WordMergerApp(root)
    root.mainloop()
