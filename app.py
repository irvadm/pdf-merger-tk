import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

from pypdf import PdfWriter

from datetime import datetime
from pathlib import Path
import subprocess
import sys

BUTTON_PADY = 5
BUTTON_IPADX = 10
BUTTON_IPADY = 10
LEFT_RIGHT_OUTER_PADDING = 20

class File:
    def __init__(self, path: str, parent_widget: tk.Frame, on_click_callback):
        self.path = path
        self.state = tk.BooleanVar(value=True) # Checked by default
        self.style = ttk.Style()
        self.style.map(
            'Custom.TCheckbutton',
            foreground=[('!selected', '#949494')]
        )
        self.checkbutton = ttk.Checkbutton(
            parent_widget,
            text=Path(self.path).name,
            variable=self.state,
            command=self.on_click,
            style='Custom.TCheckbutton'
        )
        self.on_click_callback = on_click_callback

    def __repr__(self):
        return f'File: {Path(self.path).name}'

    def on_click(self):
        print(Path(self.path).name, 'clicked')
        self.on_click_callback(self)


class PDFMergerApp:
    def __init__(self, root: tk.Tk):
        self.all_files = []
        self.selected_files = []
        self.delected_files = []
        self.width = 500
        self.height = 400
        self.root = root
        self.root.title('PDF Merger')
        self.root.minsize(self.width, self.height)
        self.root.columnconfigure(0, weight=1)
        self.root.columnconfigure(1, weight=1)

        self.frame = ttk.Frame(self.root)
        self.frame.columnconfigure(0, weight=1)
        self.frame.columnconfigure(1, weight=1)
        self.frame.grid(column=0, row=0, columnspan=2, padx=10, pady=10, sticky='nsew')

        self.center_window()
        self.create_widgets()

        # Bind events
        self.root.bind('<Escape>', sys.exit)

    def center_window(self):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (self.width // 2)
        y = (screen_height // 2) - (self.height // 2)
        self.root.geometry(f'{self.width}x{self.height}+{x}+{y}')

    def create_widgets(self):
        # Create 'PDF Merger' header
        header_style = ttk.Style()
        header_style.configure('Header.TLabel', foreground='#ab4e5b', font=('Calibri', 42, 'bold'))
        
        self.header = ttk.Label(self.frame, text='PDF Merger', anchor='center', style='Header.TLabel')
        self.header.grid(column=0, row=0, columnspan=2, pady=30, sticky='ew')
        
        # Create 'Select two or more files' button
        self.select_button = ttk.Button(self.frame, text='Select two or more files', command=self.select_files)
        self.select_button.grid(column=0, row=1, columnspan=2, padx=LEFT_RIGHT_OUTER_PADDING, pady=BUTTON_PADY, ipadx=BUTTON_IPADX, ipady=BUTTON_IPADY, sticky='ew')

        # Create 'Clear' button
        self.clear_button = ttk.Button(self.frame, text='Clear', command=self.clear_files)
        self.clear_button.grid(column=0, row=2, padx=LEFT_RIGHT_OUTER_PADDING, pady=BUTTON_PADY, ipadx=BUTTON_IPADX, ipady=BUTTON_IPADY, sticky='ew')

        # Create 'Merge' button
        self.merge_button = ttk.Button(self.frame, text='Merge', command=self.merge)
        self.merge_button.grid(column=1, row=2, padx=LEFT_RIGHT_OUTER_PADDING, pady=BUTTON_PADY, ipadx=BUTTON_IPADX, ipady=BUTTON_IPADY, sticky='ew')
        
        self.checkbutton_frame = ttk.Frame(self.frame)
        self.checkbutton_frame.grid(column=0, row=3, columnspan=2, padx=LEFT_RIGHT_OUTER_PADDING, sticky='ew')

        self.message_label = ttk.Label(self.frame)

    def clear_files(self):
        """Destroys File checkbutton widgets, and resets file related attributes."""
        for f in self.all_files:
            f.checkbutton.destroy()
        self.message_label['text'] = ''
        self.all_files = []
        self.selected_files = []
        self.delected_files = []

    def select_files(self):
        """Selected files reset when the user clicks to browse files."""
        self.clear_files()
        file_paths = filedialog.askopenfilenames(
            initialdir=Path.home(),
            title='Select two or more files',
            filetypes=[('PDF Files', '*.pdf')]
        )

        for file_path in file_paths:
            # Create checkbuttons
            f = File(
                file_path,
                self.checkbutton_frame,
                on_click_callback=self.checkbutton_handler
            )
            self.all_files.append(f)
            self.selected_files.append(f)
        self.update_files()
    
    def update_files(self):
        """Draw checkbutton widgets for selected and deselected files."""
        for i, f in enumerate(self.all_files):
            f.checkbutton.grid(column=0, row=i, columnspan=2, sticky='w')
    
    def checkbutton_handler(self, file: File):
        """
        File is added to the unselected files list if the user deselects it. Otherwise,
        File is re-added back onto the selected files list if the user reselects it.
        """
        # User unselects file
        if not file.state.get():
            file_index = self.selected_files.index(file)
            # Remove file from selected_files, and add it to unselected_files
            self.delected_files.append(self.selected_files.pop(file_index))
        # User reselects file
        else:
            # Remove file from deselected_files, and add it to selected_files
            file_index = self.delected_files.index(file)
            self.selected_files.append(self.delected_files.pop(file_index))
        self.update_files()
        print(f'SELECTED FILES: {self.selected_files}')
        print(f'DESELECTED FILES: {self.delected_files}')
        print(f'ALL FILES: {self.all_files}')

    def merge(self):
        if self.all_files == []:
            return
        # Open a Save As File Dialog, and ask user where they want to save the new file.
        save_as_path = filedialog.asksaveasfilename(
            initialdir=Path.home(),
            title='Save Location of Merged PDF',
            filetypes=[('PDF Files', '*.pdf')],
            initialfile=f'merged_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.pdf'
        )
        save_as_path = Path(save_as_path)

        # Merge PDFs
        print(f'Merging and saving new PDF file to: {save_as_path}')
        merger = PdfWriter()
        for selected_file in self.selected_files:
            merger.append(selected_file.path)
        merger.write(save_as_path)
        merger.close()

        # Ask if the user wants to open the newly created, merged PDF
        response = messagebox.askyesno('Merge Complete', 'Would you like to open your file?')
        if response:
            subprocess.run(['start', str(save_as_path)], shell=True)

        self.clear_files()

        # Display message with the save location of the new file.
        self.message_label['text'] = f'Saved PDF file at:\n{str(save_as_path)}'
        self.message_label.grid(column=0, row=4, columnspan=2, padx=LEFT_RIGHT_OUTER_PADDING, sticky='ew')


if __name__ == '__main__':
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
