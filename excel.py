import win32com.client as win32
import tkinter as tk
from tkinter import filedialog

class TextEditor:
    def __init__(self, master):
        self.master = master
        master.title("Text Editor")

        # Create an Excel application
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = True
        self.excel.Workbooks.Add()

        # Create a menu
        self.menu = tk.Menu(master)
        self.file_menu = tk.Menu(self.menu, tearoff=0)
        self.file_menu.add_command(label="New", command=self.new_file)
        self.file_menu.add_command(label="Open", command=self.open_file)
        self.file_menu.add_command(label="Save", command=self.save_file)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Exit", command=self.exit_program)
        self.menu.add_cascade(label="File", menu=self.file_menu)
        master.config(menu=self.menu)

    def new_file(self):
        # Clear the contents of the active sheet
        self.excel.ActiveSheet.Cells.Clear()

    def open_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.excel.Workbooks.Open(file_path)

    def save_file(self):
        file_path = filedialog.asksaveasfilename()
        if file_path:
            self.excel.ActiveWorkbook.SaveAs(file_path)

    def exit_program(self):
        # Close the Excel application and exit the program
        self.excel.Quit()
        self.master.quit()

root = tk.Tk()
app = TextEditor(root)
root.mainloop()