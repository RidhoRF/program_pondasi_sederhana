import os
import sys
from  subprocess import call
import tkinter as tk
from PIL import ImageTk, Image
from openpyxl import load_workbook
from tkinter import messagebox, ttk


# ONE FILE TEMP FILE
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

"""program = input("Pilih program (1 atau 2): ")

if program == "1":
    # Jalankan program 1
    call(["python", "Program_Pondasi1.py"])
elif program == "2":
    # Jalankan program 2
    call(["python", "Program_Pondasi2.py"])
else:
    print("Pilihan tidak valid.")"""


class App(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Program Pondasi")
        self.create_widgets()

    def create_widgets(self):

        logo_path = resource_path("logo.png")
        self.logoits = ImageTk.PhotoImage(Image.open(logo_path).resize((140, 50), Image.BILINEAR))
        tk.Label(self.master, image=self.logoits).place(x=10, y=10)
        tk.Label(self.master, text=" ").pack(side="top")
        tk.Label(self.master, text=" ").pack(side="top")
        tk.Label(self.master, text=" ").pack(side="top")
        tk.Label(self.master, text=" ").pack(side="top")
        

        # Input Box
        tk.Label(self.master, text="Program Daya Dukung Pondasi Dangkal").pack()
        tk.Button(self.master, text="Buka Program Daya Dukung", command=self.o1).pack()

        tk.Label(self.master, text=" ").pack(side="top")

        tk.Label(self.master, text="Program Tegangn Dasar Pondasi Dangkal").pack()
        tk.Button(self.master, text="Buka Program Tegangan", command=self.o2).pack()


        # Calculate Button
        tk.Label(self.master, text=" ").pack(side="top")
        tk.Button(self.master, text="Info", command=self.calculate).pack()

    def o1(self):
        call(["python", "Program_Pondasi1.py"])
        
    def o2(self):
        call(["python", "Program_Pondasi2.py"])

    def calculate(self):
        
        # Tampilkan pesan berhasil
        tk.messagebox.showinfo("Info", "Program ini dibuat oleh :\n\n   1. Ridho Rizky Febriansyah 2035221071 (Coder)\n\n   Made with Python 3.11, GUI Tkinter")
   
ico_path = resource_path("ico.ico")

if __name__ == "__main__":
    root = tk.Tk()
    root.iconbitmap(ico_path)
    root.geometry("420x300")
    root.resizable(False, False)
    app = App(master=root)
    app.mainloop()

