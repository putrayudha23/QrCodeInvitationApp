import tkinter as tk
from tkinter import ttk, filedialog
import customtkinter
import sqlite3 as sq
from PIL import Image, ImageTk
import random
import pyqrcode
from pyqrcode import create
from xlsxwriter.workbook import Workbook
import pandas as pd
import os
import glob
from datetime import datetime
from threading import Timer
from tkinter import messagebox

customtkinter.set_appearance_mode("Light")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(customtkinter.CTk):
    # UI
    def __init__(self):
        super().__init__()

        # connect local database (sqlite3)
        self.database()

        self.title("QR Generator and Scanner")
        self.geometry("1300x700")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # add some style for treeview
        style = ttk.Style()
        style.theme_use("winnative")
        # configure our treeview colors
        style.configure("Treeview",
            background = "white",
            foreground = "black",
            rowheight = 30,
            fieldbackground = "white"
        )

        self.qrsample_image = customtkinter.CTkImage(Image.open("./image/sampleqr.png"), size=(250, 250))
        self.search_image = customtkinter.CTkImage(Image.open("./image/search.png"))
        self.home_image = Image.open("./image/HOME.png")
        self.success_image = Image.open("./image/SUCCESS.png")
        self.unregistered_image = Image.open("./image/unregistered2.png")

        self.tabview_qr = customtkinter.CTkTabview(self,fg_color=("gray93"),text_color="White",segmented_button_fg_color="Dark Blue",segmented_button_unselected_color="Dark Blue", segmented_button_selected_hover_color="blue", segmented_button_unselected_hover_color="blue",command=self.tabview_scan)
        self.tabview_qr.grid(row=0, column=0, padx=(0, 0), pady=(0, 0), sticky="nsew")
        self.tabview_qr.add("QR Code Generator")
        self.tabview_qr.add("QR Code Reader")
        # self.tabview_qr.add("Admin Email Setting")

        #>> QR Code Generator
        self.tabview_qr.tab("QR Code Generator").grid_rowconfigure(0, weight=1)
        self.tabview_qr.tab("QR Code Generator").grid_columnconfigure((0,1), weight=1)
        #LEFT ========================================================================================================================================================
        self.qr_generator_left_frame = customtkinter.CTkFrame(self.tabview_qr.tab("QR Code Generator"), corner_radius=10, fg_color="transparent")
        self.qr_generator_left_frame.grid(row=0, column=0, sticky="nsew", padx=(0,5), pady=5)
        self.qr_generator_left_frame.grid_columnconfigure(0, weight=1)
        self.qr_generator_left_frame.grid_rowconfigure(2, weight=1)
        # title
        self.title_frame = customtkinter.CTkFrame(self.qr_generator_left_frame, corner_radius=0, fg_color="transparent")
        self.title_frame.grid(row=0, column=0, sticky="nsew", padx=5)
        self.title_label = customtkinter.CTkLabel(self.title_frame, text="Invitation List", font=customtkinter.CTkFont(size=17,weight="bold"),anchor="w")
        self.title_label.pack(side="left")
        # search and live count
        self.search_livecount_frame = customtkinter.CTkFrame(self.qr_generator_left_frame, corner_radius=0, fg_color="transparent")
        self.search_livecount_frame.grid(row=1, column=0, sticky="nsew")
        self.search_livecount_frame.grid_columnconfigure((0,1), weight=1)
        self.search_livecount_frame.grid_rowconfigure(0, weight=1)
        self.search_frame = customtkinter.CTkFrame(self.search_livecount_frame, corner_radius=0, fg_color="transparent")
        self.search_frame.grid(row=0, column=0, sticky="nsew")
        self.search_label = customtkinter.CTkLabel(self.search_frame, text="Search :", font=customtkinter.CTkFont(size=15),anchor="w")
        self.search_label.pack(side="left", padx=(5,10), pady=10)
        self.entry_search = customtkinter.CTkEntry(self.search_frame, placeholder_text="", width=250)
        self.entry_search.pack(side="left")
        self.search_button = customtkinter.CTkButton(self.search_frame, fg_color="transparent", image=self.search_image, text="", hover_color="gray80", width=20, command=self.search)
        self.search_button.pack(side="left")
        self.livecount_frame = customtkinter.CTkFrame(self.search_livecount_frame, corner_radius=0, fg_color="transparent")
        self.livecount_frame.grid(row=0, column=1, sticky="nsew")
        self.livecount_label = customtkinter.CTkLabel(self.livecount_frame, text="Jumlah Kehadiran: 0", font=customtkinter.CTkFont(size=17,weight="bold"),anchor="w")
        self.livecount_label.pack(side="right", padx=10)
        # tabel
        self.tabel_frame = customtkinter.CTkFrame(self.qr_generator_left_frame, corner_radius=0, fg_color="white")
        self.tabel_frame.grid(row=2, column=0, sticky="nsew")
        style.map("Treeview",background=[('selected',"dark blue")])
        style.configure("Treeview.Heading", font=(None,12))
        self.invitation_tabel = ttk.Treeview(self.tabel_frame)
        self.invitation_tabel.pack(side="left", expand=tk.YES, fill=tk.BOTH)
        self.invitation_tabel["columns"] = ("ID System", "Nama", "Email", "No Hp", "Kuota", "Table Number", "Kehadiran", "Waktu", "oid")
        self.invitation_tabel.column("#0", width=0,stretch=tk.NO)
        self.invitation_tabel.column("ID System", anchor=tk.CENTER, width=10, minwidth=0,stretch=tk.YES)
        self.invitation_tabel.column("Nama", anchor="w", width=250,minwidth=30,stretch=tk.YES)
        self.invitation_tabel.column("Email", anchor=tk.CENTER, width=10, minwidth=0,stretch=tk.YES)
        self.invitation_tabel.column("No Hp", anchor=tk.CENTER, width=10, minwidth=0,stretch=tk.YES)
        self.invitation_tabel.column("Kuota", anchor=tk.CENTER, width=5, minwidth=0,stretch=tk.YES)
        self.invitation_tabel.column("Table Number", anchor=tk.CENTER, width=10, minwidth=0,stretch=tk.YES)
        self.invitation_tabel.column("Kehadiran", anchor=tk.CENTER, width=10, minwidth=0,stretch=tk.YES)
        self.invitation_tabel.column("Waktu", anchor=tk.CENTER, width=10, minwidth=0,stretch=tk.YES)
        self.invitation_tabel.column("oid", width=0,stretch=tk.NO)
        self.invitation_tabel.heading("#0", text="", anchor=tk.W)
        self.invitation_tabel.heading("ID System", text="ID System", anchor=tk.CENTER)
        self.invitation_tabel.heading("Nama", text="Nama", anchor=tk.CENTER)
        self.invitation_tabel.heading("Email", text="Email", anchor=tk.CENTER)
        self.invitation_tabel.heading("No Hp", text="No Hp", anchor=tk.CENTER)
        self.invitation_tabel.heading("Kuota", text="Kuota", anchor=tk.CENTER)
        self.invitation_tabel.heading("Table Number", text="Table Number", anchor=tk.CENTER)
        self.invitation_tabel.heading("Kehadiran", text="Kehadiran", anchor=tk.CENTER)
        self.invitation_tabel.heading("Waktu", text="Waktu", anchor=tk.CENTER)
        self.invitation_tabel.heading("oid", text="", anchor=tk.W)
        self.invitation_tabel.tag_configure("oddrow",background="white",font=(None, 13))
        self.invitation_tabel.tag_configure("evenrow",background="lightblue",font=(None, 13))
        self.warehouse_tabel_scrollbar = customtkinter.CTkScrollbar(self.tabel_frame, hover=True, button_hover_color="dark blue", command=self.invitation_tabel.yview)
        self.warehouse_tabel_scrollbar.pack(side="left",fill=tk.Y)
        self.invitation_tabel.configure(yscrollcommand=self.warehouse_tabel_scrollbar.set)
        self.invitation_tabel.bind("<<TreeviewSelect>>", self.on_tree_invitation_select)
        # entry
        self.entry_frame = customtkinter.CTkFrame(self.qr_generator_left_frame, corner_radius=10, fg_color="gray70")
        self.entry_frame.grid(row=3, column=0, pady=(5,0))
        self.entry_nama_label = customtkinter.CTkLabel(self.entry_frame, text="Nama:", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_nama_label.grid(row=0,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_nama = customtkinter.CTkEntry(self.entry_frame, placeholder_text="")
        self.entry_nama.grid(row=0,column=1, sticky="ew")
        self.entry_email_label = customtkinter.CTkLabel(self.entry_frame, text="Email:", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_email_label.grid(row=1,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_email = customtkinter.CTkEntry(self.entry_frame, placeholder_text="")
        self.entry_email.grid(row=1,column=1, sticky="ew")
        self.entry_tabelNum_label = customtkinter.CTkLabel(self.entry_frame, text="Table Number:", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_tabelNum_label.grid(row=2,column=0, padx=10, pady= 5, sticky="ew")
        self.entry_tabelNum = customtkinter.CTkEntry(self.entry_frame, placeholder_text="")
        self.entry_tabelNum.grid(row=2,column=1, sticky="ew")
        self.entry_nohp_label = customtkinter.CTkLabel(self.entry_frame, text="No HP:", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_nohp_label.grid(row=0,column=2, padx=10, pady= 5, sticky="ew")
        self.entry_nohp = customtkinter.CTkEntry(self.entry_frame, placeholder_text="")
        self.entry_nohp.grid(row=0,column=3, sticky="ew")
        self.entry_kuota_label = customtkinter.CTkLabel(self.entry_frame, text="Kuota:", font=customtkinter.CTkFont(size=13),anchor="w")
        self.entry_kuota_label.grid(row=1,column=2, padx=10, pady= 5, sticky="ew")
        self.entry_kuota = customtkinter.CTkEntry(self.entry_frame, placeholder_text="")
        self.entry_kuota.grid(row=1,column=3, sticky="ew")

        self.cancel_button = customtkinter.CTkButton(self.entry_frame, fg_color="dark red", text="Cancel", hover_color="red", command=self.cancel_modify)
        self.cancel_button.grid(row=2,column=3, rowspan=2)

        self.button_importExcel = customtkinter.CTkButton(self.entry_frame, fg_color="dark green", text="Import\nExcel", hover_color="green", command=self.import_excel, height=50)
        self.button_importExcel.grid(row=0,column=4, rowspan=3, padx=10, pady=10, sticky="nswe")
        self.button_generateQR = customtkinter.CTkButton(self.entry_frame, fg_color="dark blue", text="Generate\nQR Code", hover_color="blue", command=self.generate_QR, height=50)
        self.button_generateQR.grid(row=0,column=5, rowspan=3, padx=(0,10), pady=10, sticky="nswe")
        # button
        self.button_frame = customtkinter.CTkFrame(self.qr_generator_left_frame, corner_radius=0, fg_color="transparent")
        self.button_frame.grid(row=4, column=0)
        self.add_button = customtkinter.CTkButton(self.button_frame, fg_color="dark blue", text="Add", hover_color="blue", command=self.add_invitation)
        self.add_button.grid(row=0,column=0, pady= (10,0))
        self.remove_button = customtkinter.CTkButton(self.button_frame, fg_color="dark blue", text="Delete", hover_color="blue", command=self.remove)
        self.remove_button.grid(row=0,column=1, padx=10, pady= (10,0))
        self.replace_button = customtkinter.CTkButton(self.button_frame, fg_color="dark blue", text="Modify", hover_color="blue", command=self.replace)
        self.replace_button.grid(row=0,column=2, pady= (10,0))
        self.manual_presence_button = customtkinter.CTkButton(self.button_frame, fg_color="dark green", text="Manual Presence", hover_color="green", command=self.manual_presence)
        self.manual_presence_button.grid(row=0,column=3, pady= (10,0), padx=(10,0))
        self.export_data_button = customtkinter.CTkButton(self.button_frame, fg_color="dark green", text="Export Data", hover_color="green", command=self.export_data)
        self.export_data_button.grid(row=0,column=4, pady= (10,0), padx=(10,0))
        # RIGHT ========================================================================================================================================================
        self.qr_generator_right_frame = customtkinter.CTkFrame(self.tabview_qr.tab("QR Code Generator"), corner_radius=10, fg_color="gray80")
        self.qr_generator_right_frame.grid(row=0, column=1, sticky="nsew")
        self.qr_generator_right_frame.grid_columnconfigure(0, weight=1)
        self.qr_generator_right_frame.grid_rowconfigure(0, weight=1)
        # qr code
        self.qr_frame = customtkinter.CTkFrame(self.qr_generator_right_frame, corner_radius=0, fg_color="transparent")
        self.qr_frame.grid(row=0, column=0)
        self.qr_label = customtkinter.CTkLabel(self.qr_frame,text="Nama", compound="top", pady=15, image=self.qrsample_image, font=customtkinter.CTkFont(size=10, weight="bold"))
        self.qr_label.grid(row=0, column=0, padx=10,pady=10)

        #>> QR Code Reader ================================================================================================================================================
        self.tabview_qr.tab("QR Code Reader").grid_rowconfigure(0, weight=1)
        self.tabview_qr.tab("QR Code Reader").grid_columnconfigure(0, weight=1)
        # label (image) (nama)
        self.label_scan_frame = customtkinter.CTkFrame(self.tabview_qr.tab("QR Code Reader"), corner_radius=10, fg_color="transparent")
        self.label_scan_frame.grid(row=0, column=0, padx=5, sticky="nsew")
        self.scan_label = customtkinter.CTkLabel(self.label_scan_frame,text="", compound="center", pady=15, font=customtkinter.CTkFont(size=35, weight="bold"), text_color="gray93")
        self.scan_label.pack(expand=tk.YES, fill=tk.BOTH)
        self.label_scan_frame.bind("<Configure>", self.resizer_home)
        # entry
        self.entry_ket_frame = customtkinter.CTkFrame(self.tabview_qr.tab("QR Code Reader"), corner_radius=10, fg_color="transparent")
        self.entry_ket_frame.grid(row=1, column=0, sticky="nsew", padx=5)
        self.entry_ket = customtkinter.CTkEntry(self.entry_ket_frame, placeholder_text="", width=1, height=1, fg_color="transparent", text_color="gray93", placeholder_text_color="gray93", corner_radius=0)
        self.entry_ket.pack()
        self.entry_ket.bind('<Return>', self.scan)

        #>> tampilkan data di invitation tabel
        # get data from database
        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        count = 0
        c.execute("SELECT *, oid FROM tb_invitation")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
            else:
                self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
            count+=1
        file.commit()
        file.close()

        # perbahrui live count
        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        c.execute("SELECT kuota from tb_invitation where kehadiran = "":kehadiran""",{"kehadiran":"Hadir"})
        live_count = c.fetchall()
        file.commit()
        file.close()
        jumlah = 0
        for i in live_count:
            jumlah = jumlah + int(i[0])
        self.livecount_label.configure(text="Jumlah Kehadiran: {}".format(jumlah))

    def database(self):
        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        # create tabel below
        # buat tabel packaging registration
        c.execute("""CREATE TABLE IF NOT EXISTS tb_invitation(
            id KEY text,
            nama text,
            email text,
            no_hp text,
            kuota text,
            tabel_number text,
            kehadiran text,
            waktu text
        )""")

    def search(self):
        query = self.entry_search.get()
        selections = []
        for child in self.invitation_tabel.get_children():
            if query.capitalize() in self.invitation_tabel.item(child)['values'] or query.lower() in self.invitation_tabel.item(child)['values'] or query.upper() in self.invitation_tabel.item(child)['values'] or query.lower() in self.invitation_tabel.item(child)['values'] or query.swapcase() in self.invitation_tabel.item(child)['values'] or query.title() in self.invitation_tabel.item(child)['values']:
                selections.append(child)
        self.invitation_tabel.selection_set(selections)
      
    def on_tree_invitation_select(self, event):
        global no

        self.entry_nama.configure(state="normal")

        self.entry_nama.delete(0, tk.END)
        self.entry_email.delete(0, tk.END)
        self.entry_nohp.delete(0, tk.END)
        self.entry_kuota.delete(0, tk.END)
        self.entry_tabelNum.delete(0, tk.END)

        # get data and set all entry with it
        try:
            selectedItem = self.invitation_tabel.selection()[0]
            id = self.invitation_tabel.item(selectedItem)['values'][0]
            nama = self.invitation_tabel.item(selectedItem)['values'][1]
            email = self.invitation_tabel.item(selectedItem)['values'][2]
            no_hp = self.invitation_tabel.item(selectedItem)['values'][3]
            kuota = self.invitation_tabel.item(selectedItem)['values'][4]
            tabel_number = self.invitation_tabel.item(selectedItem)['values'][5]
            no = self.invitation_tabel.item(selectedItem)['values'][8]
        except:
            pass

        # insert to entry
        try:
            self.entry_nama.insert(0, nama)
            self.entry_nama.configure(state="readonly")  # Set the entry to readonly state
            self.entry_email.insert(0, email)
            self.entry_nohp.insert(0, no_hp)
            self.entry_kuota.insert(0, kuota)
            self.entry_tabelNum.insert(0, tabel_number)
        except:
            pass

        # show qr code and name
        qr_img = customtkinter.CTkImage(Image.open('./QRCode/{}_{}.png'.format(id,nama)), size=(250, 250))
        self.qr_label.configure(text=nama, image=qr_img)

    def cancel_modify(self):
        self.entry_nama.configure(state="normal")

        self.entry_nama.delete(0, tk.END)
        self.entry_email.delete(0, tk.END)
        self.entry_nohp.delete(0, tk.END)
        self.entry_kuota.delete(0, tk.END)
        self.entry_tabelNum.delete(0, tk.END)

        query = self.entry_nama.get()
        selections = []
        for child in self.invitation_tabel.get_children():
            if query in self.invitation_tabel.item(child)['values']:
                selections.append(child)
        self.invitation_tabel.selection_remove(selections)

    def generate_QR(self):
        global my_img, my_qr
        # get data from database where oid
        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        c.execute("SELECT * from tb_invitation where oid = "":no""",{"no":no})
        rec = c.fetchall()
        qr_string = ""
        for i in rec:
            qr_string = qr_string + i[0] + ";" + i[1] + ";" + i[2] + ";" + i[3] + ";" + i[4]
            my_qr = pyqrcode.create(qr_string)
            my_qr.png('./QRCode/{}_{}.png'.format(i[0],i[1]), scale=15, module_color=[0, 0, 0, 128], background=(255, 255, 255, 255))
        file.commit()
        file.close()
        #generate qr code
        my_qr = pyqrcode.create(qr_string)
        my_qr = my_qr.xbm(scale=7)
        my_img = tk.BitmapImage(data=my_qr)
        # update qr and name label
        qr_item = qr_string.split(';')
        nama = qr_item[1]
        self.qr_label.configure(text=nama, image=my_img)

        messagebox.showinfo("Information", "Data QR Code sudah digenerate dan di simpan di folder QRCode")
    
    def add_invitation(self):
        Question = messagebox.askquestion("Add", "Anda yakin ingin menambahkan data?")
        if Question == "yes":

            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            # get data from input
            id = random.randint(00000,99999)
            nama = self.entry_nama.get()
            email = self.entry_email.get()
            no_hp = self.entry_nohp.get()
            kuota = self.entry_kuota.get()
            kehadiran = ""
            waktu = ""
            tabel_number =  self.entry_tabelNum.get()
            # write data to database
            c.execute("INSERT INTO tb_invitation VALUEs(:id,:nama,:email,:no_hp,:kuota,:tabel_number,:kehadiran,:waktu)",{
                "id": id,
                "nama": nama,
                "email": email,
                "no_hp": no_hp,
                "kuota": kuota,
                "tabel_number" : tabel_number,
                "kehadiran": kehadiran,
                "waktu" : waktu
            })
            file.commit()
            file.close()
            # hapus item di tabel GUI
            for record in self.invitation_tabel.get_children():
                self.invitation_tabel.delete(record)
            # get data from database
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            count = 0
            c.execute("SELECT *, oid FROM tb_invitation")
            rec = c.fetchall()
            # input to tabel GUI
            for i in rec:
                if count % 2 == 0:
                    self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
                else:
                    self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
                count+=1
            
            # generate qr code
            qr_string = ""
            # qr_string = qr_string + str(id) + ";" + nama + ";" + email + ";" + no_hp + ";" + kuota
            qr_string = qr_string + str(id) + ";" + nama
            my_qr = pyqrcode.create(qr_string)
            my_qr.png('./QRCode/{}_{}.png'.format(str(id),nama), scale=15, module_color=[0, 0, 0, 128], background=(255, 255, 255, 255))

            # show qr code and name
            qr_img = customtkinter.CTkImage(Image.open('./QRCode/{}_{}.png'.format(str(id),nama)), size=(250, 250))
            self.qr_label.configure(text=nama, image=qr_img)
            
            # reset input
            self.entry_nama.delete(0,tk.END)
            self.entry_email.delete(0,tk.END)
            self.entry_nohp.delete(0,tk.END)
            self.entry_kuota.delete(0,tk.END)
            self.entry_tabelNum.delete(0,tk.END)

            file.commit()
            file.close()

            messagebox.showinfo("Information", "Data QR Code sudah digenerate dan di simpan di folder QRCode")
        else:
            pass

    def remove(self):
        Question = messagebox.askquestion("Delete", "Anda yakin ingin menghapus data?")
        if Question == "yes":
            # hapus qr code
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            c.execute("SELECT nama from tb_invitation where oid = "":no""",{"no":no})
            nama = c.fetchone()
            c.execute("SELECT id from tb_invitation where oid = "":no""",{"no":no})
            id = c.fetchone()
            file.commit()
            file.close()
            os.remove('./QRCode/{}_{}.png'.format(id[0],nama[0]))
            # hapus data on database
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            c.execute("DELETE from tb_invitation where oid = "":no""",{"no":no})
            file.commit()
            file.close()
            # hapus item di tabel GUI
            for record in self.invitation_tabel.get_children():
                self.invitation_tabel.delete(record)
            # get data from database
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            count = 0
            c.execute("SELECT *, oid FROM tb_invitation")
            rec = c.fetchall()
            # input to tabel GUI
            for i in rec:
                if count % 2 == 0:
                    self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
                else:
                    self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
                count+=1
            
            # reset input
            self.entry_nama.delete(0,tk.END)
            self.entry_email.delete(0,tk.END)
            self.entry_nohp.delete(0,tk.END)
            self.entry_kuota.delete(0,tk.END)
            self.entry_tabelNum.delete(0,tk.END)

            file.commit()
            file.close()

            # perbahrui live count
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            c.execute("SELECT kuota from tb_invitation where kehadiran = "":kehadiran""",{"kehadiran":"Hadir"})
            live_count = c.fetchall()
            file.commit()
            file.close()
            jumlah = 0
            for i in live_count:
                jumlah = jumlah + int(i[0])
            self.livecount_label.configure(text="Jumlah Kehadiran: {}".format(jumlah))
        else:
            pass
        
    def replace(self):
        Question = messagebox.askquestion("Modify", "Anda yakin ingin memperbahrui data?")
        if Question == "yes":
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            # get data from input
            nama = self.entry_nama.get()
            email = self.entry_email.get()
            no_hp = self.entry_nohp.get()
            kuota = self.entry_kuota.get()
            tabel_number =  self.entry_tabelNum.get()
            # write -> update data to database
            c.execute("UPDATE tb_invitation SET nama=:nama, email=:email, no_hp=:no_hp, kuota=:kuota, tabel_number=:tabel_number WHERE oid = "":no""",{
                "nama": nama,
                "email": email,
                "no_hp": no_hp,
                "kuota": kuota,
                "tabel_number" : tabel_number,
                "no":no
            })
            file.commit()
            file.close()
            # hapus item di tabel GUI
            for record in self.invitation_tabel.get_children():
                self.invitation_tabel.delete(record)
            # get data from database
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            count = 0
            c.execute("SELECT *, oid FROM tb_invitation")
            rec = c.fetchall()
            # input to tabel GUI
            for i in rec:
                if count % 2 == 0:
                    self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
                else:
                    self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
                count+=1

            # # generate qr code
            # file = sq.connect("./dataBase/dataBase.db")
            # c = file.cursor()
            # c.execute("SELECT id from tb_invitation where oid = "":no""",{"no":no})
            # id = c.fetchone()
            # file.commit()
            # file.close()
            # qr_string = ""
            # qr_string = qr_string + id[0] + ";" + nama + ";" + email + ";" + no_hp + ";" + kuota
            # my_qr = pyqrcode.create(qr_string)
            # my_qr.png('./QRCode/{}_{}.png'.format(id[0],nama), scale=15, module_color=[0, 0, 0, 128], background=(255, 255, 255, 255))
            
            # reset input
            self.entry_nama.delete(0,tk.END)
            self.entry_email.delete(0,tk.END)
            self.entry_nohp.delete(0,tk.END)
            self.entry_kuota.delete(0,tk.END)
            self.entry_tabelNum.delete(0,tk.END)

            # show qr code and name
            qr_img = customtkinter.CTkImage(Image.open('./QRCode/{}_{}.png'.format(id[0],nama)), size=(250, 250))
            self.qr_label.configure(text=nama, image=qr_img)

            # perbahrui live count
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            c.execute("SELECT kuota from tb_invitation where kehadiran = "":kehadiran""",{"kehadiran":"Hadir"})
            live_count = c.fetchall()
            file.commit()
            file.close()
            jumlah = 0
            for i in live_count:
                jumlah = jumlah + int(i[0])
            self.livecount_label.configure(text="Jumlah Kehadiran: {}".format(jumlah))
        else:
            pass

    def manual_presence(self):
        Question = messagebox.askquestion("Manual Presence", "Anda yakin ingin menandai 'Hadir'?")
        if Question == "yes":

            now = datetime.now()
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            # get data from input
            kehadiran = "Hadir"
            waktu = now.strftime("%H:%M:%S")
            # write -> update data to database
            c.execute("UPDATE tb_invitation SET kehadiran=:kehadiran, waktu=:waktu WHERE oid = "":no""",{
                "kehadiran" : kehadiran,
                "waktu" : waktu,
                "no":no
            })
            file.commit()
            file.close()
            # hapus item di tabel GUI
            for record in self.invitation_tabel.get_children():
                self.invitation_tabel.delete(record)
            # get data from database
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            count = 0
            c.execute("SELECT *, oid FROM tb_invitation")
            rec = c.fetchall()
            # input to tabel GUI
            for i in rec:
                if count % 2 == 0:
                    self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
                else:
                    self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
                count+=1
            
            # reset input
            self.entry_nama.delete(0,tk.END)
            self.entry_email.delete(0,tk.END)
            self.entry_nohp.delete(0,tk.END)
            self.entry_kuota.delete(0,tk.END)
            self.entry_tabelNum.delete(0,tk.END)

            file.commit()
            file.close()

            # perbahrui live count
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            c.execute("SELECT kuota from tb_invitation where kehadiran = "":kehadiran""",{"kehadiran":"Hadir"})
            live_count = c.fetchall()
            file.commit()
            file.close()
            jumlah = 0
            for i in live_count:
                jumlah = jumlah + int(i[0])
            self.livecount_label.configure(text="Jumlah Kehadiran: {}".format(jumlah))
        else:
            pass

    def export_data(self):
        workbook = Workbook('./report/report.xlsx')
        worksheet = workbook.add_worksheet()
        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        mysel = c.execute("SELECT * FROM tb_invitation")
        
        header_data = ['ID System', 'Nama', 'Email', 'No HP', 'Kuota', 'Table Number', 'Kehadiran', 'Waktu']
        header_format = workbook.add_format({'bold': True, 'bottom': 2, 'bg_color': '#F9DA04'})

        for col_num, data in enumerate(header_data):
            worksheet.write(0, col_num, data, header_format)

        for i, row in enumerate(mysel):
            for j, value in enumerate(row):
                worksheet.write(i+1, j, value)
                
        workbook.close()

        messagebox.showinfo("Information", "Data report sudah digenerate dan di simpan di folder report")

    def import_excel(self):
        Question = messagebox.askquestion("Import Excel", "Anda akan menghapus seluruh data.\nData yang terhapus akan diperbarui dengan data pada file excel.\nAnda Yakin?")
        if Question == "yes":

            filename = filedialog.askopenfilename(
                initialdir="C:/",
                title = "Open A File",
                filetypes=(("xlsx files", "*.xlsx"),("All Files", "*,*"))
            )

            if filename:
                try:
                    filename = r"{}".format(filename)
                    df = pd.read_excel(filename)
                except ValueError:
                    pass
                except FileNotFoundError:
                    pass

            # please wait
            pleaseWait_img = customtkinter.CTkImage(Image.open('./image/please_wait.png'), size=(250, 250))
            self.qr_label.configure(text="On Process Please Wait", image=pleaseWait_img)
            window = tk.Toplevel(None)
            window.geometry("0x0")
            self.wait_visibility(window)
            window.destroy()

            # delete all in database
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            c.execute("DELETE from tb_invitation")
            file.commit()
            file.close()

            # hapus all item in qr code folder
            files = glob.glob('./QRCode/*')
            for f in files:
                os.remove(f)

            # generate id and import to database
            df_rows = df.to_numpy().tolist()
            for row in df_rows:
                file = sq.connect("./dataBase/dataBase.db")
                c = file.cursor()

                id = random.randint(00000,99999)
                nama = row[0]
                email = row[1]
                no_hp = row[2]
                kuota = row[3]
                tabel_number = row[4]
                kehadiran = ""
                waktu = ""

                c.execute("INSERT INTO tb_invitation VALUEs(:id,:nama,:email,:no_hp,:kuota,:tabel_number,:kehadiran,:waktu)",{
                    "id": id,
                    "nama": nama,
                    "email": email,
                    "no_hp": no_hp,
                    "kuota": kuota,
                    "tabel_number" : tabel_number,
                    "kehadiran": kehadiran,
                    "waktu": waktu
                })
                file.commit()
                file.close()
                # generate qr code tiap item di database and save to qr code folder
                file = sq.connect("./dataBase/dataBase.db")
                c = file.cursor()
                c.execute("SELECT * from tb_invitation")
                rec = c.fetchall()
                strQr = ""
                for i in rec:
                    strQr = ""
                    # strQr = strQr + i[0] + ";" + i[1] + ";" + i[2] + ";" + i[3] + ";" + i[4]
                    strQr = strQr + i[0] + ";" + i[1]
                    my_qr = pyqrcode.create(strQr)
                    my_qr.png('./QRCode/{}_{}.png'.format(i[0],i[1]), scale=15, module_color=[0, 0, 0, 128], background=(255, 255, 255, 255))
                file.commit()
                file.close()
            # delete all in tree
            for record in self.invitation_tabel.get_children():
                self.invitation_tabel.delete(record)
            # insert to tree and show to label (qr and name)
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            count = 0
            c.execute("SELECT *, oid FROM tb_invitation")
            rec = c.fetchall()
            # input to tabel GUI
            for i in rec:
                if count % 2 == 0:
                    self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
                else:
                    self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
                count+=1
            file.commit()
            file.close()

            # please wait
            complete_img = customtkinter.CTkImage(Image.open('./image/complete.png'), size=(250, 250))
            self.qr_label.configure(text="Complete", image=complete_img)

            messagebox.showinfo("Information", "Data QR Code sudah digenerate dan di simpan di folder QRCode")
        else:
            pass

    def resizer_home(self,e):
        global resize_home, dynamic_width, dynamic_height
        home_image = Image.open("./image/HOME.png")
        resize_home = home_image.resize((e.width, e.height), Image.Resampling.LANCZOS)
        dynamic_width = e.width
        dynamic_height = e.height
        resize_home = ImageTk.PhotoImage(resize_home)
        self.scan_label.configure(image=resize_home)

    def scan(self, event):
        now = datetime.now()
        # get string
        scan_string = self.entry_ket.get()
        # delete entry
        self.entry_ket.delete(0,tk.END)
        # parsing string
        scan_string = scan_string.split(';')
        # get id
        key =scan_string[0]
        # try update kehadiran n waktu based on id -> label configure -> hitung waktu -> label configure
        # get data from input
        kehadiran = "Hadir"
        waktu = now.strftime("%H:%M:%S")
        try:
            file = sq.connect("./dataBase/dataBase.db")
            c = file.cursor()
            c.execute("SELECT nama, tabel_number, kuota from tb_invitation where id = "":id""",{"id":key})
            rec = c.fetchone()
            nama = rec[0]
            tabel_number = rec[1]
            kuota = rec[2]
            file.commit()
            file.close()
        except:
            pass
        # write -> update data to database
        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        try:
            c.execute("UPDATE tb_invitation SET kehadiran=:kehadiran, waktu=:waktu WHERE id = "":id""",{
                "kehadiran" : kehadiran,
                "waktu" : waktu,
                "id": key
            })
            # label configure
            resize_success = self.success_image.resize((dynamic_width, dynamic_height), Image.Resampling.LANCZOS)
            resize_success = ImageTk.PhotoImage(resize_success)
            self.scan_label.configure(text="{}\nTable Number: {} ({} PAX)".format(nama,tabel_number,kuota), image=resize_success, text_color="white")
            # timer
            t = Timer(5.0, self.initial_scan)
            t.start()
        except:
            # label configure
            resize_unregistered = self.unregistered_image.resize((dynamic_width, dynamic_height), Image.Resampling.LANCZOS)
            resize_unregistered = ImageTk.PhotoImage(resize_unregistered)
            self.scan_label.configure(image=resize_unregistered)
            # timer
            t = Timer(5.0, self.initial_scan)
            t.start()
        file.commit()
        file.close()
        # hapus item di tabel GUI
        for record in self.invitation_tabel.get_children():
            self.invitation_tabel.delete(record)
        # get data from database
        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        count = 0
        c.execute("SELECT *, oid FROM tb_invitation")
        rec = c.fetchall()
        # input to tabel GUI
        for i in rec:
            if count % 2 == 0:
                self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("evenrow",))
            else:
                self.invitation_tabel.insert(parent="",index="end", iid=count, text="", values=(i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]), tags=("oddrow",))
            count+=1

        # perbahrui live count
        file = sq.connect("./dataBase/dataBase.db")
        c = file.cursor()
        c.execute("SELECT kuota from tb_invitation where kehadiran = "":kehadiran""",{"kehadiran":"Hadir"})
        live_count = c.fetchall()
        file.commit()
        file.close()
        jumlah = 0
        for i in live_count:
            jumlah = jumlah + int(i[0])
        self.livecount_label.configure(text="Jumlah Kehadiran: {}".format(jumlah))

    def initial_scan(self):
        self.scan_label.configure(text_color="gray93")
        self.scan_label.configure(text="",image=resize_home, text_color="gray93")

    def tabview_scan(self):
        tab_name = self.tabview_qr.get()
        if tab_name == "QR Code Reader":
            self.entry_ket.focus()

if __name__ == "__main__":
    app = App()
    app.mainloop()