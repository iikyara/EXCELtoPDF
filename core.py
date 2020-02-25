# coding: utf-8
import sys
import os
import shutil
import json
import csv
import win32com.client
import pyminizip
import pikepdf
from pikepdf import Pdf
import pyminizip
import tkinter as tk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
from tkinter import ttk
from utils import *

class App:
    def __init__(self):
        self.view = View()
        self.controller = Controller(self.view)
        #self.run()

    def run(self):
        self.view.run()

class View:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("EXCELtoPDF")
        self.window.geometry("700x480")
        self.window.iconbitmap(default='icon.ico')

        # info data
        self.info_data = tk.StringVar()

        # list data
        self.projectlist_data = tk.StringVar()
        self.namelist_data = tk.StringVar()
        self.pdflist_data = tk.StringVar()

        # combobox data
        self.sortlist = ['昇順', '降順']
        self.projectlistop_sort_data = tk.StringVar()
        self.namelistop_sort_data = tk.StringVar()
        self.pdflistop_sort_data = tk.StringVar()
        self.setting_sheet_data = tk.StringVar()

        # checkbox data
        self.setting_genpdf_data = tk.BooleanVar()

        # text box
        self.namebox = None
        self.edit_name = None
        self.edit_pdfpass = None
        self.edit_zippass = None
        self.edit_pdfname = None
        self.edit_zipname = None
        self.setting_input = None
        self.setting_offset_column = None
        self.setting_offset_row = None
        self.setting_range_column = None
        self.setting_range_row = None
        self.setting_output = None

        # button
        self.projectlistop_new = None
        self.projectlistop_trash = None
        self.addbutton = None
        self.namelistop_trash = None
        self.listop_add = None
        self.pdflistop_trash = None
        self.edit_save = None
        self.setting_input_button = None
        self.setting_sheet_button = None
        self.setting_output_button = None
        self.genpdf_button = None

        # list
        self.projectlist = None
        self.namelist = None
        self.pdflist = None

        # combobox
        self.namelistop_sort = None
        self.pdflistop_sort = None
        self.setting_sheet = None

        # listner func
        self.on_push_projectlistop_new = []
        self.on_push_projectlistop_trash = []
        self.on_push_addbutton = []
        self.on_push_namelistop_trash = []
        self.on_push_listop_add = []
        self.on_push_pdflistop_trash = []
        self.on_push_edit_save = []
        self.on_push_setting_input_button = []
        self.on_push_setting_sheet_button = []
        self.on_push_setting_output_button = []
        self.on_push_genpdf_button = []

        self.on_select_projectlist = []
        self.on_select_namelist = []
        self.on_select_pdflist = []

        self.on_change_projectlistop_sort = []
        self.on_change_namelistop_sort = []
        self.on_change_pdflistop_sort = []
        self.on_change_setting_sheet = []

        self.on_change_setting_genpdf = []

        self.on_close_window = []

        self.createGUI()

        # set listener
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)
        self.projectlistop_new['command'] = self.push_projectlistop_new
        self.projectlistop_trash['command'] = self.push_projectlistop_trash
        self.addbutton['command'] = self.push_addbutton
        self.namelistop_trash['command'] = self.push_namelistop_trash
        self.listop_add['command'] = self.push_listop_add
        self.pdflistop_trash['command'] = self.push_pdflistop_trash
        self.edit_save['command'] = self.push_edit_save
        self.setting_input_button['command'] = self.push_setting_input_button
        self.setting_sheet_button['command'] = self.push_setting_sheet_button
        self.setting_output_button['command'] = self.push_setting_output_button
        self.genpdf_button['command'] = self.push_genpdf_button


        self.projectlist.bind('<<ListboxSelect>>', self.select_projectlist)
        self.namelist.bind('<<ListboxSelect>>', self.select_namelist)
        self.pdflist.bind('<<ListboxSelect>>', self.select_pdflist)

        self.projectlistop_sort.bind('<<ComboboxSelected>>', self.change_projectlistop_sort)
        self.namelistop_sort.bind('<<ComboboxSelected>>', self.change_namelistop_sort)
        self.pdflistop_sort.bind('<<ComboboxSelected>>', self.change_pdflistop_sort)
        self.setting_sheet.bind('<<ComboboxSelected>>', self.change_setting_sheet)

        self.setting_genpdf['command'] = self.change_setting_genpdf

    def createGUI(self):
        # frames
        self.main_frame = ttk.Frame(self.window, style="Main.TFrame")
        self.project_frame = ttk.Frame(self.main_frame, style="A.TFrame")
        self.left_frame = ttk.Frame(self.main_frame, style="A.TFrame")
        self.right_frame = ttk.Frame(self.main_frame, style="B.TFrame")

        self.main_frame.grid(column=0, row=0, sticky=tk.NSEW)
        self.project_frame.grid(column=0, row=0, sticky=tk.NSEW, padx=1, pady=1)
        self.left_frame.grid(column=1, row=0, sticky=tk.NSEW, padx=1, pady=1)
        self.right_frame.grid(column=2, row=0, sticky=tk.NS, padx=1, pady=1)

        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(0, weight=1)
        self.project_frame.columnconfigure(0, weight=1)
        self.project_frame.rowconfigure(1, weight=1)
        self.left_frame.columnconfigure(0, weight=1)
        self.left_frame.rowconfigure(3, weight=1)
        self.left_frame.rowconfigure(7, weight=1)
        self.right_frame.columnconfigure(0, weight=1)
        self.right_frame.rowconfigure(0, weight=1)
        self.right_frame.rowconfigure(1, weight=1)
        self.right_frame.rowconfigure(2, weight=1)

        # info field
        self.info_label = ttk.Label(self.main_frame, textvariable=self.info_data)

        self.info_label.grid(column=0, row=1, columnspan=3, sticky=tk.EW)

        # project list
        self.projectlist_label = ttk.Label(self.project_frame, text='プロジェクトリスト')
        self.projectlist_frame = ttk.Frame(self.project_frame)
        self.projectlist = tk.Listbox(self.projectlist_frame, listvariable=self.projectlist_data)
        self.projectlist_scrollbar = ttk.Scrollbar(
            self.projectlist_frame,
            orient=tk.VERTICAL,
            command=self.projectlist.yview
        )
        self.projectlist["yscrollcommand"] = self.projectlist_scrollbar.set

        self.projectlist_label.grid(column=0, row=0, sticky=tk.W)
        self.projectlist_frame.grid(column=0, row=1, sticky=tk.NSEW)
        self.projectlist.grid(column=0, row=0, sticky=tk.NSEW)
        self.projectlist_scrollbar.grid(column=1, row=0, sticky=tk.NS)

        self.projectlist_frame.rowconfigure(0, weight=1)
        self.projectlist_frame.columnconfigure(0, weight=1)

        # project list operation
        self.projectlistop_frame = ttk.Frame(self.project_frame)
        # sort
        self.projectlistop_sort_label = ttk.Label(self.projectlistop_frame, text="並び順：")
        self.projectlistop_sort = ttk.Combobox(self.projectlistop_frame, textvariable=self.projectlistop_sort_data, values=self.sortlist, state='readonly', width=10)
        self.projectlistop_sort.current(0)
        # project operation
        self.projectlistop_new = ttk.Button(self.projectlistop_frame, text="新規作成", width=10)
        self.projectlistop_trash = ttk.Button(self.projectlistop_frame, text="削除", width=5)

        self.projectlistop_frame.grid(column=0, row=4, sticky=tk.EW)
        self.projectlistop_sort_label.grid(column=0, row=0, sticky=tk.W)
        self.projectlistop_sort.grid(column=1, row=0, sticky=tk.W)
        self.projectlistop_new.grid(column=2, row=0, sticky=tk.E)
        self.projectlistop_trash.grid(column=3, row=0, sticky=tk.E)

        self.projectlistop_frame.columnconfigure(2, weight=1)

        # name box
        self.namebox_frame = ttk.Frame(self.left_frame)
        self.namebox_label = ttk.Label(self.namebox_frame, text="名前入力：")
        self.namebox = ttk.Entry(self.namebox_frame)

        self.namebox_frame.grid(column=0, row=0, sticky=tk.EW)
        self.namebox_label.grid(column=0, row=0)
        self.namebox.grid(column=1, row=0, sticky=tk.EW)

        self.namebox_frame.columnconfigure(1, weight=1)

        # add button
        self.addbutton = ttk.Button(self.left_frame, text="↓ 追加 ↓")
        self.addbutton.grid(column=0, row=1, sticky=tk.EW)

        # name list
        self.namelist_label = ttk.Label(self.left_frame, text="名前リスト")
        self.namelist_frame = ttk.Frame(self.left_frame)
        self.namelist = tk.Listbox(self.namelist_frame, listvariable=self.namelist_data)
        self.namelist_scrollbar = ttk.Scrollbar(
            self.namelist_frame,
            orient=tk.VERTICAL,
            command=self.namelist.yview
        )
        self.namelist["yscrollcommand"] = self.namelist_scrollbar.set

        self.namelist_label.grid(column=0, row=2, sticky=tk.W)
        self.namelist_frame.grid(column=0, row=3, sticky=tk.NSEW)
        self.namelist.grid(column=0, row=0, sticky=tk.NSEW)
        self.namelist_scrollbar.grid(column=1, row=0, sticky=tk.NS)

        self.namelist_frame.rowconfigure(0, weight=1)
        self.namelist_frame.columnconfigure(0, weight=1)

        # name list operation
        self.namelistop_frame = ttk.Frame(self.left_frame)
        # sort
        self.namelistop_sort_label = ttk.Label(self.namelistop_frame, text="並び順：")
        self.namelistop_sort = ttk.Combobox(self.namelistop_frame, textvariable=self.namelistop_sort_data, values=self.sortlist, state='readonly', width=10)
        self.namelistop_sort.current(0)
        # namelist operation
        self.namelistop_trash = ttk.Button(self.namelistop_frame, text="削除", width=5)

        self.namelistop_frame.grid(column=0, row=4, sticky=tk.EW)
        self.namelistop_sort_label.grid(column=0, row=0, sticky=tk.W)
        self.namelistop_sort.grid(column=1, row=0, sticky=tk.W)
        self.namelistop_trash.grid(column=2, row=0, sticky=tk.E)

        self.namelistop_frame.columnconfigure(2, weight=1)

        # list operation
        self.listop_add = ttk.Button(self.left_frame, text="↓追加↓", width=10)
        self.listop_add.grid(column=0, row=5, sticky=tk.EW)

        # pdf list
        self.pdflist_label = ttk.Label(self.left_frame, text="PDF出力一覧")
        self.pdflist_frame = ttk.Frame(self.left_frame)
        self.pdflist = tk.Listbox(self.pdflist_frame, listvariable=self.pdflist_data)
        self.pdflist_scrollbar = ttk.Scrollbar(
            self.pdflist_frame,
            orient=tk.VERTICAL,
            command=self.pdflist.yview
        )
        self.pdflist["yscrollcommand"] = self.pdflist_scrollbar.set

        self.pdflist_label.grid(column=0, row=6, sticky=tk.W)
        self.pdflist_frame.grid(column=0, row=7, sticky=tk.NSEW)
        self.pdflist.grid(column=0, row=0, sticky=tk.NSEW)
        self.pdflist_scrollbar.grid(column=1, row=0, sticky=tk.NS)

        self.pdflist_frame.rowconfigure(0, weight=1)
        self.pdflist_frame.columnconfigure(0, weight=1)

        # pdf list operation
        self.pdflistop_frame = ttk.Frame(self.left_frame)
        # sort
        self.pdflistop_sort_label = ttk.Label(self.pdflistop_frame, text="並び順：")
        self.pdflistop_sort = ttk.Combobox(self.pdflistop_frame, textvariable=self.pdflistop_sort_data, values=self.sortlist, state='readonly', width=10)
        self.pdflistop_sort.current(0)
        # pdflist operation
        self.pdflistop_trash = ttk.Button(self.pdflistop_frame, text="削除", width=5)

        self.pdflistop_frame.grid(column=0, row=8, sticky=tk.EW)
        self.pdflistop_sort_label.grid(column=0, row=0, sticky=tk.W)
        self.pdflistop_sort.grid(column=1, row=0, sticky=tk.W)
        self.pdflistop_trash.grid(column=2, row=0, sticky=tk.E)

        self.pdflistop_frame.columnconfigure(2, weight=1)

        # edit info
        self.edit_frame = ttk.Frame(self.right_frame)
        self.edit_label = ttk.Label(self.edit_frame, text="編集")
        # edit name
        self.edit_name_label = ttk.Label(self.edit_frame, text="名前：")
        self.edit_name = ttk.Entry(self.edit_frame)
        # pdf pass
        self.edit_pdfpass_label = ttk.Label(self.edit_frame, text="PDFパスワード：")
        self.edit_pdfpass = ttk.Entry(self.edit_frame)
        # zip pass
        self.edit_zippass_label = ttk.Label(self.edit_frame, text="ZIPパスワード：")
        self.edit_zippass = ttk.Entry(self.edit_frame)
        # pdf filename
        self.edit_pdfname_label = ttk.Label(self.edit_frame, text="PDFファイル名：")
        self.edit_pdfname = ttk.Entry(self.edit_frame)
        # zip filename
        self.edit_zipname_label = ttk.Label(self.edit_frame, text="ZIPファイル名：")
        self.edit_zipname = ttk.Entry(self.edit_frame)
        # save button
        self.edit_save = ttk.Button(self.edit_frame, text="変更を保存")

        self.edit_frame.grid(column=0, row=0, sticky=tk.NSEW)
        self.edit_label.grid(column=0, row=0, sticky=tk.W)
        self.edit_name_label.grid(column=0, row=1, sticky=tk.E)
        self.edit_name.grid(column=1, row=1, sticky=tk.EW)
        self.edit_pdfpass_label.grid(column=0, row=2, sticky=tk.E)
        self.edit_pdfpass.grid(column=1, row=2, sticky=tk.EW)
        self.edit_zippass_label.grid(column=0, row=3, sticky=tk.E)
        self.edit_zippass.grid(column=1, row=3, sticky=tk.EW)
        self.edit_pdfname_label.grid(column=0, row=4, sticky=tk.E)
        self.edit_pdfname.grid(column=1, row=4, sticky=tk.EW)
        self.edit_zipname_label.grid(column=0, row=5, sticky=tk.E)
        self.edit_zipname.grid(column=1, row=5, sticky=tk.EW)
        self.edit_save.grid(column=1, row=6, sticky=tk.E)

        self.edit_frame.columnconfigure(1, weight=1)

        # setting
        self.setting_frame = ttk.Frame(self.right_frame)
        self.setting_lable = ttk.Label(self.setting_frame, text="設定")
        # input file
        self.setting_input_label = ttk.Label(self.setting_frame, text="入力ファイル：")
        self.setting_input_frame = ttk.Frame(self.setting_frame)
        self.setting_input = ttk.Entry(self.setting_input_frame)
        self.setting_input_button = ttk.Button(self.setting_input_frame, text="参照", width=5)
        # sheet select
        self.setting_sheet_label = ttk.Label(self.setting_frame, text="対象シート：")
        self.setting_sheet_frame = ttk.Frame(self.setting_frame)
        self.setting_sheet = ttk.Combobox(self.setting_sheet_frame, textvariable=self.setting_sheet_data, state='readonly')
        self.setting_sheet_button = ttk.Button(self.setting_sheet_frame, text="更新", width=5)
        # offset
        self.setting_offset_label = ttk.Label(self.setting_frame, text="ズレ（右×下）：")
        self.setting_offset_frame = ttk.Frame(self.setting_frame)
        self.setting_offset_column = ttk.Entry(self.setting_offset_frame, width=10)
        self.setting_offset_middle = ttk.Label(self.setting_offset_frame, text="×")
        self.setting_offset_row = ttk.Entry(self.setting_offset_frame, width=10)
        # range
        self.setting_range_label = ttk.Label(self.setting_frame, text="範囲（横×縦）：")
        self.setting_range_frame = ttk.Frame(self.setting_frame)
        self.setting_range_column = ttk.Entry(self.setting_range_frame, width=10)
        self.setting_range_middle = ttk.Label(self.setting_range_frame, text="×")
        self.setting_range_row = ttk.Entry(self.setting_range_frame, width=10)
        # output file
        self.setting_output_label = ttk.Label(self.setting_frame, text="出力場所：")
        self.setting_output_frame = ttk.Frame(self.setting_frame)
        self.setting_output = ttk.Entry(self.setting_output_frame)
        self.setting_output_button = ttk.Button(self.setting_output_frame, text="参照", width=5)
        # is pdf
        self.setting_genpdf_label = ttk.Label(self.setting_frame, text="PDF化：")
        self.setting_genpdf = ttk.Checkbutton(self.setting_frame, variable=self.setting_genpdf_data)

        self.setting_frame.grid(column=0, row=1, sticky=tk.NSEW)
        self.setting_lable.grid(column=0, row=0, sticky=tk.W)
        self.setting_input_label.grid(column=0, row=1, sticky=tk.E)
        self.setting_input_frame.grid(column=1, row=1, sticky=tk.EW)
        self.setting_input.grid(column=0, row=0, sticky=tk.EW)
        self.setting_input_button.grid(column=1, row=0, sticky=tk.E)
        self.setting_sheet_label.grid(column=0, row=2, sticky=tk.E)
        self.setting_sheet_frame.grid(column=1, row=2, sticky=tk.EW)
        self.setting_sheet.grid(column=0, row=0, sticky=tk.EW)
        self.setting_sheet_button.grid(column=1, row=0, sticky=tk.E)
        self.setting_offset_label.grid(column=0, row=3, sticky=tk.E)
        self.setting_offset_frame.grid(column=1, row=3, sticky=tk.EW)
        self.setting_offset_column.grid(column=0, row=0)
        self.setting_offset_middle.grid(column=1, row=0)
        self.setting_offset_row.grid(column=2, row=0)
        self.setting_range_label.grid(column=0, row=4, sticky=tk.E)
        self.setting_range_frame.grid(column=1, row=4, sticky=tk.EW)
        self.setting_range_column.grid(column=0, row=0)
        self.setting_range_middle.grid(column=1, row=0)
        self.setting_range_row.grid(column=2, row=0)
        self.setting_output_label.grid(column=0, row=5, sticky=tk.E)
        self.setting_output_frame.grid(column=1, row=5, sticky=tk.EW)
        self.setting_output.grid(column=0, row=0, sticky=tk.EW)
        self.setting_output_button.grid(column=1, row=0, sticky=tk.E)
        self.setting_genpdf_label.grid(column=0, row=6, sticky=tk.E)
        self.setting_genpdf.grid(column=1, row=6, sticky=tk.EW)

        self.setting_frame.columnconfigure(1, weight=1)
        self.setting_input_frame.columnconfigure(0, weight=1)
        self.setting_output_frame.columnconfigure(0, weight=1)

        # generate pdf button
        self.genpdf_button = ttk.Button(self.right_frame, text="PDFを出力する")
        self.genpdf_button.grid(column=0, row=2, sticky=tk.NSEW)

        ttk.Style().configure("Main.TFrame", padding=6, relief="flat",   background="#000")
        #ttk.Style().configure("A.TFrame", padding=6, relief="flat",   background="#F00")
        #ttk.Style().configure("B.TFrame", padding=6, relief="flat",  background="#0F0")

    def setEdit(self, name):
        self.edit_name.delete(0, tk.END)
        self.edit_name.insert(tk.END, name.name)
        self.edit_pdfpass.delete(0, tk.END)
        self.edit_pdfpass.insert(tk.END, name.pdf_password or "")
        self.edit_zippass.delete(0, tk.END)
        self.edit_zippass.insert(tk.END, name.zip_password or "")
        self.edit_pdfname.delete(0, tk.END)
        self.edit_pdfname.insert(tk.END, name.pdf_filename or "")
        self.edit_zipname.delete(0, tk.END)
        self.edit_zipname.insert(tk.END, name.zip_filename or "")

    def setInfo(self, info):
        self.info_data.set(info)
        self.window.update()

    def messageInfo(self, title, info):
        messagebox.showinfo(title, info)

    def enableGen(self):
        self.setGen(True)

    def disableGen(self):
        self.setGen(False)

    def setGen(self, is_gen):
        #text box
        self.namebox.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.edit_name.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.edit_pdfpass.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.edit_zippass.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.edit_pdfname.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.edit_zipname.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_input.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_offset_column.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_offset_row.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_range_column.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_range_row.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_output.configure(state=tk.DISABLED if is_gen else tk.NORMAL)

        #button
        self.projectlistop_new.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.projectlistop_trash.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.addbutton.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.namelistop_trash.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.listop_add.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.pdflistop_trash.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.edit_save.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_input_button.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_sheet_button.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_output_button.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.genpdf_button.configure(state=tk.DISABLED if is_gen else tk.NORMAL)

        #list
        self.projectlist.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.namelist.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.pdflist.configure(state=tk.DISABLED if is_gen else tk.NORMAL)

        #combobox
        self.projectlistop_sort.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.namelistop_sort.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.pdflistop_sort.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_sheet.configure(state=tk.DISABLED if is_gen else tk.NORMAL)

        #checkbox
        self.setting_genpdf.configure(state=tk.DISABLED if is_gen else tk.NORMAL)

        self.window.update()

    def push_projectlistop_new(self):
        for func in self.on_push_projectlistop_new:
            func()

    def push_projectlistop_trash(self):
        for func in self.on_push_projectlistop_trash:
            func()

    def push_addbutton(self):
        for func in self.on_push_addbutton:
            func()

    def push_namelistop_trash(self):
        for func in self.on_push_namelistop_trash:
            func()

    def push_listop_add(self):
        for func in self.on_push_listop_add:
            func()

    def push_pdflistop_trash(self):
        for func in self.on_push_pdflistop_trash:
            func()

    def push_edit_save(self):
        for func in self.on_push_edit_save:
            func()

    def push_setting_input_button(self):
        for func in self.on_push_setting_input_button:
            func()

    def push_setting_sheet_button(self):
        for func in self.on_push_setting_sheet_button:
            func()

    def push_setting_output_button(self):
        for func in self.on_push_setting_output_button:
            func()

    def push_genpdf_button(self):
        for func in self.on_push_genpdf_button:
            func()

    def select_projectlist(self, event):
        for func in self.on_select_projectlist:
            func(event)

    def select_namelist(self, event):
        for func in self.on_select_namelist:
            func(event)

    def select_pdflist(self, event):
        for func in self.on_select_pdflist:
            func(event)

    def change_projectlistop_sort(self, event):
        for func in self.on_change_projectlistop_sort:
            func(event)

    def change_namelistop_sort(self, event):
        for func in self.on_change_namelistop_sort:
            func(event)

    def change_pdflistop_sort(self, event):
        for func in self.on_change_pdflistop_sort:
            func(event)

    def change_setting_sheet(self, event):
        for func in self.on_change_setting_sheet:
            func(event)

    def change_setting_genpdf(self):
        for func in self.on_change_setting_genpdf:
            func()

    def close_window(self):
        for func in self.on_close_window:
            func()
        self.window.destroy()

    def getData(self):
        return {
            "project_list" : list(self.projectlist.get(0, self.projectlist.size() - 1)),
            "name_list" : list(self.namelist.get(0, self.namelist.size() - 1)),
            "pdf_list" : list(self.pdflist.get(0, self.pdflist.size() - 1)),
            "project_sort" : self.projectlistop_sort_data.get(),
            "name_sort" : self.namelistop_sort_data.get(),
            "pdf_sort" : self.pdflistop_sort_data.get(),
            "input_file" : self.setting_input.get(),
            "sheet" : self.setting_sheet_data.get(),
            "output_file" : self.setting_output.get(),
            "offset" : (
                tryParseInt(self.setting_offset_column.get(), default=0),
                tryParseInt(self.setting_offset_row.get(), default=0)
            ),
            "range" : (
                tryParseInt(self.setting_range_column.get(), default=1),
                tryParseInt(self.setting_range_row.get(), default=1)
            ),
            "genpdf" : self.setting_genpdf_data.get(),
            "project_list_selected" : self.projectlist.curselection(),
            "name_list_selected" : self.namelist.curselection(),
            "pdf_list_selected" : self.pdflist.curselection(),
        }

    def setData(self, project_list, name_list, pdf_list, project_sort, name_sort, pdf_sort, input_file, sheet, output_file, sheet_list, offset, range, genpdf, project_list_selected, name_list_selected, pdf_list_selected):
        self.projectlist['listvariable'] = tk.StringVar(value=project_list)
        self.namelist['listvariable'] = tk.StringVar(value=name_list)
        self.pdflist['listvariable'] = tk.StringVar(value=pdf_list)
        self.setting_input.delete(0, tk.END)
        self.setting_input.insert(tk.END, input_file)
        self.setting_sheet['values'] = sheet_list
        self.setting_sheet_data.set(sheet)
        self.setting_output.delete(0, tk.END)
        self.setting_output.insert(tk.END, output_file)
        self.setting_offset_column.delete(0, tk.END)
        self.setting_offset_column.insert(tk.END, offset[0])
        self.setting_offset_row.delete(0, tk.END)
        self.setting_offset_row.insert(tk.END, offset[1])
        self.setting_range_column.delete(0, tk.END)
        self.setting_range_column.insert(tk.END, range[0])
        self.setting_range_row.delete(0, tk.END)
        self.setting_range_row.insert(tk.END, range[1])
        self.setting_genpdf_data.set(genpdf)

    def run(self):
        self.window.mainloop()

class Controller:
    def __init__(self, view):
        self.view = view

        # params
        self.projects = []
        self.project_sort = ""
        self.name_sort = ""
        self.pdf_sort = ""
        # extend params
        self.project_list_selected = []
        self.name_list_selected = []
        self.pdf_list_selected = []
        # other param
        self.current_project = None
        self.current_name = None

        # attributes
        self.updateData_attrs = ["project_sort", "name_sort", "pdf_sort", "input_file", "sheet", "output_file", "offset", "range", "genpdf", "project_list_selected", "name_list_selected", "pdf_list_selected"]
        self.updateView_attrs = ["project_list", "name_list", "pdf_list", "project_sort", "name_sort", "pdf_sort", "input_file", "sheet_list", "sheet", "output_file", "offset", "range", "genpdf", "project_list_selected", "name_list_selected", "pdf_list_selected"]
        self.save_attrs = ["projects", "project_sort", "name_sort", "pdf_sort"]
        #self.save_attrs = ["_name_list", "name_sort", "pdf_sort", "_input_file", "sheet", "output_file", "offset", "range", "name_list_selected", "pdf_list_selected"]
        # set listener
        self.view.on_push_projectlistop_new = [self.updateData, self.push_projectlistop_new, self.updateView]
        self.view.on_push_projectlistop_trash = [self.updateData, self.push_projectlistop_trash, self.updateView]
        self.view.on_push_addbutton = [self.updateData, self.push_addbutton, self.updateView]
        self.view.on_push_namelistop_trash = [self.updateData, self.push_namelistop_trash, self.updateView]
        self.view.on_push_listop_add = [self.updateData, self.push_listop_add, self.updateView]
        self.view.on_push_pdflistop_trash = [self.updateData, self.push_pdflistop_trash, self.updateView]
        self.view.on_push_edit_save = [self.updateData, self.push_edit_save, self.updateView]
        self.view.on_push_setting_input_button = [self.updateData, self.push_setting_input_button, self.updateView]
        self.view.on_push_setting_sheet_button = [self.updateData, self.push_setting_sheet_button, self.updateView]
        self.view.on_push_setting_output_button = [self.updateData, self.push_setting_output_button, self.updateView]
        self.view.on_push_genpdf_button = [self.updateData, self.push_genpdf_button, self.updateView]

        self.view.on_select_projectlist = [self.updateData, self.select_projectlist, self.updateView]
        self.view.on_select_namelist = [self.updateData, self.select_namelist, self.updateView]
        self.view.on_select_pdflist = [self.updateData, self.select_pdflist, self.updateView]

        self.view.on_change_projectlistop_sort = [self.updateData, self.change_projectlistop_sort, self.updateView]
        self.view.on_change_namelistop_sort = [self.updateData, self.change_namelistop_sort, self.updateView]
        self.view.on_change_pdflistop_sort = [self.updateData, self.change_pdflistop_sort, self.updateView]
        self.view.on_change_setting_sheet = [self.updateData, self.change_setting_sheet, self.updateView]

        self.view.on_change_setting_genpdf = [self.updateData, self.change_setting_genpdf, self.updateView]

        self.view.on_close_window = [self.updateData, self.close_window]
        # params initialize
        self.loadData()
        self.updateView()

    @property
    def project_list(self):
        return Project.getProjectList(self.projects, self.current_project)

    @property
    def name_list(self):
        if self.current_project is None:
            return []
        else:
            return Name.getNameList(self.current_project.name_list)

    @property
    def pdf_list(self):
        if self.current_project is None:
            return []
        else:
            return Name.getPdfList(self.current_project.name_list)

    @property
    def _name_list(self):
        if self.current_project is None:
            return []
        else:
            return self.current_project.name_list

    @property
    def _pdf_list(self):
        if self.current_project is None:
            return []
        else:
            return self.current_project.pdf_list

    @property
    def input_file(self):
        if self.current_project is None:
            return ""
        else:
            return self.current_project.input_file.filename

    @input_file.setter
    def input_file(self, value):
        if self.current_project is not None:
            self.current_project.input_file.setFilename(value)

    @property
    def sheet_list(self):
        if self.current_project is None:
            return []
        else:
            return self.current_project.input_file.sheets

    @property
    def sheet(self):
        if self.current_project is None:
            return ""
        else:
            if len(self.current_project.input_file.sheets) == 0:
                return "ファイルが存在しません"
            else:
                #return self.current_project.input_file.sheets[self.current_project.input_file.enable_sheet]
                return self.current_project.input_file.enable_sheet

    @sheet.setter
    def sheet(self, value):
        if self.current_project is not None:
            if value in self.current_project.input_file.sheets:
                #self.current_project.input_file.enable_sheet = self.current_project.input_file.sheets.index(value)
                self.current_project.input_file.enable_sheet = value

    @property
    def output_file(self):
        if self.current_project is None:
            return ""
        else:
            return self.current_project.output_file

    @output_file.setter
    def output_file(self, value):
        if self.current_project is not None:
            self.current_project.output_file = value

    @property
    def offset(self):
        if self.current_project is None:
            return (0, 0)
        else:
            return self.current_project.offset

    @offset.setter
    def offset(self, value):
        if self.current_project is not None:
            self.current_project.offset = value

    @property
    def range(self):
        if self.current_project is None:
            return (0, 0)
        else:
            return self.current_project.range

    @range.setter
    def range(self, value):
        if self.current_project is not None:
            self.current_project.range = value

    @property
    def genpdf(self):
        if self.current_project is None:
            return False
        else:
            return self.current_project.genpdf

    @genpdf.setter
    def genpdf(self, value):
        if self.current_project is not None:
            self.current_project.genpdf = value

    def updateData(self, *arg):
        for key, value in self.view.getData().items():
            if key in self.updateData_attrs:
                setattr(self, key, value)
        Project.sort(self.projects, self.project_sort)
        if self.current_project is not None:
            Name.sort(self.current_project.name_list, self.name_sort, self.pdf_sort)

    def updateView(self, *arg):
        Project.sort(self.projects, self.project_sort)
        if self.current_project is not None:
            Name.sort(self.current_project.name_list, self.name_sort, self.pdf_sort)
        self.view.setData(**{k : getattr(self, k) for k in self.updateView_attrs})

    def loadData(self, filepath="./userdata.dat"):
        try:
            with open(filepath, mode='r', encoding='cp932', errors='ignore') as f:
                data = json.loads(f.read())
                for key, value in data.items():
                    setattr(self, key, value)
        except:
            print("Error : loadData")
        self.projects = Project.loadProjects(self.projects)

    def exportData(self, filepath="./userdata.dat"):
        data = {x : getattr(self, x) for x in self.save_attrs}
        data['projects'] = Project.toDictProjects(self.projects)
        with open(filepath, mode='w', encoding='cp932', errors='ignore') as f:
            f.write(json.dumps(data))

    def push_projectlistop_new(self):
        log("push_projectlistop_new")
        self.projects.append(Project())

    def push_projectlistop_trash(self):
        log("push_projectlistop_trash")
        selected = self.project_list_selected[0]
        popins = [ins for ins in self.projects if ins.list_index is selected][0]
        self.projects.pop(self.projects.index(popins))
        self.view.setInfo('「' + popins.name + '」を削除しました．')

    def push_addbutton(self):
        log("push_addbutton")
        name = self.view.namebox.get()
        if name is "" or name is None:
            return
        self._name_list.append(Name(name))

    def push_namelistop_trash(self):
        if len(self.name_list_selected) is 0:
            return
        log("push_namelistop_trash")
        selected = self.name_list_selected[0]
        popins = [ins for ins in self._name_list if ins.name_list_index is selected][0]
        self._name_list.pop(self._name_list.index(popins))
        self.view.setInfo('「' + popins.name + '」を削除しました．')

    def push_listop_add(self):
        if len(self.name_list_selected) is 0:
            return
        log("push_listop_add")
        selected = self.name_list_selected[0]
        addins = [ins for ins in self._name_list if ins.name_list_index is selected][0]
        addins.in_pdf_list = True
        self.view.setInfo('「' + addins.name + '」をPDF化リストに追加しました．')

    def push_pdflistop_trash(self):
        if len(self.pdf_list_selected) is 0:
            return
        log("push_pdflistop_trash")
        selected = self.pdf_list_selected[0]
        popins = [ins for ins in self._name_list if ins.pdf_list_index is selected][0]
        popins.in_pdf_list = False
        self.view.setInfo('「' + popins.name + '」をPDF化リストから削除しました．')

    def push_edit_save(self):
        if self.current_name is None:
            return
        log("push_edit_save")
        name = self.view.edit_name.get()
        pdf_password = self.view.edit_pdfpass.get()
        zip_password = self.view.edit_zippass.get()
        pdf_filename = self.view.edit_pdfname.get()
        zip_filename = self.view.edit_zipname.get()
        # check
        if name == "" or name is None:
            print("Error : invalid name")
            self.view.setInfo('名前を入力してください．')
            return
        if pdf_filename == "":
            print("Error : invalid pdf_filename")
            self.view.setInfo('PDFファイル名を入力してください．')
            return
        if zip_filename == "":
            print("Error : invalid zip_filename")
            self.view.setInfo('ZIPファイル名を入力してください．')
            return
        # correct
        pdf_password = pdf_password if pdf_password != "" else None
        zip_password = zip_password if zip_password != "" else None
        pdf_filename = pdf_filename if pdf_filename.endswith(".pdf") else pdf_filename + ".pdf"
        zip_filename = zip_filename if zip_filename.endswith(".zip") else zip_filename + ".zip"
        # save data
        #selected = self.name_list_selected[0]
        #editins = [ins for ins in self._name_list if ins.name_list_index is selected][0]
        self.current_name.name = name
        self.current_name.pdf_password = pdf_password
        self.current_name.zip_password = zip_password
        self.current_name.pdf_filename = pdf_filename
        self.current_name.zip_filename = zip_filename
        self.view.setEdit(self.current_name)
        self.view.setInfo('「' + name + '」の変更内容を保存しました．')

    def push_setting_input_button(self):
        log("push_setting_input_button")
        # select file
        dir = os.path.dirname(self.view.setting_input.get())
        file_type = [('エクセルファイル', '*.xlsx'), ('', '*')]
        file = filedialog.askopenfilename(filetypes=file_type, initialdir=dir)
        if len(file) == 0:
            return
        self.input_file = file
        # load
        self.view.setInfo(self.current_project.input_file.filename + 'をロード中')
        if not self.current_project.input_file.reloadSheets():
            self.view.setInfo(self.current_project.input_file.filename + 'が見つかりませんでした．')
        else:
            self.view.setInfo(self.current_project.input_file.filename + 'を正常に読み込みました．')

    def push_setting_sheet_button(self):
        log("push_setting_sheet_button")
        self.view.setInfo(self.current_project.input_file.filename + 'をロード中')
        self.current_project.input_file.reloadSheets()
        if not self.current_project.input_file.reloadSheets():
            self.view.setInfo(self.current_project.input_file.filename + 'が見つかりませんでした．')
        else:
            self.view.setInfo(self.current_project.input_file.filename + 'を正常に読み込みました．')

    def push_setting_output_button(self):
        log("push_setting_output_button")
        dir = os.path.dirname(self.view.setting_output.get())
        directory = filedialog.askdirectory(initialdir=dir)
        if directory is "":
            return
        self.output_file = directory

    def push_genpdf_button(self):
        log("push_genpdf_button")
        # ready
        self.view.setInfo('生成を開始します．（しばらくウィンドウが固まります．）')
        self.view.enableGen()

        # generate pdf
        is_success = True
        for project in self.projects:
            if project.genpdf:
                is_success &= project.genPDF(self.view.setInfo)

        # message
        if is_success:
            self.view.setInfo('正常に終了しました．')
        else:
            self.view.setInfo('何らかの問題が発生しました．出力ファイルが欠損している可能性があります．')

        self.view.disableGen()

    def select_projectlist(self, event):
        if len(self.project_list_selected) is 0:
            return
        log("select_projectlist : ", self.project_list[self.project_list_selected[0]])
        selected = self.project_list_selected[0]
        editins = [ins for ins in self.projects if ins.list_index is selected][0]
        self.current_project = editins

    def select_namelist(self, event):
        if len(self.name_list_selected) is 0:
            return
        log("select_namelist : ", self.name_list[self.name_list_selected[0]])
        selected = self.name_list_selected[0]
        editins = [ins for ins in self._name_list if ins.name_list_index is selected][0]
        self.current_name = editins
        self.view.setEdit(editins)

    def select_pdflist(self, event):
        if len(self.pdf_list_selected) is 0:
            return
        log("select_pdflist : ", self.pdf_list[self.pdf_list_selected[0]])

    def change_projectlistop_sort(self, event):
        log("change_projectlistop_sort", self.project_sort)

    def change_namelistop_sort(self, event):
        log("change_namelistop_sort", self.name_sort)

    def change_pdflistop_sort(self, event):
        log("change_pdflistop_sort", self.pdf_sort)

    def change_setting_sheet(self, event):
        log("change_setting_sheet", self.sheet_list)

    def change_setting_genpdf(self):
        log("change_setting_genpdf")

    def close_window(self):
        log("close_window")
        self.exportData()

class Project:
    @staticmethod
    def loadProjects(list):
        return [Project().load(data) for data in list]

    @staticmethod
    def toDictProjects(projects):
        return [project.toDict() for project in projects]

    @staticmethod
    def getProjectList(projects, current_project=None):
        list = [project for project in sorted(projects, key=lambda ins : ins.list_index)]
        return [project.toStr(project is current_project) for project in list]

    @staticmethod
    def sort(projects, project_sort="降順"):
        sorted_list = []
        if project_sort == "昇順":
            sorted_list = sorted(projects, key=lambda ins : ins.name, reverse=False)
        elif project_sort == "降順":
            sorted_list = sorted(projects, key=lambda ins : ins.name, reverse=True)
        for i in range(0, len(sorted_list)):
            sorted_list[i].list_index = i

    def __init__(self,
        name_list=None,
        input_file=None,
        list_index=None,
        output_file=None,
        offset=None,
        range=None,
        genpdf=None
    ):
        # params (initial value)
        self.name_list = name_list or []
        self.input_file = input_file or ExcelFile()
        self.list_index = list_index or -1
        self.output_file = output_file or ""
        self.offset = offset or (0, 0)
        self.range = range or (3, 3)
        self.genpdf = genpdf or False

        self.save_attrs = ["name_list", "input_file", "list_index", "output_file", "offset", "range", "genpdf"]

    def __repr__(self):
        s = ""
        for k,v in Project.toDictProjects([self])[0].items():
            s += str(k) + " : " + str(v) + ", "
        return s

    def __str__(self):
        s = ""
        for k,v in Project.toDictProjects([self])[0].items():
            s += str(k) + " : " + str(v) + ", "
        return s

    @property
    def name(self):
        if self.input_file.filename != "":
            return os.path.splitext(os.path.basename(
                self.input_file.filename
            ))[0] + ' - ' + self.input_file.enable_sheet_name
        else:
            return '新規プロジェクト'

    def load(self, dict):
        for k, v in dict.items():
            if k == 'name_list':
                value = Name.load(v)
            elif k == 'input_file':
                value = ExcelFile().load(v)
            else:
                value = v
            setattr(self, k, value)
        return self

    def toDict(self):
        data = {x : getattr(self, x) for x in self.save_attrs}
        data['name_list'] = Name.toDictNames(self.name_list)
        data['input_file'] = self.input_file.toDict()
        return data

    def toStr(self, isEdit=False):
        name = "有効 " if self.genpdf else "　　 "
        name += "編集中 " if isEdit else "　　　 "
        name += self.name
        return name

    def genPDF(self, infofunc=None):
        pdflist = Name.getPdfGroup(self.name_list)
        Name.resetSuccess(pdflist)
        result = self.input_file.generatePDF(self.output_file, self.name, pdflist, self.offset, self.range, infofunc)
        return result

class Name:
    @staticmethod
    def load(list):
        return [Name(**data) for data in list]

    @staticmethod
    def toDictNames(names):
        save_attrs = ['name', 'name_list_index', 'pdf_list_index', 'in_pdf_list', 'pdf_password', 'zip_password', 'pdf_filename', 'zip_filename']
        return [{key : getattr(name, key) for key in save_attrs} for name in names]

    @staticmethod
    def getNameList(names):
        return [name.name for name in sorted(names, key=lambda ins : ins.name_list_index)]

    @staticmethod
    def getPdfList(names):
        pdflist = [name for name in names if name.in_pdf_list]
        return [name.name for name in sorted(pdflist, key=lambda ins : ins.pdf_list_index)]

    @staticmethod
    def getPdfGroup(names):
        return [name for name in names if name.in_pdf_list]

    @staticmethod
    def sort(names, name_sort="降順", pdf_sort="降順"):
        Name.sortName(names, name_sort)
        Name.sortPdf(names, pdf_sort)

    @staticmethod
    def sortName(names, name_sort="降順"):
        sorted_list = []
        if name_sort == "昇順":
            sorted_list = sorted(names, key=lambda ins : ins.name, reverse=False)
        elif name_sort == "降順":
            sorted_list = sorted(names, key=lambda ins : ins.name, reverse=True)
        for i in range(0, len(sorted_list)):
            sorted_list[i].name_list_index = i

    @staticmethod
    def sortPdf(names, pdf_sort="降順"):
        pdflist = [name for name in names if name.in_pdf_list]
        sorted_list = []
        if pdf_sort == "昇順":
            sorted_list = sorted(pdflist, key=lambda ins : ins.name, reverse=False)
        elif pdf_sort == "降順":
            sorted_list = sorted(pdflist, key=lambda ins : ins.name, reverse=True)
        for i in range(0, len(sorted_list)):
            sorted_list[i].pdf_list_index = i

    @staticmethod
    def resetSuccess(names):
        for name in names:
            name.is_success = False

    def __init__(self,
        name,
        name_list_index=None,
        pdf_list_index=None,
        in_pdf_list=None,
        pdf_password=None,
        zip_password=None,
        pdf_filename=None,
        zip_filename=None
    ):
        self.name = name
        self.name_list_index = name_list_index or -1
        self.pdf_list_index = pdf_list_index or -1
        self.in_pdf_list = in_pdf_list or False
        self.pdf_password = pdf_password
        self.zip_password = zip_password
        self.pdf_filename = pdf_filename or "pdffile.pdf"
        self.zip_filename = zip_filename or "zipfile.zip"

        # non param
        self.is_success = False

    def __repr__(self):
        s = ""
        for k,v in Name.toDictNames([self])[0].items():
            s += str(k) + " : " + str(v) + ", "
        return s

    def __str__(self):
        s = ""
        for k,v in Name.toDictNames([self])[0].items():
            s += str(k) + " : " + str(v) + ", "
        return s

class ExcelFile:
    excel = None
    tmp_dir = "excelfile_temporary"
    id = -1

    '''
    @staticmethod
    def load(list):
        return [ExcelFile(**data) for data in list]
    '''

    @staticmethod
    def toDictExcelFiles(names):
        save_attrs = ['filename', 'password']
        return [{key : getattr(name, key) for key in save_attrs} for name in names]
    '''
    @staticmethod
    def genPDF(excel_files, pdf_file, name_list, offset, range, infofunc):
        for ef in excel_files:
            ef.generatePDF(pdf_file, name_list, offset, range, infofunc)
    '''

    @classmethod
    def loadExcel(cls):
        try:
            _ = cls.excel.Workbooks
        except:
            try:
                cls.excel = win32com.client.DispatchEx("Excel.Application")
            except Exception as e:
                print("Error : Don't run excel com. Details - ", e)
        if cls.excel is None:
            print('Error : No app found')
            return False
        return True

    @classmethod
    def genId(cls):
        cls.id += 1
        return cls.id

    @classmethod
    def Temporary(cls):
        tmp_dir = os.path.abspath(cls.tmp_dir)
        if not os.path.isdir(tmp_dir):
            os.mkdir(tmp_dir)
        return tmp_dir

    def __init__(self, filename="", password=""):
        self.id = ExcelFile.genId()
        self.filename = filename
        self.password = password

        # data fields
        self.sheets = []

        # state
        self.enable_sheet = 0

        # private params
        self.tmp_filename = ""

        self.setFilename(filename)

        self.workbook = None

        self.save_attrs = ['filename', 'password', 'enable_sheet']

    def __del__(self):
        self.discard()

    def discard(self):
        try:
            self.workbook.Close(False)
        except:
            pass
        tryRemoveFile(self.tmp_filename)

    def setFilename(self, filename):
        self.filename = filename
        self.setTmpFilename()

    def setTmpFilename(self):
        self.discard()
        self.tmp_filename = os.path.abspath(os.path.join(
            ExcelFile.Temporary(),
            str(self.id) + os.path.splitext(self.filename)[1]
        ))

    @property
    def enable_sheet_name(self):
        if self.enable_sheet in self.sheets:
            return self.enable_sheet
        else:
            return ""

    def load(self, dict):
        for k, v in dict.items():
            setattr(self, k, v)
        self.reloadfile()
        self.reloadSheets()
        return self

    def toDict(self):
        data = {x : getattr(self, x) for x in self.save_attrs}
        return data

    def reloadfile(self):
        try:
            tryRemoveFile(self.tmp_filename)
            log('remove - '+ self.tmp_filename + ' : ', not os.path.exists(self.tmp_filename))
            shutil.copyfile(self.filename, self.tmp_filename)
            log('copy - '+ self.tmp_filename + ' : ', os.path.exists(self.tmp_filename))
        except Exception as e:
            print("Error : cannot copy file - ", self.filename, e)
            return False
        return True

    def reloadWorkbook(self):
        ExcelFile.loadExcel()
        try:
            self.workbook.Close(False)
        except Exception as e:
            print("failed to close workbook - ", e)
        try:
            self.workbook = ExcelFile.excel.Workbooks.Open(
                self.tmp_filename,
                None,
                True,
                Password=self.password
            )
            ExcelFile.excel.Visible = False
            # check
            _ = self.workbook.Worksheets
            self.usable = True
        except Exception as e:
            print("Error : cannot open file - ", self.tmp_filename, e)
            self.usable = False
            return False
        return True

    def reloadSheets(self):
        if not self.reloadfile() or not self.reloadWorkbook():
            log("faild reload")
            self.sheets = []
            return False
        try:
            self.sheets = [sheet.name for sheet in self.workbook.Worksheets]
            log("success reload")
        except Exception as e:
            self.sheets = []
            print('Error : excel client error - ', e)
            return False
        return True

    def generatePDF(self, root_dir, project_name, name_list, offset, range, infofunc=None):
        if not self.reloadSheets():
            if infofunc:
                infofunc(self.filename + 'の読み込みに失敗しました．')
            return False

        is_success = True

        # create directory
        rawpdf_dir = os.path.abspath(os.path.join(root_dir, "rawpdf"))
        encrypt_dir = os.path.abspath(os.path.join(root_dir, "encrypt"))
        save_dir = os.path.abspath(os.path.join(root_dir, project_name))
        if not os.path.isdir(rawpdf_dir):
            os.mkdir(rawpdf_dir)
        if not os.path.isdir(encrypt_dir):
            os.mkdir(encrypt_dir)
        if not os.path.isdir(save_dir):
            os.mkdir(save_dir)

        try:
            sheet = self.workbook.Worksheets[self.enable_sheet]
            sheet.Activate()
            for name in name_list:
                try:
                    # erase print range
                    sheet.ResetAllPageBreaks()

                    # search name
                    result = sheet.UsedRange.Find(name.name)
                    if result is None:
                        continue

                    # calc range
                    upperleft = sheet.Cells(
                        result.Row + offset[1],
                        result.Column + offset[0]
                    )
                    bottomright = sheet.Cells(
                        result.Row + offset[1] + range[1] - 1,
                        result.Column + offset[0] + range[0] - 1
                    )

                    # set print range
                    print_range = upperleft.Address + ":" + bottomright.Address
                    sheet.PageSetup.PrintArea = print_range

                    # save as pdf file
                    rawpdffile = os.path.abspath(os.path.join(rawpdf_dir, name.name.replace(" ", "") + ".pdf"))
                    if os.path.exists(rawpdffile):
                        os.remove(rawpdffile)
                    log("save : ", rawpdffile, ", range : ", print_range)
                    if infofunc:
                        infofunc('「' + name.name + '」のPDFファイルを作成中')
                    sheet.ExportAsFixedFormat(0, rawpdffile)

                    # set password to pdf file
                    encryptfile = os.path.abspath(os.path.join(encrypt_dir, name.name.replace(" ", "") + ".pdf"))
                    if os.path.exists(encryptfile):
                        os.remove(encryptfile)
                    log("save : ", encryptfile)
                    if infofunc:
                        infofunc('「' + name.name + '」のPDFファイルを暗号化中')
                    rawpdf = Pdf.open(rawpdffile)
                    encryptpdf = Pdf.new()
                    encryptpdf.pages.extend(rawpdf.pages)
                    encryptpdf.save(encryptfile, encryption=pikepdf.Encryption(
                        user=name.pdf_password or "", owner=name.pdf_password or ""
                    ))
                    rawpdf.close()
                    encryptpdf.close()

                    # create zip file
                    zip_dir = os.path.abspath(os.path.join(save_dir, name.name))
                    if not os.path.isdir(zip_dir):
                        os.mkdir(zip_dir)
                    zipfile = os.path.abspath(os.path.join(zip_dir, name.zip_filename.replace(" ", "")))
                    if os.path.exists(zipfile):
                        os.remove(zipfile)
                    log("save : ", zipfile)
                    if infofunc:
                        infofunc('「' + name.name + '」をZIPに圧縮中')
                    pyminizip.compress(
                        encryptfile.encode('cp932'), '', zipfile.encode('cp932'), name.zip_password or "", int(0)
                    )
                    name.is_success = True
                    if infofunc:
                        infofunc('「' + name.name + '」のファイル生成を完了しました．')
                except:
                    if infofunc:
                        infofunc('「' + name.name + '」のファイル生成に失敗しました．')
                    is_success &= False
        except Exception as e:
            print('Error : cannot save as pdf.', e)
            if infofunc:
                infofunc('エクセルファイルの処理中に問題が発生しました．')
            is_success &= False
        return is_success
