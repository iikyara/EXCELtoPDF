# coding: utf-8
import sys
import os
import re
import datetime
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
from tkinter_extend import *
from settings import *
from utils import *
from aipo_message import *

print("core - DEBUG :", DEBUG)

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
        self.window.title("EXCELtoPDF" + (" - debug mode" if DEBUG else ""))
        self.window.geometry("1000x480")
        self.window.iconbitmap(default='icon.ico')

        # info data
        self.info_data = tk.StringVar()

        # textbox data
        self.edit_pdfname_data = tk.StringVar()
        self.edit_zipname_data = tk.StringVar()
        self.edit_aipo_id_data = tk.StringVar()

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
        self.edit_default_pdf_data = tk.BooleanVar()
        self.edit_default_zip_data = tk.BooleanVar()
        self.edit_send_aipo_data = tk.BooleanVar()
        self.setting_genpdf_data = tk.BooleanVar()

        # text box
        self.namebox = None
        self.edit_name = None
        self.edit_pdfpass = None
        self.edit_zippass = None
        self.edit_pdfname = None
        self.edit_zipname = None
        self.edit_aipo_id = None
        self.setting_input = None
        self.setting_input_password = None
        self.setting_offset_column = None
        self.setting_offset_row = None
        self.setting_range_column = None
        self.setting_range_row = None
        self.setting_output = None

        # button
        self.projectlistop_duplicate = None
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
        self.sendaipo_button = None

        # list
        self.projectlist = None
        self.namelist = None
        self.pdflist = None

        # combobox
        self.namelistop_sort = None
        self.pdflistop_sort = None
        self.setting_sheet = None

        # checkbox
        self.edit_default_pdf = None
        self.edit_default_zip = None
        self.edit_send_aipo = None
        self.setting_genpdf = None

        # listner func
        self.on_push_projectlistop_duplicate = []
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
        self.on_push_sendaipo_button = []

        self.on_select_projectlist = []
        self.on_select_namelist = []
        self.on_select_pdflist = []

        self.on_change_projectlistop_sort = []
        self.on_change_namelistop_sort = []
        self.on_change_pdflistop_sort = []
        self.on_change_setting_sheet = []

        self.on_change_edit_default_pdf = []
        self.on_change_edit_default_zip = []
        self.on_change_edit_send_aipo = []
        self.on_change_setting_genpdf = []

        self.on_change_edit_aipo_id = []

        self.on_close_window = []

        # create GUI
        self.createGUI()

        # set listener
        # close window listener
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)
        # button listener
        self.projectlistop_duplicate['command'] = self.push_projectlistop_duplicate
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
        self.sendaipo_button['command'] = self.push_sendaipo_button
        # listbox changed listner
        self.projectlist.bind('<<ListboxSelect>>', self.select_projectlist)
        self.namelist.bind('<<ListboxSelect>>', self.select_namelist)
        self.pdflist.bind('<<ListboxSelect>>', self.select_pdflist)
        # combobox selected listener
        self.projectlistop_sort.bind('<<ComboboxSelected>>', self.change_projectlistop_sort)
        self.namelistop_sort.bind('<<ComboboxSelected>>', self.change_namelistop_sort)
        self.pdflistop_sort.bind('<<ComboboxSelected>>', self.change_pdflistop_sort)
        self.setting_sheet.bind('<<ComboboxSelected>>', self.change_setting_sheet)
        # checkbox changed listener
        self.edit_default_pdf['command'] = self.change_edit_default_pdf
        self.edit_default_zip['command'] = self.change_edit_default_zip
        self.edit_send_aipo['command'] = self.change_edit_send_aipo
        self.setting_genpdf['command'] = self.change_setting_genpdf
        # entry widget listener
        self.edit_aipo_id.bind('<KeyRelease>', self.change_edit_aipo_id)

    def createGUI(self):
        # frames
        self.main_frame = ttk.Frame(self.window, style="Main.TFrame")
        self.project_frame = ttk.Frame(self.main_frame, style="A.TFrame")
        self.left_frame = ttk.Frame(self.main_frame, style="A.TFrame")
        self.right_frame = ttk.Frame(self.main_frame, style="B.TFrame")

        self.main_frame.grid(column=0, row=0, sticky=tk.NSEW)
        self.project_frame.grid(column=0, row=0, sticky=tk.NSEW, padx=1, pady=1)
        self.left_frame.grid(column=1, row=0, sticky=tk.NSEW, padx=1, pady=1)
        self.right_frame.grid(column=2, row=0, sticky=tk.NSEW, padx=1, pady=1)

        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        #self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.columnconfigure(2, weight=1)
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
        self.projectlistop_duplicate = ttk.Button(self.projectlistop_frame, text="複製", width=5)
        self.projectlistop_new = ttk.Button(self.projectlistop_frame, text="新規作成", width=10)
        self.projectlistop_trash = ttk.Button(self.projectlistop_frame, text="削除", width=5)

        self.projectlistop_frame.grid(column=0, row=4, sticky=tk.EW)
        self.projectlistop_sort_label.grid(column=0, row=0, sticky=tk.W)
        self.projectlistop_sort.grid(column=1, row=0, sticky=tk.W)
        self.projectlistop_duplicate.grid(column=2, row=0, sticky=tk.E)
        self.projectlistop_new.grid(column=3, row=0, sticky=tk.E)
        self.projectlistop_trash.grid(column=4, row=0, sticky=tk.E)

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
        self.edit_pdfname = ttk.Entry(self.edit_frame, textvariable=self.edit_pdfname_data)
        self.edit_default_pdf = ttk.Checkbutton(self.edit_frame, text="デフォルトのファイル名を使用", variable=self.edit_default_pdf_data)
        # zip filename
        self.edit_zipname_label = ttk.Label(self.edit_frame, text="ZIPファイル名：")
        self.edit_zipname = ttk.Entry(self.edit_frame, textvariable=self.edit_zipname_data)
        self.edit_default_zip = ttk.Checkbutton(self.edit_frame, text="デフォルトのファイル名を使用", variable=self.edit_default_zip_data)
        # aipo id
        self.edit_aipo_id_label = ttk.Label(self.edit_frame, text="Aipo ID：")
        self.edit_aipo_id = ttk.Entry(self.edit_frame, textvariable=self.edit_aipo_id_data)
        # aipo check
        self.edit_send_aipo_label = ttk.Label(self.edit_frame, text="Aipoでメッセージを送信：")
        self.edit_send_aipo = ttk.Checkbutton(self.edit_frame, variable=self.edit_send_aipo_data)
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
        self.edit_default_pdf.grid(column=1, row=5, sticky=tk.E)
        self.edit_zipname_label.grid(column=0, row=6, sticky=tk.E)
        self.edit_zipname.grid(column=1, row=6, sticky=tk.EW)
        self.edit_default_zip.grid(column=1, row=7, sticky=tk.E)
        self.edit_aipo_id_label.grid(column=0, row=8, sticky=tk.E)
        self.edit_aipo_id.grid(column=1, row=8, sticky=tk.EW)
        self.edit_send_aipo_label.grid(column=0, row=9, sticky=tk.E)
        self.edit_send_aipo.grid(column=1, row=9, sticky=tk.W)
        #self.edit_save.grid(column=1, row=8, sticky=tk.E)

        self.edit_frame.columnconfigure(1, weight=1)

        # setting
        self.setting_frame = ttk.Frame(self.right_frame)
        self.setting_lable = ttk.Label(self.setting_frame, text="設定")
        # input file
        self.setting_input_label = ttk.Label(self.setting_frame, text="入力ファイル：")
        self.setting_input_frame = ttk.Frame(self.setting_frame)
        self.setting_input = ttk.Entry(self.setting_input_frame)
        self.setting_input_button = ttk.Button(self.setting_input_frame, text="参照", width=5)
        self.setting_input_password_label = ttk.Label(self.setting_frame, text="パスワード：")
        self.setting_input_password = ttk.Entry(self.setting_frame)
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
        self.setting_input_password_label.grid(column=0, row=2, sticky=tk.E)
        self.setting_input_password.grid(column=1, row=2, sticky=tk.EW)
        self.setting_sheet_label.grid(column=0, row=3, sticky=tk.E)
        self.setting_sheet_frame.grid(column=1, row=3, sticky=tk.EW)
        self.setting_sheet.grid(column=0, row=0, sticky=tk.EW)
        self.setting_sheet_button.grid(column=1, row=0, sticky=tk.E)
        self.setting_offset_label.grid(column=0, row=4, sticky=tk.E)
        self.setting_offset_frame.grid(column=1, row=4, sticky=tk.EW)
        self.setting_offset_column.grid(column=0, row=0)
        self.setting_offset_middle.grid(column=1, row=0)
        self.setting_offset_row.grid(column=2, row=0)
        self.setting_range_label.grid(column=0, row=5, sticky=tk.E)
        self.setting_range_frame.grid(column=1, row=5, sticky=tk.EW)
        self.setting_range_column.grid(column=0, row=0)
        self.setting_range_middle.grid(column=1, row=0)
        self.setting_range_row.grid(column=2, row=0)
        self.setting_output_label.grid(column=0, row=6, sticky=tk.E)
        self.setting_output_frame.grid(column=1, row=6, sticky=tk.EW)
        self.setting_output.grid(column=0, row=0, sticky=tk.EW)
        self.setting_output_button.grid(column=1, row=0, sticky=tk.E)
        self.setting_genpdf_label.grid(column=0, row=7, sticky=tk.E)
        self.setting_genpdf.grid(column=1, row=7, sticky=tk.EW)

        self.setting_frame.columnconfigure(1, weight=1)
        self.setting_input_frame.columnconfigure(0, weight=1)
        self.setting_output_frame.columnconfigure(0, weight=1)
        self.setting_sheet_frame.columnconfigure(0, weight=1)

        # right buttom button frame
        self.right_buttom_frame = ttk.Frame(self.right_frame)

        # generate pdf button
        self.genpdf_button = ttk.Button(self.right_buttom_frame, text="PDFを出力する")
        # send aipo button
        self.sendaipo_button = ttk.Button(self.right_buttom_frame, text="Aipoでメッセージを送信")

        self.right_buttom_frame.grid(column=0, row=2, sticky=tk.NSEW)
        self.genpdf_button.grid(column=0, row=0, sticky=tk.NSEW)
        self.sendaipo_button.grid(column=0, row=1, sticky=tk.NSEW)

        self.right_buttom_frame.rowconfigure(0, weight=1)
        self.right_buttom_frame.rowconfigure(1, weight=1)
        self.right_buttom_frame.columnconfigure(0, weight=1)

        ttk.Style().configure("Main.TFrame", padding=6, relief="flat",   background="#000")
        #ttk.Style().configure("A.TFrame", padding=6, relief="flat",   background="#F00")
        #ttk.Style().configure("B.TFrame", padding=6, relief="flat",  background="#0F0")

    def setEdit(self, name):
        if not name:
            self.edit_name.delete(0, tk.END)
            self.edit_pdfpass.delete(0, tk.END)
            self.edit_zippass.delete(0, tk.END)
            self.edit_pdfname_data.set("")
            self.edit_pdfname.configure(state=tk.NORMAL)
            self.edit_default_pdf_data.set(False)
            self.edit_zipname_data.set("")
            self.edit_zipname.configure(state=tk.NORMAL)
            self.edit_default_zip_data.set(False)
            self.edit_aipo_id.delete(0, tk.END)
            self.edit_send_aipo_data.set(False)
            return
        self.edit_name.delete(0, tk.END)
        self.edit_name.insert(tk.END, name.name)
        self.edit_pdfpass.delete(0, tk.END)
        self.edit_pdfpass.insert(tk.END, name.pdf_password or "")
        self.edit_zippass.delete(0, tk.END)
        self.edit_zippass.insert(tk.END, name.zip_password or "")
        self.edit_pdfname_data.set(name.pdf_filename or "")
        self.edit_pdfname.configure(state=tk.DISABLED if name.is_default_pdf_filename else tk.NORMAL)
        self.edit_default_pdf_data.set(name.is_default_pdf_filename)
        self.edit_zipname_data.set(name.zip_filename or "")
        self.edit_zipname.configure(state=tk.DISABLED if name.is_default_zip_filename else tk.NORMAL)
        self.edit_default_zip_data.set(name.is_default_zip_filename)
        self.edit_aipo_id.delete(0, tk.END)
        self.edit_aipo_id.insert(tk.END, "" if not name.aipo_id or name.aipo_id == -1 else name.aipo_id)
        self.edit_send_aipo_data.set(name.send_aipo)

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
        self.setting_input_password.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_offset_column.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_offset_row.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_range_column.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_range_row.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_output.configure(state=tk.DISABLED if is_gen else tk.NORMAL)

        #button
        self.projectlistop_duplicate.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
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
        self.edit_default_pdf.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.edit_default_zip.configure(state=tk.DISABLED if is_gen else tk.NORMAL)
        self.setting_genpdf.configure(state=tk.DISABLED if is_gen else tk.NORMAL)

        self.window.update()

    def push_projectlistop_duplicate(self):
        for func in self.on_push_projectlistop_duplicate:
            func()

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

    def push_sendaipo_button(self):
        for func in self.on_push_sendaipo_button:
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

    def change_edit_default_pdf(self):
        for func in self.on_change_edit_default_pdf:
            func()

    def change_edit_default_zip(self):
        for func in self.on_change_edit_default_zip:
            func()

    def change_edit_send_aipo(self):
        for func in self.on_change_edit_send_aipo:
            func()

    def change_setting_genpdf(self):
        for func in self.on_change_setting_genpdf:
            func()

    def change_edit_aipo_id(self, event):
        for func in self.on_change_edit_aipo_id:
            func(event)

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
            "edit_name" : Name(
                name=self.edit_name.get(),
                pdf_password=self.edit_pdfpass.get(),
                zip_password=self.edit_zippass.get(),
                _pdf_filename=self.edit_pdfname.get(),
                _zip_filename=self.edit_zipname.get(),
                is_default_pdf_filename=self.edit_default_pdf_data.get(),
                is_default_zip_filename=self.edit_default_zip_data.get(),
                aipo_id=self.edit_aipo_id.get(),
                send_aipo=self.edit_send_aipo_data.get()
            ),
            "input_file" : self.setting_input.get(),
            "input_password" : self.setting_input_password.get(),
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

    def setData(self, project_list, name_list, pdf_list, project_sort, name_sort, pdf_sort, edit_name, input_file, input_password, sheet, output_file, sheet_list, offset, range, genpdf, project_list_selected, name_list_selected, pdf_list_selected):
        self.projectlist['listvariable'] = tk.StringVar(value=project_list)
        self.namelist['listvariable'] = tk.StringVar(value=name_list)
        self.pdflist['listvariable'] = tk.StringVar(value=pdf_list)
        self.setEdit(edit_name)
        self.setting_input.delete(0, tk.END)
        self.setting_input.insert(tk.END, input_file)
        self.setting_input_password.delete(0, tk.END)
        self.setting_input_password.insert(tk.END, input_password)
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

class ResultView:
    def __init__(self, app_window, projects):
        self.window = tk.Toplevel(app_window)

        self.window.title("EXCELtoPDF - 結果")
        self.window.geometry("700x480")
        self.window.iconbitmap(default='icon.ico')

        self.projects = projects

        self.createGUI()

    def createGUI(self):
        self.main_frame = ttk.Frame(self.window)
        self.window_label = ttk.Label(self.main_frame, text="結果")
        self.result_frame = ttk.Frame(self.main_frame, style="Table.TFrame")

        self.main_frame.grid(column=0, row=0, sticky=tk.NSEW)
        self.window_label.grid(column=0, row=0, sticky=tk.NSEW, padx=1, pady=1)
        self.result_frame.grid(column=0, row=1, sticky=[tk.N, tk.EW], padx=5, pady=5)

        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        self.result_frame.columnconfigure(3, weight=1)

        row = 0
        # header
        ttk.Label(self.result_frame, text="プロジェクト名", style="Table.TLabel").grid(column=0, row=row, sticky=tk.EW, padx=1, pady=1)
        ttk.Label(self.result_frame, text="名前", style="Table.TLabel").grid(column=1, row=row, sticky=tk.EW, padx=1, pady=1)
        ttk.Label(self.result_frame, text="結果", style="Table.TLabel").grid(column=2, row=row, sticky=tk.EW, padx=1, pady=1)
        ttk.Label(self.result_frame, text="備考", style="Table.TLabel").grid(column=3, row=row, sticky=tk.EW, padx=1, pady=1)
        row += 1
        # table
        for project in self.projects:
            if not project.genpdf:
                continue

            projectname = ttk.Label(self.result_frame, text=project.name, style="Table.TLabel")

            pdf_group = Name.getPdfGroup(project.name_list)
            projectname.grid(column=0, row=row, rowspan=len(pdf_group), sticky=tk.NSEW, padx=1, pady=1)

            for name in pdf_group:
                namename = ttk.Label(self.result_frame, text=name.name, style="Table.TLabel")
                namesuccess = ttk.Label(self.result_frame, text="成功" if name.is_success else "失敗", style="Table.TLabel")
                namemessage = ttk.Label(self.result_frame, text=name.error_message, style="Table.TLabel")

                namename.grid(column=1, row=row, sticky=tk.EW, padx=1, pady=1)
                namesuccess.grid(column=2, row=row, sticky=tk.EW, padx=1, pady=1)
                namemessage.grid(column=3, row=row, sticky=tk.EW, padx=1, pady=1)
                row += 1

            row += 1

        ttk.Style().configure("Table.TLabel", relief="flat", background="#fff")
        ttk.Style().configure("Table.TFrame", relief="flat", background="#111")

class SendSettingView:
    def __init__(self, app_window, project):
        self.window = tk.Toplevel(app_window)

        self.window.title("EXCELtoPDF - Aipoへメッセージを送信")
        self.window.geometry("1000x480")
        self.window.iconbitmap(default='icon.ico')

        self.project = project

        # text area
        self.sendmsg = None

        # button
        self.cancel_button = None
        self.send_button = None

        # listner func
        self.on_push_cancel_button = []
        self.on_push_send_button = []
        self.on_change_sendmsg = []

        self.on_close_window = []

        # create GUI
        self.createGUI()

        # set listener
        # close winder listener
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)
        # button listener
        self.cancel_button['command'] = self.push_cancel_button
        self.send_button['command'] = self.push_send_button
        # text listener
        self.sendmsg.bind('<<Change>>', self.change_sendmsg)

        # update
        self.setText(self.project.aipo_message)
        self.update()

    def createGUI(self):
        # frames
        self.main_frame = ttk.Frame(self.window)
        # over wrap
        self.window_label = ttk.Label(self.main_frame, text="Aipoへメッセージを送信する人，ファイル名，メッセージを確認してください．")
        # aipo message list
        self.list_frame = ttk.Frame(self.main_frame, style="Table.TFrame")
        '''
        self.list_frame_scrollbar = ttk.Scrollbar(
            self.list_frame,
            orient=tk.VERTICAL,
            command=self.list_frame.yview
        )
        self.list_frame["yscrollcommand"] = self.list_frame_scrollbar.set
        '''
        # send message text box
        self.sendmsg_frame = ttk.Frame(self.main_frame)
        self.sendmsg_label = ttk.Label(self.sendmsg_frame, text="送信メッセージ")
        self.sendmsg = CustomText(self.sendmsg_frame, height=5)
        self.sendmsg_scrollbar = ttk.Scrollbar(
            self.sendmsg,
            orient=tk.VERTICAL,
            command=self.sendmsg.yview
        )
        self.sendmsg["yscrollcommand"] = self.sendmsg_scrollbar.set
        self.sendmsg_preview_label = ttk.Label(self.sendmsg_frame, text="プレビュー（実際に送信されるメッセージ）")
        self.sendmsg_preview = ttk.Label(self.sendmsg_frame, text="")
        # bottom button frame
        self.bottom_frame = ttk.Frame(self.main_frame)

        self.main_frame.grid(column=0, row=0, sticky=tk.NSEW)
        self.window_label.grid(column=0, row=0, sticky=tk.NSEW)
        self.list_frame.grid(column=0, row=1, sticky=[tk.N, tk.EW], padx=5, pady=5)
        self.sendmsg_frame.grid(column=0, row=2, sticky=tk.NSEW)
        self.sendmsg_label.grid(column=0, row=0, sticky=tk.NSEW)
        self.sendmsg.grid(column=0, row=1, sticky=tk.NSEW)
        self.sendmsg_preview_label.grid(column=1, row=0, sticky=tk.NSEW)
        self.sendmsg_preview.grid(column=1, row=1, sticky=tk.NSEW)
        self.bottom_frame.grid(column=0, row=4, sticky=tk.NSEW)

        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        #self.main_frame.rowconfigure(3, weight=1)
        self.list_frame.columnconfigure(3, weight=1)
        self.sendmsg_frame.columnconfigure(0, weight=1)
        self.sendmsg_frame.columnconfigure(1, weight=1)
        self.sendmsg_frame.rowconfigure(1, weight=1)
        self.bottom_frame.columnconfigure(0, weight=1)

        # buttons
        self.cancel_button = ttk.Button(self.bottom_frame, text="キャンセル", width=12)
        self.send_button = ttk.Button(self.bottom_frame, text="送信", width=5)

        self.cancel_button.grid(column=1, row=0, sticky=tk.E)
        self.send_button.grid(column=2, row=0, sticky=tk.E)

        ttk.Style().configure("Table.TLabel", relief="flat", background="#fff")
        ttk.Style().configure("Table.TFrame", relief="flat", background="#111")

    def update(self):
        self.updateTable()
        self.saveText()

    def saveText(self):
        self.project.aipo_message = self.sendmsg.get('1.0', 'end -1c')

    def updateTable(self):
        # destroy table
        for child in self.list_frame.winfo_children():
            child.destroy()
        row = 0
        # header
        ttk.Label(self.list_frame, text="名前", style="Table.TLabel").grid(column=0, row=row, sticky=tk.EW, padx=1, pady=1)
        ttk.Label(self.list_frame, text="AipoID", style="Table.TLabel").grid(column=1, row=row, sticky=tk.EW, padx=1, pady=1)
        ttk.Label(self.list_frame, text="送信ファイル名", style="Table.TLabel").grid(column=2, row=row, sticky=tk.EW, padx=1, pady=1)
        ttk.Label(self.list_frame, text="更新日", style="Table.TLabel").grid(column=3, row=row, sticky=tk.EW, padx=1, pady=1)
        row += 1
        # table
        namelist = sorted(Name.getAipoGroup(self.project.name_list), key=lambda ins : ins.name, reverse=False)
        for name in namelist:
            # name
            ttk.Label(
                self.list_frame,
                text=name.name,
                style="Table.TLabel"
            ).grid(column=0, row=row, sticky=tk.EW, padx=1, pady=1)
            # aipo id
            ttk.Label(
                self.list_frame,
                text=name.aipo_id,
                style="Table.TLabel"
            ).grid(column=1, row=row, sticky=tk.EW, padx=1, pady=1)
            # send file name
            filename = os.path.abspath(os.path.join(self.project.output_file, self.project.name, name.zip_filename))
            ttk.Label(
                self.list_frame,
                text=filename,
                style="Table.TLabel"
            ).grid(column=2, row=row, sticky=tk.EW, padx=1, pady=1)
            # update
            update = datetime.datetime.fromtimestamp(os.stat(filename).st_mtime).strftime('%Y/%m/%d %H:%M:%S') if os.path.exists(filename) else "ファイルが存在しません"
            ttk.Label(
                self.list_frame,
                text=update,
                style="Table.TLabel"
            ).grid(column=3, row=row, sticky=tk.EW, padx=1, pady=1)
            row += 1

    def setText(self, str):
        self.sendmsg.delete('1.0', 'end')
        self.sendmsg.insert('1.0', str)

    def push_cancel_button(self):
        for func in self.on_push_cancel_button:
            func()

    def push_send_button(self):
        for func in self.on_push_send_button:
            func()

    def change_sendmsg(self, event):
        for func in self.on_change_sendmsg:
            func(event)

    def close_window(self):
        for func in self.on_close_window:
            func()
        self.window.destroy()

class SendResultView:
    def __init__(self, app_window, project):
        self.window = tk.Toplevel(app_window)

        self.window.title("EXCELtoPDF - メッセージ送信結果")
        self.window.geometry("700x480")
        self.window.iconbitmap(default='icon.ico')

        self.project = project

        self.createGUI()

    def createGUI(self):
        self.main_frame = ttk.Frame(self.window)
        self.main_frame.grid(column=0, row=0, sticky=tk.NSEW)

        if not self.project.is_success:
            ttk.Label(self.main_frame, text="送信エラー").grid(column=0, row=0, sticky=tk.E)
            ttk.Label(self.main_frame, text=self.project.error_message).grid(column=0, row=1, sticky=tk.E)
            return

        self.window_label = ttk.Label(self.main_frame, text="送信結果")
        self.result_frame = ttk.Frame(self.main_frame, style="Table.TFrame")

        self.window_label.grid(column=0, row=0, sticky=tk.NSEW, padx=1, pady=1)
        self.result_frame.grid(column=0, row=1, sticky=[tk.N, tk.EW], padx=5, pady=5)

        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        self.result_frame.columnconfigure(2, weight=1)

        row = 0
        # header
        ttk.Label(self.result_frame, text="名前", style="Table.TLabel").grid(column=0, row=row, sticky=tk.EW, padx=1, pady=1)
        ttk.Label(self.result_frame, text="結果", style="Table.TLabel").grid(column=1, row=row, sticky=tk.EW, padx=1, pady=1)
        ttk.Label(self.result_frame, text="備考", style="Table.TLabel").grid(column=2, row=row, sticky=tk.EW, padx=1, pady=1)
        row += 1
        # table
        # table
        namelist = sorted(Name.getAipoGroup(self.project.name_list), key=lambda ins : ins.name, reverse=False)
        for name in namelist:
            # name
            ttk.Label(
                self.result_frame,
                text=name.name,
                style="Table.TLabel"
            ).grid(column=0, row=row, sticky=tk.EW, padx=1, pady=1)
            # result
            ttk.Label(
                self.result_frame,
                text="完了" if name.is_success else "送信失敗",
                style="Table.TLabel"
            ).grid(column=1, row=row, sticky=tk.EW, padx=1, pady=1)
            # note
            ttk.Label(
                self.result_frame,
                text=name.error_message,
                style="Table.TLabel"
            ).grid(column=2, row=row, sticky=tk.EW, padx=1, pady=1)
            row += 1

        ttk.Style().configure("Table.TLabel", relief="flat", background="#fff")
        ttk.Style().configure("Table.TFrame", relief="flat", background="#111")

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
        self.sendaipo_window = None

        # attributes
        self.updateData_attrs = ["project_sort", "name_sort", "pdf_sort", "edit_name", "input_file", "input_password", "sheet", "output_file", "offset", "range", "genpdf", "project_list_selected", "name_list_selected", "pdf_list_selected"]
        self.updateView_attrs = ["project_list", "name_list", "pdf_list", "project_sort", "name_sort", "pdf_sort", "edit_name", "input_file", "input_password", "sheet_list", "sheet", "output_file", "offset", "range", "genpdf", "project_list_selected", "name_list_selected", "pdf_list_selected"]
        self.save_attrs = ["projects", "project_sort", "name_sort", "pdf_sort"]
        #self.save_attrs = ["_name_list", "name_sort", "pdf_sort", "_input_file", "sheet", "output_file", "offset", "range", "name_list_selected", "pdf_list_selected"]
        # set listener
        self.view.on_push_projectlistop_duplicate = [self.updateData, self.push_projectlistop_duplicate, self.updateView]
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
        self.view.on_push_sendaipo_button = [self.updateData, self.push_sendaipo_button, self.updateView]

        self.view.on_select_projectlist = [self.updateData, self.select_projectlist, self.updateView]
        self.view.on_select_namelist = [self.updateData, self.select_namelist, self.updateView]
        self.view.on_select_pdflist = [self.updateData, self.select_pdflist, self.updateView]

        self.view.on_change_projectlistop_sort = [self.updateData, self.change_projectlistop_sort, self.updateView]
        self.view.on_change_namelistop_sort = [self.updateData, self.change_namelistop_sort, self.updateView]
        self.view.on_change_pdflistop_sort = [self.updateData, self.change_pdflistop_sort, self.updateView]
        self.view.on_change_setting_sheet = [self.updateData, self.change_setting_sheet, self.updateView]

        self.view.on_change_edit_default_pdf = [self.updateData, self.change_edit_default_pdf, self.updateView]
        self.view.on_change_edit_default_zip = [self.updateData, self.change_edit_default_zip, self.updateView]
        self.view.on_change_edit_send_aipo = [self.updateData, self.change_edit_send_aipo, self.updateView]
        self.view.on_change_setting_genpdf = [self.updateData, self.change_setting_genpdf, self.updateView]

        self.view.on_change_edit_aipo_id = [self.updateData, self.change_edit_aipo_id, self.updateView]

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
    def edit_name(self):
        return self.current_name

    @edit_name.setter
    def edit_name(self, value):
        if not self.current_name:
            return
        # check
        if value.name == "" or value.name is None:
            print("Error : invalid name")
            return
        if value.pdf_filename == "":
            print("Error : invalid pdf_filename")
            return
        if value.zip_filename == "":
            print("Error : invalid zip_filename")
            return
        # correct
        value.pdf_password = value.pdf_password if value.pdf_password != "" else None
        value.zip_password = value.zip_password if value.zip_password != "" else None
        value.pdf_filename = value.pdf_filename if value.pdf_filename.endswith(".pdf") else value.pdf_filename + ".pdf"
        value.zip_filename = value.zip_filename if value.zip_filename.endswith(".zip") else value.zip_filename + ".zip"
        # set
        self.current_name.name = value.name
        self.current_name.pdf_password = value.pdf_password
        self.current_name.zip_password = value.zip_password
        if not self.current_name.is_default_pdf_filename and not value.is_default_pdf_filename or not self.current_name.is_default_pdf_filename and value.is_default_pdf_filename:
            self.current_name.pdf_filename = value.pdf_filename
        if not self.current_name.is_default_zip_filename and not value.is_default_zip_filename or not self.current_name.is_default_zip_filename and value.is_default_zip_filename:
            self.current_name.zip_filename = value.zip_filename
        self.current_name.is_default_pdf_filename = value.is_default_pdf_filename
        self.current_name.is_default_zip_filename = value.is_default_zip_filename
        self.current_name.aipo_id = tryParseInt(value.aipo_id, self.current_name.aipo_id)
        self.current_name.send_aipo = value.send_aipo

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
    def input_password(self):
        if self.current_project is None:
            return ""
        else:
            return self.current_project.input_file.password

    @input_password.setter
    def input_password(self, value):
        if self.current_project is not None:
            self.current_project.input_file.password = value

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
                return "読み込み失敗"
            else:
                return self.current_project.input_file.enable_sheet

    @sheet.setter
    def sheet(self, value):
        if self.current_project is not None:
            if value in self.current_project.input_file.sheets:
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
        # update paramater
        for key, value in self.view.getData().items():
            if key in self.updateData_attrs:
                setattr(self, key, value)
        # other update
        Project.sort(self.projects, self.project_sort)
        if self.current_project is not None:
            Name.sort(self.current_project.name_list, self.name_sort, self.pdf_sort)
        Project.updateProjects(self.projects)

    def updateView(self, *arg):
        # sort projects and names
        Project.sort(self.projects, self.project_sort)
        if self.current_project is not None:
            Name.sort(self.current_project.name_list, self.name_sort, self.pdf_sort)
        # update
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

    def push_projectlistop_duplicate(self):
        log("push_projectlistop_duplicate")
        selected = self.project_list_selected[0]
        dupins = [ins for ins in self.projects if ins.list_index is selected][0]
        self.projects.append(dupins.copy())
        self.view.setInfo('「' + dupins.name + '」を複製しました．')

    def push_projectlistop_new(self):
        log("push_projectlistop_new")
        self.projects.append(Project())

    def push_projectlistop_trash(self):
        log("push_projectlistop_trash")
        selected = self.project_list_selected[0]
        popins = [ins for ins in self.projects if ins.list_index is selected][0]
        #確認する
        result = messagebox.askquestion(title="確認", message=popins.name + "を削除しようとしています。よろしいですか？")
        if result == "yes":
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
        #確認する
        result = messagebox.askquestion(title="確認", message=popins.name + "を削除しようとしています。よろしいですか？")
        if result == "yes":
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
        popins = [ins for ins in self._name_list if ins.pdf_list_index is selected and ins.in_pdf_list][0]
        popins.in_pdf_list = False
        self.view.setInfo('「' + popins.name + '」をPDF化リストから削除しました．')

    def push_edit_save(self):
        if self.current_name is None:
            return
        log("push_edit_save")

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
            self.view.setInfo(self.current_project.input_file.filename + 'というファイルが存在しないか，パスワードが間違っています．')
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

        window = ResultView(self.view.window, self.projects)

        self.view.disableGen()

    def push_sendaipo_button(self):
        log("push_sendaipo_button")
        if self.sendaipo_window:
            self.view.setInfo('二つ同時に開けません')
            return
        if not self.current_project:
            self.view.setInfo('プロジェクトを選んでください')
            return
        self.sendaipo_window = SendSettingView(self.view.window, self.current_project)

        # set listener
        self.sendaipo_window.on_push_cancel_button = [self.SSV_update, self.SSV_push_cancel_button]
        self.sendaipo_window.on_push_send_button = [self.SSV_update, self.SSV_push_send_button]
        self.sendaipo_window.on_change_sendmsg = [self.SSV_update, self.SSV_change_sendmsg]
        self.sendaipo_window.on_close_window = [self.SSV_update, self.SSV_close_window]

    def select_projectlist(self, event):
        if len(self.project_list_selected) is 0:
            return
        log("select_projectlist : ", self.project_list[self.project_list_selected[0]])
        selected = self.project_list_selected[0]
        editins = [ins for ins in self.projects if ins.list_index is selected][0]
        self.current_project = editins
        self.current_name = None

    def select_namelist(self, event):
        if len(self.name_list_selected) is 0:
            return
        log("select_namelist : ", self.name_list[self.name_list_selected[0]])
        selected = self.name_list_selected[0]
        editins = [ins for ins in self._name_list if ins.name_list_index is selected][0]
        self.current_name = editins

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

    def change_edit_default_pdf(self):
        log("change_edit_default_pdf")

    def change_edit_default_zip(self):
        log("change_edit_default_zip")

    def change_edit_send_aipo(self):
        log("change_edit_send_aipo")
        if self.sendaipo_window:
            self.SSV_update()

    def change_setting_genpdf(self):
        log("change_setting_genpdf")

    def change_edit_aipo_id(self, event):
        log("change_edit_aipo_id", str(event))
        if self.sendaipo_window:
            self.SSV_update()

    def close_window(self):
        log("close_window")
        self.exportData()

    # SendSettingView listener
    def SSV_update(self, *arg):
        self.sendaipo_window.update()

    def SSV_push_cancel_button(self):
        log("SendSettingView - push_cancel_button")
        self.sendaipo_window.close_window()
        self.sendaipo_window = None

    def SSV_push_send_button(self):
        log("SendSettingView - push_send_button")
        result = messagebox.askquestion(title="確認", message="メッセージを送信してもよろしいですか？" + ("（デバッグ中．送信されません．）" if DEBUG else ""))
        log(result)
        if result == "yes":
            self.sendaipo_window.project.sendMessage(self.view.setInfo)
            SendResultView(self.view.window, self.sendaipo_window.project)
            self.sendaipo_window.close_window()
            self.sendaipo_window = None

    def SSV_change_sendmsg(self, event):
        log("SendSettingView - change_sendmsg")
        project = self.sendaipo_window.project
        self.sendaipo_window.sendmsg_preview['text'] = Name(name="山田太郎").create_aipo_message(project.aipo_message, searchMonth(project.name))

    def SSV_close_window(self):
        log("SendSettingView - close_window")
        self.sendaipo_window = None

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
    def updateProjects(projects):
        for project in projects:
            project.updateNameList()

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
        genpdf=None,
        aipo_message=None
    ):
        # params (initial value)
        self.name_list = name_list or []
        self.input_file = input_file or ExcelFile()
        self.list_index = list_index or -1
        self.output_file = output_file or ""
        self.offset = offset or (0, 0)
        self.range = range or (3, 3)
        self.genpdf = genpdf or False
        self.aipo_message = aipo_message or "%MONTH%月度給与明細です。\nお疲れ様でした。"

        self.save_attrs = ["name_list", "input_file", "list_index", "output_file", "offset", "range", "genpdf", "aipo_message"]

        # temp
        self.is_success = False
        self.error_message = ""

    def copy(self):
        nlist = []
        for i in range(len(self.name_list)):
            nlist.append(self.name_list[i].copy())
        log(type(self.list_index))
        newins = Project(
            nlist,
            self.input_file.copy(),
            self.list_index,
            self.output_file,
            self.offset,
            self.range,
            self.genpdf,
            self.aipo_message
        )
        return newins

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

    def updateNameList(self):
        for name in self.name_list:
            name.project = self

    def load(self, dict):
        for k, v in dict.items():
            if k == 'name_list':
                value = Name.loadNames(v)
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

    def sendMessage(self, infofunc=None):
        # メッセージを送信する人のリスト
        aipolist = Name.getAipoGroup(self.name_list)

        # プロジェクト名から月を抽出
        month = searchMonth(self.name)

        # 記録の初期化
        Name.resetSuccess(aipolist)

        # デバッグ中は送信しない
        if DEBUG:
            for member in aipolist:
                attachment_file = os.path.join(self.output_file, self.name,  member.zip_filename)
                message = member.create_aipo_message(self.aipo_message, month)
                if not os.path.exists(attachment_file):
                    member.is_success = False
                    member.error_message = "添付ファイルが見つからなかったため，送信できませんでした．"
                    continue
                log(member.name + " - AipoID：" + str(member.aipo_id) + "添付ファイル：" + attachment_file + " メッセージ：" + message)
                #if post_message(jsessionid, member.aipo_id, attachment_file, message):
                if "杉野森" in member.name:
                    member.is_success = True
                    member.error_message = "送信成功"
                else:
                    member.is_success = False
                    member.error_message = "送信中にエラーが発生しました．"
            self.is_success = True
            return True

        # メッセージの送信
        print("Warning!!! : send message on aipo")
        jsessionid = get_aipo_session()
        if not jsessionid:
            infofunc("ネットワークエラーのため，メッセージを送信できませんでした．") if infofunc else 0
            self.is_success = False
            self.error_message = "ネットワークエラーのため，メッセージを送信できませんでした．"
            return False

        jsessionid = aipo_login(jsessionid, username, password)
        if not jsessionid:
            infofunc("ユーザー名またはパスワードが違ったため，メッセージを送信できませんでした．") if infofunc else 0
            self.is_success = False
            self.error_message = "ユーザー名またはパスワードが違ったため，メッセージを送信できませんでした．"
            return False

        for member in aipolist:
            attachment_file = os.path.join(self.output_file, self.name,  member.zip_filename)
            message = member.create_aipo_message(self.aipo_message, month)
            if not os.path.exists(attachment_file):
                member.is_success = False
                member.error_message = "添付ファイルが見つからなかったため，送信できませんでした．"
                continue
            if post_message(jsessionid, member.aipo_id, attachment_file, message):
                member.is_success = True
                member.error_message = "送信成功"
            else:
                member.is_success = False
                member.error_message = "送信中にエラーが発生しました．"

        self.is_success = True

        return True

class Name:
    @staticmethod
    def loadNames(list):
        return [Name().load(data) for data in list]

    @staticmethod
    def toDictNames(names):
        return [name.toDict() for name in names]

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
    def getAipoGroup(names):
        return [name for name in names if name.send_aipo]

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
        name=None,
        name_list_index=None,
        pdf_list_index=None,
        in_pdf_list=None,
        pdf_password=None,
        zip_password=None,
        _pdf_filename=None,
        _zip_filename=None,
        is_default_pdf_filename=True,
        is_default_zip_filename=True,
        aipo_id=None,
        send_aipo=None
    ):
        self.name = name or "新規ユーザー"
        self.name_list_index = name_list_index or -1
        self.pdf_list_index = pdf_list_index or -1
        self.in_pdf_list = in_pdf_list or False
        self.pdf_password = pdf_password
        self.zip_password = zip_password
        self._pdf_filename = _pdf_filename or "pdffile.pdf"
        self._zip_filename = _zip_filename or "zipfile.zip"
        self.is_default_pdf_filename = is_default_pdf_filename
        self.is_default_zip_filename = is_default_zip_filename
        self.aipo_id = aipo_id or -1
        self.send_aipo = send_aipo or False

        self.save_attrs = ['name', 'name_list_index', 'pdf_list_index', 'in_pdf_list', 'pdf_password', 'zip_password', '_pdf_filename', '_zip_filename', 'is_default_pdf_filename', 'is_default_zip_filename', 'aipo_id', 'send_aipo']

        # non param
        self.project = None
        self.is_success = False
        self.error_message = ""

    def copy(self):
        newins = Name(
            self.name,
            self.name_list_index,
            self.pdf_list_index,
            self.in_pdf_list,
            self.pdf_password,
            self.zip_password,
            self._pdf_filename,
            self._zip_filename,
            self.is_default_pdf_filename,
            self.is_default_zip_filename,
            self.aipo_id,
            self.send_aipo
        )
        return newins

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

    @property
    def pdf_filename(self):
        if self.project and self.is_default_pdf_filename:
            #return self.name + self.project.input_file.enable_sheet + "明細.pdf"
            return self.project.input_file.enable_sheet + "（" + self.name + "）.pdf"
        else:
            return self._pdf_filename

    @pdf_filename.setter
    def pdf_filename(self, value):
        self._pdf_filename = value

    @property
    def zip_filename(self):
        if self.project and self.is_default_zip_filename:
            #return self.name + self.project.input_file.enable_sheet + "明細.zip"
            return self.project.input_file.enable_sheet + "（" + self.name + "）.zip"
        else:
            return self._zip_filename

    @zip_filename.setter
    def zip_filename(self, value):
        self._zip_filename = value

    def load(self, dict):
        for k, v in dict.items():
            setattr(self, k, v)
        return self

    def toDict(self):
        data = {x : getattr(self, x) for x in self.save_attrs}
        return data

    def create_aipo_message(self, msg, month):
        return msg.replace('%NAME%', self.name).replace('%MONTH%', month or "X")

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

    def copy(self):
        newins = ExcelFile(
            self.filename,
            self.password
        )
        newins.enable_sheet = self.enable_sheet
        newins.reloadfile()
        newins.reloadSheets()
        return newins

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
                False,
                True,
                None,
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
                        name.error_message = "名前が見つかりませんでした。"
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
                    rawpdffile = os.path.abspath(os.path.join(rawpdf_dir, name.pdf_filename))
                    if os.path.exists(rawpdffile):
                        os.remove(rawpdffile)
                    log("save : ", rawpdffile, ", range : ", print_range)
                    if infofunc:
                        infofunc('「' + name.name + '」のPDFファイルを作成中')
                    sheet.ExportAsFixedFormat(0, rawpdffile)

                    # set password to pdf file
                    encryptfile = os.path.abspath(os.path.join(encrypt_dir, name.pdf_filename))
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
                    #zip_dir = os.path.abspath(os.path.join(save_dir, name.name))
                    zip_dir = save_dir
                    if not os.path.isdir(zip_dir):
                        os.mkdir(zip_dir)
                    zipfile = os.path.abspath(os.path.join(zip_dir, name.zip_filename))
                    if os.path.exists(zipfile):
                        os.remove(zipfile)
                    log("save : ", zipfile)
                    if infofunc:
                        infofunc('「' + name.name + '」をZIPに圧縮中')
                    pyminizip.compress(
                        encryptfile.encode('cp932'), '', zipfile.encode('cp932'), name.zip_password or "", int(0)
                    )
                    name.is_success = True
                    name.error_message = ""
                    if infofunc:
                        infofunc('「' + name.name + '」のファイル生成を完了しました．')
                except:
                    if infofunc:
                        infofunc('「' + name.name + '」のファイル生成に失敗しました．')
                    name.error_message = "ファイル生成中にエラーが発生しました。"
                    is_success &= False
        except Exception as e:
            print('Error : cannot save as pdf.', e)
            if infofunc:
                infofunc('エクセルファイルの処理中に問題が発生しました．')
            is_success &= False
        return is_success
