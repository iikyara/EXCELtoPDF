# coding: utf-8
import sys
import os
import shutil
import win32com.client
#import pandas as pd
import json
import csv
import tkinter as tk
import tkinter.filedialog as filedialog
from tkinter import ttk
from utils import *

class App:
    def __init__(self):
        self.view = View()
        self.controller = Controller(self.view)

    def run(self):
        self.view.run()

class View:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("EXCELtoPDF")
        self.window.geometry("640x480")
        self.window.iconbitmap(default='icon.ico')

        # list data
        self.namelist_data = tk.StringVar()
        self.pdflist_data = tk.StringVar()

        # combobox data
        self.sortlist = ['昇順', '降順']
        self.namelistop_sort_data = tk.StringVar()
        self.pdflistop_sort_data = tk.StringVar()

        #text box
        self.namebox = None
        self.edit_name = None
        self.setting_input = None
        self.setting_offset_column = None
        self.setting_offset_row = None
        self.setting_range_column = None
        self.setting_range_row = None
        self.setting_output = None

        #button
        self.addbutton = None
        self.namelistop_trash = None
        self.listop_add = None
        self.pdflistop_trash = None
        self.edit_save = None
        self.setting_input_button = None
        self.setting_output_button = None
        self.genpdf_button = None

        #list
        self.namelist = None
        self.pdflist = None

        #combobox
        self.namelistop_sort = None
        self.pdflistop_sort = None

        # listner func
        self.on_push_addbutton = []
        self.on_push_namelistop_trash = []
        self.on_push_listop_add = []
        self.on_push_pdflistop_trash = []
        self.on_push_edit_save = []
        self.on_push_setting_input_button = []
        self.on_push_setting_output_button = []
        self.on_push_genpdf_button = []

        self.on_select_namelist = []
        self.on_select_pdflist = []

        self.on_change_namelistop_sort = []
        self.on_change_pdflistop_sort = []

        self.on_close_window = []

        self.createGUI()

        # set listener
        self.window.protocol("WM_DELETE_WINDOW", self.close_window)
        self.addbutton['command'] = self.push_addbutton
        self.namelistop_trash['command'] = self.push_namelistop_trash
        self.listop_add['command'] = self.push_listop_add
        self.pdflistop_trash['command'] = self.push_pdflistop_trash
        self.edit_save['command'] = self.push_edit_save
        self.setting_input_button['command'] = self.push_setting_input_button
        self.setting_output_button['command'] = self.push_setting_output_button
        self.genpdf_button['command'] = self.push_genpdf_button

        self.namelist.bind('<<ListboxSelect>>', self.select_namelist)
        self.pdflist.bind('<<ListboxSelect>>', self.select_pdflist)

        self.namelistop_sort.bind('<<ComboboxSelected>>', self.change_namelistop_sort)
        self.pdflistop_sort.bind('<<ComboboxSelected>>', self.change_pdflistop_sort)

    def createGUI(self):
        # frames
        self.main_frame = ttk.Frame(self.window, style="Main.TFrame")
        self.left_frame = ttk.Frame(self.main_frame, style="A.TFrame")
        self.right_frame = ttk.Frame(self.main_frame, style="B.TFrame")

        self.main_frame.grid(column=0, row=0, sticky=tk.NSEW)
        self.right_frame.grid(column=1, row=0, sticky=tk.NS, padx=1, pady=1)
        self.left_frame.grid(column=0, row=0, sticky=tk.NSEW, padx=1, pady=1)

        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)
        self.left_frame.columnconfigure(0, weight=1)
        self.left_frame.rowconfigure(3, weight=1)
        self.left_frame.rowconfigure(7, weight=1)
        self.right_frame.columnconfigure(0, weight=1)
        self.right_frame.rowconfigure(0, weight=1)
        self.right_frame.rowconfigure(1, weight=1)
        self.right_frame.rowconfigure(2, weight=1)

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
        #self.namelist_data = ['項目1', '項目2', '項目3', '項目4']
        #self.namelist_elem = tk.StringVar(value=self.namelist_data)
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
        # right operation
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
        #self.pdflist_data = ['項目1', '項目2', '項目3', '項目4']
        #self.pdflist_elem = tk.StringVar(value=self.pdflist_data)
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
        # right operation
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
        self.edit_name_frame = ttk.Label(self.edit_frame)
        self.edit_name_label = ttk.Label(self.edit_name_frame, text="名前：")
        self.edit_name = ttk.Entry(self.edit_name_frame)
        # save button
        self.edit_save = ttk.Button(self.edit_frame, text="変更を保存")

        self.edit_frame.grid(column=0, row=0, sticky=tk.NSEW)
        self.edit_label.grid(column=0, row=0, sticky=tk.W)
        self.edit_name_frame.grid(column=0, row=1, sticky=tk.EW)
        self.edit_name_label.grid(column=0, row=0, sticky=tk.W)
        self.edit_name.grid(column=1, row=0, sticky=tk.EW)
        self.edit_save.grid(column=0, row=2, sticky=tk.E)

        self.edit_frame.columnconfigure(0, weight=1)
        self.edit_name_frame.columnconfigure(1, weight=1)

        # setting
        self.setting_frame = ttk.Frame(self.right_frame)
        self.setting_lable = ttk.Label(self.setting_frame, text="設定")
        # input file
        self.setting_input_label = ttk.Label(self.setting_frame, text="入力ファイル：")
        self.setting_input_frame = ttk.Frame(self.setting_frame)
        self.setting_input = ttk.Entry(self.setting_input_frame)
        self.setting_input_button = ttk.Button(self.setting_input_frame, text="参照", width="5")
        # offset
        self.setting_offset_label = ttk.Label(self.setting_frame, text="ズレ（右×下）：")
        self.setting_offset_frame = ttk.Frame(self.setting_frame)
        self.setting_offset_column = ttk.Entry(self.setting_offset_frame, width=10)
        self.setting_offset_label = ttk.Label(self.setting_offset_frame, text="×")
        self.setting_offset_row = ttk.Entry(self.setting_offset_frame, width=10)
        # range
        self.setting_range_label = ttk.Label(self.setting_frame, text="範囲（横×縦）：")
        self.setting_range_frame = ttk.Frame(self.setting_frame)
        self.setting_range_column = ttk.Entry(self.setting_range_frame, width=10)
        self.setting_range_label = ttk.Label(self.setting_range_frame, text="×")
        self.setting_range_row = ttk.Entry(self.setting_range_frame, width=10)
        # output file
        self.setting_output_label = ttk.Label(self.setting_frame, text="出力ファイル：")
        self.setting_output_frame = ttk.Frame(self.setting_frame)
        self.setting_output = ttk.Entry(self.setting_output_frame)
        self.setting_output_button = ttk.Button(self.setting_output_frame, text="参照", width="5")

        self.setting_frame.grid(column=0, row=1, sticky=tk.NSEW)
        self.setting_lable.grid(column=0, row=0, sticky=tk.W)
        self.setting_input_label.grid(column=0, row=1, sticky=tk.E)
        self.setting_input_frame.grid(column=1, row=1, sticky=tk.EW)
        self.setting_input.grid(column=0, row=0, sticky=tk.EW)
        self.setting_input_button.grid(column=1, row=0, sticky=tk.E)
        self.setting_offset_label.grid(column=0, row=2, sticky=tk.E)
        self.setting_offset_frame.grid(column=1, row=2, sticky=tk.EW)
        self.setting_offset_column.grid(column=0, row=0)
        self.setting_offset_label.grid(column=1, row=0)
        self.setting_offset_row.grid(column=2, row=0)
        self.setting_range_label.grid(column=0, row=3, sticky=tk.E)
        self.setting_range_frame.grid(column=1, row=3, sticky=tk.EW)
        self.setting_range_column.grid(column=0, row=0)
        self.setting_range_label.grid(column=1, row=0)
        self.setting_range_row.grid(column=2, row=0)
        self.setting_output_label.grid(column=0, row=4, sticky=tk.E)
        self.setting_output_frame.grid(column=1, row=4, sticky=tk.EW)
        self.setting_output.grid(column=0, row=0, sticky=tk.EW)
        self.setting_output_button.grid(column=1, row=0, sticky=tk.W)

        self.setting_frame.columnconfigure(1, weight=1)
        self.setting_input_frame.columnconfigure(0, weight=1)
        self.setting_output_frame.columnconfigure(0, weight=1)

        # generate pdf button
        self.genpdf_button = ttk.Button(self.right_frame, text="PDFを出力する")
        self.genpdf_button.grid(column=0, row=2, sticky=tk.NSEW)

        ttk.Style().configure("Main.TFrame", padding=6, relief="flat",   background="#000")
        #ttk.Style().configure("A.TFrame", padding=6, relief="flat",   background="#F00")
        #ttk.Style().configure("B.TFrame", padding=6, relief="flat",  background="#0F0")

    def setEditName(self, name):
        self.edit_name.delete(0, tk.END)
        self.edit_name.insert(tk.END, name)

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

    def push_setting_output_button(self):
        for func in self.on_push_setting_output_button:
            func()

    def push_genpdf_button(self):
        self.namelist['listvariable'] = tk.StringVar(value=[0,1,2])
        for func in self.on_push_genpdf_button:
            func()

    def select_namelist(self, event):
        for func in self.on_select_namelist:
            func(event)

    def select_pdflist(self, event):
        for func in self.on_select_pdflist:
            func(event)

    def change_namelistop_sort(self, event):
        for func in self.on_change_namelistop_sort:
            func(event)

    def change_pdflistop_sort(self, event):
        for func in self.on_change_pdflistop_sort:
            func(event)

    def close_window(self):
        for func in self.on_close_window:
            func()
        self.window.destroy()

    def getData(self):
        return {
            "name_list" : list(self.namelist.get(0, self.namelist.size() - 1)),
            "pdf_list" : list(self.pdflist.get(0, self.pdflist.size() - 1)),
            "name_sort" : self.namelistop_sort_data.get().strip(),
            "pdf_sort" : self.pdflistop_sort_data.get().strip(),
            "input_file" : self.setting_input.get(),
            "output_file" : self.setting_output.get(),
            "offset" : (
                int(self.setting_offset_column.get()),
                int(self.setting_offset_row.get())
            ),
            "range" : (
                int(self.setting_range_column.get()),
                int(self.setting_range_row.get())
            ),
            "name_list_selected" : self.namelist.curselection(),
            "pdf_list_selected" : self.pdflist.curselection()
        }

    def setData(self, name_list, pdf_list, name_sort, pdf_sort, input_file, output_file, offset, range, name_list_selected, pdf_list_selected):
        self.namelist['listvariable'] = tk.StringVar(value=name_list)
        self.pdflist['listvariable'] = tk.StringVar(value=pdf_list)
        self.setting_input.delete(0, tk.END)
        self.setting_input.insert(tk.END, input_file)
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

    def run(self):
        self.window.mainloop()

class Controller:
    def __init__(self, view):
        self.view = view
        # params (initial value)
        self._name_list = []
        self.name_sort = ""
        self.pdf_sort = ""
        self.input_file = ""
        self.output_file = ""
        self.offset = (0, 0)
        self.range = (3, 3)
        # extend params
        self.name_list_selected = []
        self.pdf_list_selected = []
        # attributes
        self.updateData_attrs = ["name_sort", "pdf_sort", "input_file", "output_file", "offset", "range", "name_list_selected", "pdf_list_selected"]
        self.updateView_attrs = ["name_list", "pdf_list", "name_sort", "pdf_sort", "input_file", "output_file", "offset", "range", "name_list_selected", "pdf_list_selected"]
        self.save_attrs = ["_name_list", "name_sort", "pdf_sort", "input_file", "output_file", "offset", "range", "name_list_selected", "pdf_list_selected"]
        # set listener
        self.view.on_push_addbutton = [self.updateData, self.push_addbutton, self.updateView]
        self.view.on_push_namelistop_trash = [self.updateData, self.push_namelistop_trash, self.updateView]
        self.view.on_push_listop_add = [self.updateData, self.push_listop_add, self.updateView]
        self.view.on_push_pdflistop_trash = [self.updateData, self.push_pdflistop_trash, self.updateView]
        self.view.on_push_edit_save = [self.updateData, self.push_edit_save, self.updateView]
        self.view.on_push_setting_input_button = [self.updateData, self.push_setting_input_button, self.updateView]
        self.view.on_push_setting_output_button = [self.updateData, self.push_setting_output_button, self.updateView]
        self.view.on_push_genpdf_button = [self.updateData, self.push_genpdf_button, self.updateView]

        self.view.on_select_namelist = [self.updateData, self.select_namelist, self.updateView]
        self.view.on_select_pdflist = [self.updateData, self.select_pdflist, self.updateView]

        self.view.on_change_namelistop_sort = [self.updateData, self.change_namelistop_sort, self.updateView]
        self.view.on_change_pdflistop_sort = [self.updateData, self.change_pdflistop_sort, self.updateView]

        self.view.on_close_window = [self.updateData, self.close_window]
        # params initialize
        self.loadData()
        self.updateView()

    @property
    def name_list(self):
        return Name.getNameList(self._name_list)

    @property
    def pdf_list(self):
        return Name.getPdfList(self._name_list)

    def updateData(self, *arg):
        for key, value in self.view.getData().items():
            if key in self.updateData_attrs:
                setattr(self, key, value)
        Name.sort(self._name_list, self.name_sort, self.pdf_sort)

    def updateView(self, *arg):
        Name.sort(self._name_list, self.name_sort, self.pdf_sort)
        self.view.setData(**{k : getattr(self, k) for k in self.updateView_attrs})

    def loadData(self, filepath="./userdata.dat"):
        try:
            with open(filepath, mode='r', encoding='cp932', errors='ignore') as f:
                data = json.loads(f.read())
                for key, value in data.items():
                    setattr(self, key, value)
        except:
            print("Error : loadData")
        self._name_list = Name.load(self._name_list)

    def exportData(self, filepath="./userdata.dat"):
        data = {x : getattr(self, x) for x in self.save_attrs}
        data['_name_list'] = Name.toDictNames(self._name_list)
        with open(filepath, mode='w', encoding='cp932', errors='ignore') as f:
            f.write(json.dumps(data))

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

    def push_listop_add(self):
        if len(self.name_list_selected) is 0:
            return
        log("push_listop_add")
        selected = self.name_list_selected[0]
        addins = [ins for ins in self._name_list if ins.name_list_index is selected][0]
        addins.in_pdf_list = True

    def push_pdflistop_trash(self):
        if len(self.pdf_list_selected) is 0:
            return
        log("push_pdflistop_trash")
        selected = self.pdf_list_selected[0]
        addins = [ins for ins in self._name_list if ins.pdf_list_index is selected][0]
        addins.in_pdf_list = False

    def push_edit_save(self):
        if len(self.name_list_selected) is 0:
            return
        log("push_edit_save")
        name = self.view.edit_name.get()
        if name is "" or name is None:
            print("Error : invalid name")
        selected = self.name_list_selected[0]
        editins = [ins for ins in self._name_list if ins.name_list_index is selected][0]
        editins.name = name

    def push_setting_input_button(self):
        log("push_setting_input_button")
        dir = os.path.dirname(self.view.setting_input.get())
        file_type = [('エクセルファイル', '*.xlsx'), ('', '*')]
        files = filedialog.askopenfilenames(filetypes=file_type, initialdir=dir)
        if files is "":
            return
        self.input_file = '"' + '","'.join(files) + '"'

    def push_setting_output_button(self):
        log("push_setting_output_button")
        dir = os.path.dirname(self.view.setting_output.get())
        file_type = [('PDFファイル', '*.pdf'), ('', '*')]
        file = filedialog.asksaveasfilename(filetypes=file_type, initialdir=dir)
        if file is "":
            return
        self.output_file = '"' + file + '"'

    def push_genpdf_button(self):
        log("push_genpdf_button")
        for file in [x.strip('" ') for x in self.input_file.split(",")]:
            if not os.path.exists(file):
                print("Error : file doesn't exist. - ", file)
                continue
            genPDF(file, self.output_file.strip('" '), self.pdf_list, self.offset, self.range)

    def select_namelist(self, event):
        if len(self.name_list_selected) is 0:
            return
        log("select_namelist : ", self.name_list[self.name_list_selected[0]])
        self.view.setEditName(self.name_list[self.name_list_selected[0]])

    def select_pdflist(self, event):
        if len(self.pdf_list_selected) is 0:
            return
        log("select_pdflist : ", self.pdf_list[self.pdf_list_selected[0]])

    def change_namelistop_sort(self, event):
        log("change_namelistop_sort", self.name_sort)

    def change_pdflistop_sort(self, event):
        log("change_pdflistop_sort", self.pdf_sort)

    def close_window(self):
        log("close_window")
        self.exportData()

class Name:
    @staticmethod
    def load(list):
        return [Name(**data) for data in list]

    @staticmethod
    def toDictNames(names):
        save_attrs = ['name', 'name_list_index', 'pdf_list_index', 'in_pdf_list']
        return [{key : getattr(name, key) for key in save_attrs} for name in names]

    @staticmethod
    def getNameList(names):
        return [name.name for name in sorted(names, key=lambda ins : ins.name_list_index)]

    @staticmethod
    def getPdfList(names):
        pdflist = [name for name in names if name.in_pdf_list]
        return [name.name for name in sorted(pdflist, key=lambda ins : ins.pdf_list_index)]

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

    def __init__(self,
        name,
        name_list_index=None,
        pdf_list_index=None,
        in_pdf_list=None
    ):
        self.name = name
        self.name_list_index = name_list_index or -1
        self.pdf_list_index = pdf_list_index or -1
        self.in_pdf_list = in_pdf_list or False

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

def genPDF(xlsx_file, pdf_file, name_list, offset, range):
    excel = None
    #コピーファイルを一時的に作成
    tmpfile = os.path.join(os.path.dirname(pdf_file), "_" + os.path.basename(xlsx_file))
    shutil.copyfile(xlsx_file, tmpfile)
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
    except Exception as e:
        print("Error : Don't run excel com. Details - ", e)
    if excel is None:
        print('Error : No app found')
        return
    try:
        wb = excel.Workbooks.Open(tmpfile, None, True)
        excel.Visible = False
        for sheet in wb.Worksheets:
            # erase print range
            sheet.Activate()
            for name in name_list:
                sheet.ResetAllPageBreaks()
                result = sheet.UsedRange.Find(name)
                if result is None:
                    continue
                print(offset, range)
                print(result.Row, ',', result.Column)
                upperleft = sheet.Cells(
                    result.Row + offset[1],
                    result.Column + offset[0]
                )
                bottomright = sheet.Cells(
                    result.Row + offset[1] + range[1] - 1,
                    result.Column + offset[0] + range[0] - 1
                )
                print_range = upperleft.Address + ":" + bottomright.Address
                sheet.PageSetup.PrintArea = print_range
                filename = os.path.join(os.path.dirname(pdf_file), name.replace(" ", "") + ".pdf")
                if os.path.exists(filename):
                    os.remove(filename)
                log("save : ", filename, ", range : ", print_range)
                sheet.ExportAsFixedFormat(0, filename)
    except Exception as e:
        print('Error : cannot save as pdf.', e)
    finally:
        wb.Close(False)
        excel.Quit()
    #コピーファイルの削除
    os.remove(tmpfile)
