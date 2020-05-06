import os
import re
from settings import *

print("utils - DEBUG :", DEBUG)

def log(*arg, **darg):
    if DEBUG:
        print(*arg, **darg)

def tryParseInt(src, default=0):
    return tryParse(src, int, default)

def tryParseFloat(src, default=0.0):
    return tryParse(src, float, default)

def tryParseStr(src, default=""):
    return tryParse(src, str, default)

def tryParse(src, type, default=0):
    try:
        return type(src)
    except:
        return default

def tryRemoveFile(filename):
    try:
        os.remove(filename)
    except Exception as e:
        log('tryRemoveFile - ', e)

def searchMonth(str):
    return (re.compile(r'([0-9]+)月').findall(str) or [''])[0]

def getSheetsName(xlsx_files, infofunc=None):
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
    except Exception as e:
        log("Error : Don't run excel com. Details - ", e)
    if excel is None:
        log('Error : No app found')
        return
    try:
        wb = excel.Workbooks.Open(tmpfile, None, True)
        excel.Visible = False
        for sheet in wb.Worksheets:
            pass
    except Exception as e:
        log('Error : cannot save as pdf.', e)
        infofunc or infofunc('エクセルファイルの処理中に問題が発生しました．')
        is_success = False
        #raise e
    finally:
        wb.Close(False)
        excel.Quit()

# pdf_dir
# |- rawpdf
# |  |- pdffile1.pdf
# |  |- pdffile2.pdf
# |  |-      :
# |- encrypt
# |  |- encrypt1.pdf
# |  |- encrypt2.pdf
# |  |-      :
# |- result1.zip
# |- result2.zip
# |-    :
def genPDF(xlsx_file, pdf_file, name_list, offset, range, infofunc):
    excel = None
    is_success = True
    rawpdf_dir = os.path.abspath(os.path.join(pdf_file, "rawpdf"))
    encrypt_dir = os.path.abspath(os.path.join(pdf_file, "encrypt"))
    if not os.path.isdir(rawpdf_dir):
        os.mkdir(rawpdf_dir)
    if not os.path.isdir(encrypt_dir):
        os.mkdir(encrypt_dir)
    #コピーファイルを一時的に作成
    #tmpfile = os.path.join(os.path.dirname(pdf_file), "_" + os.path.basename(xlsx_file))
    tmpfile = os.path.abspath(os.path.join(pdf_file, "_" + os.path.basename(xlsx_file)))
    shutil.copyfile(xlsx_file, tmpfile)
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
    except Exception as e:
        log("Error : Don't run excel com. Details - ", e)
    if excel is None:
        log('Error : No app found')
        return
    try:
        wb = excel.Workbooks.Open(tmpfile, None, True)
        excel.Visible = False
        for sheet in wb.Worksheets:
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
                    #rawpdffile = os.path.join(os.path.dirname(pdf_file), name.replace(" ", "") + ".pdf")
                    # save as pdf file
                    rawpdffile = os.path.abspath(os.path.join(rawpdf_dir, name.name.replace(" ", "") + ".pdf"))
                    if os.path.exists(rawpdffile):
                        os.remove(rawpdffile)
                    log("save : ", rawpdffile, ", range : ", print_range)
                    infofunc or infofunc('「' + name.name + '」のPDFファイルを作成中')
                    sheet.ExportAsFixedFormat(0, rawpdffile)
                    # set password to pdf file
                    encryptfile = os.path.abspath(os.path.join(encrypt_dir, name.name.replace(" ", "") + ".pdf"))
                    if os.path.exists(encryptfile):
                        os.remove(encryptfile)
                    log("save : ", encryptfile)
                    infofunc or infofunc('「' + name.name + '」のPDFファイルを暗号化中')
                    rawpdf = Pdf.open(rawpdffile)
                    encryptpdf = Pdf.new()
                    encryptpdf.pages.extend(rawpdf.pages)
                    encryptpdf.save(encryptfile, encryption=pikepdf.Encryption(
                        user=name.pdf_password or "", owner=name.pdf_password or ""
                    ))
                    rawpdf.close()
                    encryptpdf.close()
                    # create zip file
                    zip_dir = os.path.abspath(os.path.join(pdf_file, name.name))
                    if not os.path.isdir(zip_dir):
                        os.mkdir(zip_dir)
                    zipfile = os.path.abspath(os.path.join(zip_dir, name.zip_filename.replace(" ", "")))
                    if os.path.exists(zipfile):
                        os.remove(zipfile)
                    log("save : ", zipfile)
                    infofunc or infofunc('「' + name.name + '」をZIPに圧縮中')
                    pyminizip.compress(
                        encryptfile.encode('cp932'), '', zipfile.encode('cp932'), name.zip_password or "", int(0)
                    )
                    name.is_success = True
                    infofunc or infofunc('「' + name.name + '」のファイル生成を完了しました．')
                except:
                    infofunc or infofunc('「' + name.name + '」のファイル生成に失敗しました．')
                    is_success = False
    except Exception as e:
        log('Error : cannot save as pdf.', e)
        infofunc or infofunc('エクセルファイルの処理中に問題が発生しました．')
        is_success = False
        #raise e
    finally:
        wb.Close(False)
        excel.Quit()
    #コピーファイルの削除
    os.remove(tmpfile)
    return is_success
