#icon = https://www.iconfinder.com/icons/5232552/coffee_cup_drink_hot_mug_icon
import os
import PySimpleGUI as sg
import ctypes
import threading
from PIL import Image
from win10toast import ToastNotifier
from pdf2docx import Converter
import docx2pdf
import time
from PyPDF2 import PdfFileReader, PdfFileWriter
from datetime import datetime
from pylovepdf.ilovepdf import ILovePdf

myappid = 'Bean Convert.Bean Converter.Wizard.1'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

today_pdf = datetime.today().strftime("%d-%m-%Y")

tn = ToastNotifier()
global files_list, document_list, progress_num
gif_path = os.path.realpath(__file__).replace("\\main.py","\\introbean.gif")
icon_path = os.path.realpath(__file__).replace("\\main.py","\\cicon.ico")
files_list = ["This area will populate with files after selected.","Feel free to select as many as you want.","Bean Convert will recgonize the file by its type.",
            "","NOTE: Bean Convert currently only supports the following conversions:",
            "PDF =====> DOCX",
            "",
            "More to come in a future update! \U0001f600"]

pdf_convert = []
docx_convert = []

document_list = []

progress_num = 0
progress_files = []

sg.theme("LightBrown4")


for i in range(2000):
    sg.popup_animated(image_source=gif_path, time_between_frames=100, title="Bean Convert", background_color="white")
sg.popup_animated(image_source=None)

sg.popup("                Welcome to Bean Convert!\n\nThis is an open source program in development.\n    For documentation and source code visit:\n   www.zachgoreczny.com\python-projects", icon=icon_path)


layout = [
    [sg.Text("Bean Convert", font="bold 17")],
    [sg.FilesBrowse("Select File or Files", key="in-files", target=(None,None)), sg.Button("Click to Load", key="load")],
    [sg.InputText(key="output-dir"), sg.FolderBrowse("Choose Save Location", target=(2,0))],
    [sg.Listbox(values=files_list,size=(70, 8), key="files-list")],
    [sg.Button("Convert!", key="convert"), sg.Checkbox("Show Files After Conversion?", key="check")],
    [sg.Text("       The Current File will Show Here When Conversion Begins       ", key="current-file")],
    [sg.Text("  {}  /  {}  ".format(progress_num, 0), key="counts"), sg.ProgressBar(max_value=100, orientation='h', size=(50,15), key="progressbar")],
    [sg.Text("v 1.00"), sg.Button("Converter Output", key="coutput", visible=False)],
    [sg.Button("    Exit!    ", key="quit")]
]


window = sg.Window(title="Bean Convert", layout=layout, icon=icon_path, element_justification="c")

files_ele = window["files-list"]
in_files_ele = window["in-files"]
output_dir = window["output-dir"]
progress = window["counts"]
pbar = window["progressbar"]
cfile = window["current-file"]
coutput = window["coutput"]

bad_files = []
good_files = 0
progress_num = 0

def converter(window):
    global bad_files, good_files, progress_num

    progress.update("  {}  /  {}  ".format(progress_num, len(progress_files)))
    for file in docx_convert:
        cv = Converter(file)
        size = os.path.getsize(file)/1024
        print(size)
        if size < float(20480):
            try:
                cfile.update("Current File: {}".format(file))
                splits = file.split("/")
                new_pdf = values["output-dir"]+"/"+splits[-1].replace(".pdf", ".docx")
                cv.convert(new_pdf)
                cv.close()
                good_files = good_files + 1
            except:
                bad_files.append(file)
        else:
            try:
                dirout = values["output-dir"]
                splits = file.split("\\")
                try:
                    fileout = (splits[-2]+"_"+splits[-1].replace(".pdf","")+"_compress_"+today_pdf+".pdf")
                except:
                    fileout = (splits[-1].replace(".pdf","")+"_compress_"+today_pdf+".pdf")
                public_key = "project_public_9356f7da439067e11b4112ac13526d9f_TVgOSc2ab95c720b8040dcfe704e7d75f6b9a"

                lovepdf = ILovePdf(public_key=public_key, verify_ssl=True)
                task = lovepdf.new_task("compress")
                task.add_file(file)
                task.set_output_folder(dirout)
                task.execute()
                task.download()
                task.delete_current_task()

                cfile.update("Current File: {}".format(dirout+fileout))
                splits = file.split("/")
                new_pdf = values["output-dir"]+"/"+splits[-1].replace(".pdf", ".docx")
                cv.convert(new_pdf)
                cv.close()
                good_files = good_files + 1
            except:
                bad_files.append(file)
        progress_num = progress_num + 1
        current_progress = 100*(float(progress_num)/float(len(progress_files)))
        progress.update("  {}  /  {}  ".format(progress_num, len(progress_files)))
        pbar.update(current_progress)


    """
    for file in files_list:
        if file.endswith(".mp4"):
            try:
                splits = file.split("/")
                audio_name = values["output-dir"]+"/"+splits[-1].replace(".mp4", ".mp3")
                new_files.append(audio_name)
                video = mp.VideoFileClip(file)
                audio = video.audio
                audio.write_audiofile(audio_name)
                video.close()
                audio.close()
            except:
                bad_files.append(file)
    
        if file.endswith(".png"):
            try:
                splits = file.split("/")
                png_name = splits[-1]
                img = Image.open(png_name)
                img.save(values["output-dir"]+"/"+splits[-1].replace(".png",".ico"))
            except:
                bad_files.append(file)
        """
    sucessful_files = good_files-len(bad_files)
    tn.show_toast(title="Conversion Complete!", msg="{} files sucessfully converted!\nClick 'Conversion Output' for more info".format(sucessful_files), icon_path=icon_path, duration=9)

    if values["check"] == True:
        os.system("start "+values["output-dir"])
    else:
        pass
    

def complete_win(window):
    final_layout = [
        [sg.Text("{} Files Converted!".format(good_files))],
        [sg.Text("{} Files Failed to Convert. See Below for Details".format(len(bad_files)))],
        [sg.Listbox(values=bad_files, size=(40,5))],
        [sg.Button("Thanks!", key="finclose")]
    ]

    final_window = sg.Window(title="Conversion Complete!", layout=final_layout, finalize=True, icon=icon_path)

    while True:
        event, values = final_window.Read(timeout=300)
        if event == "finclose" or event == sg.WIN_CLOSED:
            final_window.close()
            break


def converter_threading():
    threading.Thread(target=converter, args=(window,), daemon=True).start()

def doc_win():
    doc_layout = [
    [sg.Text("{} Document Files Loaded".format(len(document_list)))],
    [sg.Text("Click each file below and select how you want it converted.")],
    [sg.Listbox(values=document_list, size=(40,6), select_mode="multiple", key="doc_box")],
    [sg.Button("To PDF", key="pdf"), sg.Button("To DOC", key="doc"), sg.Button("To DOCX", key="docx")],
    [sg.Button("All To PDF", key="pdfa"), sg.Button("All To DOC", key="doca"), sg.Button("All To DOCX", key="docxa")],
    [sg.Button("Convert!", key="doc_convert")]
]
    doc_window = sg.Window(title="Document Converter", layout=doc_layout, element_justification="c", icon=icon_path)
    doc_window.finalize()


    while True:
        event, values = doc_window.Read()
        if event == "close" or event == sg.WIN_CLOSED:
            break
        if event == "pdf":
            pdfs = values["doc_box"]
            for item in pdfs:
                pdf_convert.append(item)
        if event == "doc":
            sg.popup("This file type is not yet supported!", title="Whoops!", icon=icon_path)
        if event == "docx":
            for item in values["doc_box"]:
                docx_convert.append(item)
                progress_files.append(item)
        if event == "doc_convert":
            converter_threading()
            doc_window.close()

while True:
    event, values = window.Read(timeout=500)
    if event  == "quit" or event == sg.WIN_CLOSED:
        break
    if event == "load":
        files_list = []
        new_vals = values["in-files"].split(";")
        for item in new_vals:
            if str(item).endswith(".mp4") or str(item).endswith(".png") or str(item).endswith(".pdf"):
                files_list.append(item)
            else:
                sg.popup("Some files could not be added as only .mp4 and .png files are currently supported.", icon=icon_path)
                files_list = ["This area will populate with files after selected.","Feel free to select as many as you want.","Bean Convert will recgonize the file by its type.","","NOTE: Bean Convert currently only supports converting .MP4 to .MP3"]
        sg.popup("Files Loaded: {}".format(len(files_list)))
        files_ele.update(files_list)
    if event == "convert":
        print(values["check"])
        if values["output-dir"] != "":
            photo_list = []
            document_list = []
            video_list = []
            audio_list = []
            for file in files_list:
                if file.endswith(".pdf") or file.endswith(".doc") or file.endswith(".docx"):
                    document_list.append(file)
            if len(document_list) > 0:    
                doc_win()
            sg.popup_auto_close("Conversion started!", auto_close_duration=5, icon=icon_path)
            coutput.update(visible=True)

        else:
            sg.popup("Whoops! Please select an output folder.", icon=icon_path)
    if event == "coutput":
        complete_win(window)
    
        
