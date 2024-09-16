import os
# from datetime import time
from tkinter import *
from tkinter import filedialog, messagebox
from PIL import Image  # ImageDraw, ImageFilter, ImageOps, ImageFont
import fitz  # PyMuPDF
import win32print
import win32api
import time
from pathlib import Path
from glob import glob


def input_dir():
    global input_dir
    pt = r"p:\\XEROX700\\МУРАВЬЁВ\\"
    try:
        input_dir = filedialog.askdirectory(initialdir=pt)
        label2.config(text=input_dir)
        label2.pack()

    except:
        messagebox.showinfo('Предупреждение', 'Укажите INPUT')


def printer(save_pdf):
    PRINTER_DEFAULTS = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}

    pHandle = win32print.OpenPrinter('KONICA MINOLTA C14000Series PS', PRINTER_DEFAULTS)
    # pHandle = win32print.OpenPrinter('KONICA MINOLTA C1070/C1060PS', PRINTER_DEFAULTS)
    properties = win32print.GetPrinter(pHandle, 2)
    pDevModeObj = properties["pDevMode"]

    kolvo_copies = message.get()
    pDevModeObj.Copies = int(kolvo_copies)

    win32print.SetPrinter(pHandle, 2, properties, 0)
    win32api.ShellExecute(0, 'print', save_pdf, '.', '.', 0)

    if kol_vo_pages < 10:
        kol_vo_min = 30
    # elif kol_vo_pages > 50:
    #     kol_vo_min = 20
    else:
        kol_vo_min = 18

    time.sleep(kol_vo_min)

    pDevModeObj.Copies = 1
    win32print.SetPrinter(pHandle, 2, properties, 0)

    return save_pdf


def start():
    global kol_vo_pages, c

    mas_print = []

    # переименование папок
    papkii = os.listdir(input_dir)
    c = 1
    for p in papkii:
        old = os.path.join(input_dir, p)
        new = os.path.join(input_dir, f"0{c}__{p}")
        os.rename(old, new)
        mas_print.append(new)
        c = c + 1

    # список папок для печати
    for papka in mas_print:

        jpgs = sorted(glob(f'{papka}/*.jpg'))  # список jpg в папке для печати
        mas_jpg = jpgs[:-2]

        kol_vo_pages = len(mas_jpg) * 2
        pdf = fitz.open("шаблоны/шаблон.pdf")
        for y in range(1, kol_vo_pages):  # кол-во страниц в пдф
            pdf.new_page(width=907.2, height=1275.6)  # pt

        pages = 0
        for jpg in mas_jpg:

            path = Path(jpg)
            name_jps = path.name
            folder = path.parent.name
            img = Image.open(jpg)
            img_crop_left = img.crop((0, 0, 3602.5, 2598))
            img_crop_left.save("left.jpg")
            img_crop_right = img.crop((3602.5, 0, 7205, 2598))
            img_crop_right.save("right.jpg")

            if "01.jpg" != name_jps:
                if kol_vo_pages != pages:
                    # ЛИЦО листа srA3
                    # определить позицию (верхняя)
                    image_rectangle = fitz.Rect(21.4, 14.4, 886, 638)
                    page = pdf[pages]
                    # добавить изображение в pdf
                    page.insert_image(image_rectangle, filename='left.jpg')
                    # определить позицию (нижняя)
                    image_rectangle = fitz.Rect(21.4, 638, 886, 1261)
                    page.insert_image(image_rectangle, filename='left.jpg')
                    pages = pages + 1

            if kol_vo_pages != pages:
                # ОБОРОТ листа srA3
                # определить позицию (верхняя)
                image_rectangle = fitz.Rect(21.4, 14.4, 886, 638)
                page = pdf[pages]
                page.insert_image(image_rectangle, filename='right.jpg')
                # определить позицию (нижняя)
                image_rectangle = fitz.Rect(21.4, 638, 886, 1261)
                page.insert_image(image_rectangle, filename='right.jpg')
                pages = pages + 1

        # pdf.delete_page(0)
        # k = int(kol_vo_pages - 2)
        pdf.delete_page(-1)
        pdf.delete_page(-1)

        path_save_pdf = f"C:\\Users\\Yurius\\PycharmProjects\\Orlan_print\\в печать\\{folder}__save_pdf.pdf"
        pdf.save(path_save_pdf)
        time.sleep(10)

        if ismarried_print.get() == True:
            print(f"Печатаю:  {folder}__листов: {(int(kol_vo_pages) - 2) / 2}")
            printer(path_save_pdf)

    print("[**INFO**]  Все готово!!\n просто + кол-во листов * кол-во шт.")
    print("---------------------------------------------------")



# Дизайн окна
root = Tk()
root.title("Орланы")
root.geometry("450x150")

label1 = Label(text="Количество шт./2:", font="Arial 9")
label1.place(relx=.4, rely=.6)

message = StringVar()
message.set(1)
message_entry = Entry(textvariable=message)
message_entry.place(relx=.5, rely=.8, width=50, anchor="c")

message_button_input = Button(text="INPUT", command=input_dir)
message_button_input.place(relx=.2, rely=.8, height=40, width=100, anchor="c")

label2 = Label(font="Arial 9")

message_button = Button(text="START", command=start)
message_button.place(relx=.8, rely=.8, height=40, width=100, anchor="c")

ismarried_print = IntVar()
ismarried_print.set(True)
ismarried_checkbutton_print = Checkbutton(text="Печатать...", variable=ismarried_print)
ismarried_checkbutton_print.place(relx=.7, rely=.5, )

root.mainloop()
