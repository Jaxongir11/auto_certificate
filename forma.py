# kutubxonalar ro‘yxati
import tkinter as tk
from tkinter import ttk
from tkinter import *
from PIL import Image, ImageTk
import tkinter.font as font
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from docxtpl import DocxTemplate, InlineImage, RichText
import qrcode
import docx2pdf
import os
#import time
#from tkinter .ttk import Progressbar

#Initialize the directory name
word_papka = "sertifikatlar_word"
pdf_papka="sertifikatlar_pdf"
#Check the directory name exist or not
if os.path.isdir(word_papka) == False:
    #Create the directory
    os.mkdir(word_papka)
else:
    pass

if os.path.isdir(pdf_papka) == False:
    #Create the directory
    os.mkdir(pdf_papka)
else:
    pass
# forma hosil qilish
class Window(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.pack(fill=BOTH, expand=1)
        load = Image.open("iiv.png")
        load=load.resize((100, 100), Image.ANTIALIAS)
        render = ImageTk.PhotoImage(load)
        img = Label(self, image=render)
        img.image = render
        img.place(x=window_width/2-50, y=10)

# csv faylni yuklash
def csv_fayl_ochish():
    csv_file_turi = (
        ('text files', '*.csv'),
        ('All files', '*.*')
    )
    csv_fayl = fd.askopenfilename(
    title='csv faylni tanlang',
    initialdir='/',
    filetypes=csv_file_turi)
    showinfo(title='Tanlangan .csv fayl',message=csv_fayl)
    global csv_file_name
    csv_file_name=csv_fayl
    return csv_file_name

# docx faylni yuklash
def docx_fayl_ochish():
    docx_file_turi = (
        ('text files', '*.docx'),
        ('All files', '*.*')
    )
    docx_fayl = fd.askopenfilename(
    title='Shablon docx faylni tanlang',
    initialdir='/',
    filetypes=docx_file_turi)
    showinfo(title='Tanlangan .docx fayl',message=docx_fayl)
    global docx_file_name
    docx_file_name=docx_fayl
    return docx_file_name

# enable button1
def Switch():
    if(sertifikat_var.get()>0):
        button1.config(state='normal')
    else:
        button1.config(state='disabled')
    docx_fayl_ochish()
# sertifikat tayyorlash funksiyasi
def Tayyorlash():
    #a=sertifikat_var.get()
    #print(docx_file_name)
    #print(csv_file_name)
    #loading ...
    #master=tk.Toplevel(root)
    #master.title("Iltimos biroz kuting ...")
    #bar=Progressbar(master,orient=HORIZONTAL,length=100,mode='determinate')
    #persentLabel=Label(master,text='0%')

    #global csv_file_name
    csv_file=open(csv_file_name,"r")
    op=csv_file.readlines()
    #sertifikatlar_soni=len(op)-1

    for i in op[1:]:
        ser_nomi=i.split(";")[0]
        nomeri=i.split(";")[1]
        fio=i.split(";")[2]
        qr_code=i.split(";")[3]
        raqami=i.split(";")[4]

        qr = qrcode.QRCode(version = 6,
                   error_correction=qrcode.constants.ERROR_CORRECT_H,
                   box_size = 2,
                   border = 1,)
        doc=DocxTemplate(docx_file_name)
        qr.add_data(qr_code)
        qr.make(fit = True)
        img = qr.make_image(fill_color = 'black', back_color = 'white')
        img.save('qr_code.png')

        context={
                 'nomeri':nomeri,
                 'fio':RichText(fio.upper(), size=32, color='000000', bold=True),
                 'qr_code':InlineImage(doc,"qr_code.png"),
                 'raqami':RichText(raqami, italic=True)
                }

        doc.render(context)
        doc.save('sertifikatlar_word/'f"{ser_nomi}.docx")
        docx2pdf.convert('sertifikatlar_word/'f"{ser_nomi}.docx",'sertifikatlar_pdf/'f"{ser_nomi}.pdf")

        #loading ...

        #bar['value']+=100/sertifikatlar_soni
        #persentLabel['text']=int(bar['value']),'%'
        #time.sleep(0.05)
        #master.update_idletasks()
        #bar.pack(padx=10,pady=10)
        #persentLabel.pack()
        #master.mainloop()
    #master.after(3000,lambda:master.destroy())
    showinfo(title='Xabar',message='Muvaffaqiyatli bajarildi!!!')

# sertifikatlar ro‘yxati
sertifikatlar={
    1: 'Malaka oshirish',
    2: 'Qayta tayyorlash',
    3: 'Boshlang‘ich kasbiy tayyorgarlik',
    4: 'Podpolkovnik',
    5: 'Mayor',
    6: 'Kapitan',
    7: 'Katta serjant',
}

# boshlang‘ich parametrlar
root = tk.Tk()
sertifikat_var=tk.IntVar()
root.title('Sertifikat dasturi')
window_width=800
window_height=650
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# menyu
menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Faylni yuklash", command=csv_fayl_ochish)
filemenu.add_separator()
filemenu.add_command(label="Chiqish", command=root.quit)
menubar.add_cascade(label="Fayl", menu=filemenu)

# find the center point
center_x = int(screen_width/2 - window_width / 2)
center_y = int(screen_height/2 - window_height / 2)
root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')

# logotipni yuklab olish
root.iconbitmap('python.ico')
window = Window(root)

# tashkilot nomi
label3 = ttk.Label(text="O‘zbekiston Respublikasi Ichki ishlar vazirligi",font=("Times", 24))
label4 = ttk.Label(text="Malaka oshirish instituti",font=("Times", 24))
label3.pack(anchor=CENTER)
label4.pack(anchor=CENTER)

frame_top=LabelFrame(root,text="Sertifikatlardan birini tanlang:",font="Times 16", labelanchor='n', bd=5,highlightbackground='blue',padx=40)

# radiobutton
for sertifikat in sorted(sertifikatlar):
    tk.Radiobutton(frame_top,text=sertifikatlar[sertifikat],font="Times 14", variable=sertifikat_var,value=sertifikat,command=Switch).pack(anchor=tk.W, side = TOP, ipady = 3)

# butto1 tugmasi
buttonFont = font.Font(family='Helvetica', size=16, weight='bold')
button1=Button(frame_top,text='Tayyorlash',font=buttonFont, state='disabled',command=Tayyorlash)

frame_top.pack(side=TOP,padx=100,pady=20)
button1.pack(pady=10)
root.resizable(False, False)
root.config(menu=menubar)

# mualliflik huquqi
copyright = u"\u00A9"
label2 = ttk.Label(text=copyright+ " IIV Malaka oshirish instituti 2023",font="Times 12")
label2.pack(anchor='s')
root.mainloop()


#https://pythonbasics.org/tkinter-image/
#https://www.plus2net.com/python/tkinter-text-editor.php
#https://stackoverflow.com/questions/68435780/how-to-display-image-in-python-docx-template-docxtpl-django-python
#https://docxtpl.readthedocs.io/en/latest/
#https://nagasudhir.blogspot.com/2021/10/automating-word-files-to-pdf-files.html
#https://docs-python.ru/packages/modul-python-docx-python/modul-docx-template/
#https://linuxhint.com/create-directory-python/