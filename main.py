from docxtpl import DocxTemplate, InlineImage, RichText
import qrcode
import docx2pdf

with open("kapitan_namuna.csv","r") as csv_file:
  op=csv_file.readlines()
print(op)

for i in op[1:]:
  raqam=i.split(";")[0]
  fio=i.split(";")[1]
  qr_code=i.split(";")[2]
  q_raqam=i.split(";")[3]
  qr = qrcode.QRCode(version = 6,
                   error_correction=qrcode.constants.ERROR_CORRECT_H,
                   box_size = 2,
                   border = 1,)
  doc=DocxTemplate("kapitan.docx")
  qr.add_data(qr_code)
  qr.make(fit = True)
  img = qr.make_image(fill_color = 'black',
                    back_color = 'white')
  img.save('qr_code.png')
  context={
    'raqam':raqam,
    'fio':RichText(fio.upper(), size=32, color='000000', bold=True),
    'qr_code':InlineImage(doc,"qr_code.png"),
    'q_raqam':RichText(q_raqam, italic=True)
  }

  doc.render(context)
  doc.save('word_natija/'f"{raqam}.docx")
  docx2pdf.convert('word_natija/'f"{raqam}.docx",'pdf_natija/'f"{raqam}.pdf")



#https://stackoverflow.com/questions/68435780/how-to-display-image-in-python-docx-template-docxtpl-django-python
#https://docxtpl.readthedocs.io/en/latest/
#https://nagasudhir.blogspot.com/2021/10/automating-word-files-to-pdf-files.html
#https://docs-python.ru/packages/modul-python-docx-python/modul-docx-template/