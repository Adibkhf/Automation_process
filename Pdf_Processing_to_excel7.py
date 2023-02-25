import re 
import pandas as pd
import os
import pdfx
import PyPDF2

chemain = "CAB - RDP Divers -27.01.2023.pdf"
pdf = pdfx.PDFx(chemain)
pdf_txt = pdf.get_text()
pdf_txt = pdf_txt.split('\n')
pdf_reader = PyPDF2.PdfFileReader(chemain)
res = pdf_reader.getPage(1).extractText()

from pdf2image import convert_from_path
import pytesseract

pages = convert_from_path('RDP - Mehdi Bensaid- 29.01.2023.pdf', dpi=200,poppler_path=r'C:\Users\HP\Desktop\poppler-0.68.0\bin')
pages[1].save('page.jpg', 'JPEG')

import cv2
img = cv2.imread("page.jpg")
tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
dt_image = pytesseract.image_to_string(img,lang='fra',config='--psm 4 --oem 3')


text = "Afrique: L'Association nationale des médias et \
des éditeurs célèbre l'exploit de l'équipe \
nationale au Mondial du Qatar  ..........................  3 \
وزيران يتفقدان أشغال ترميم مسجد تنمل الموحدي بالحوز ...... 6 \
  اتحاد الصحفيين العرب يهنئ نقابة المغرب بمناسبة مرور60    عاماً\
على تأسيسها  ..........................................................  11 \
  وفاة هذه الفنانة الشهيرة يفجّر ضجّة.. لن تصدّق من هي وكيف\
ماتت بطريقة مأساوية صعقت الجميع!! ........................... 13 "

# use a regular expression to extract the titles and page numbers
matches = re.findall(r'^(.+?)\s+\.{3,}\s+(\d+)', res, re.MULTILINE | re.DOTALL)

# print the extracted titles and page numbers
for title, page_number in matches:
    print(f'Title: {title}')
    print(f'Page number: {page_number}')
    
res = pdf_reader.getPage(2).extractText()
    

    
#if re.search(r'Autres sources  : ',res):
 
def index(tab,txt):
    index = ''
    for i in range(len(tab)):
        if txt in tab[i].replace(' ',''):
            index = i
            break
    return index
    
res = pdf_reader.getPage(3).extractText()


def titre(tab,index_src):
    r = ''
    for x in range(index_src):
        if 'CONFIDENTIEL' not in tab[x]:
            r+=tab[x]
    r = r.strip()
    while r[0].isdigit():
        r = r[1:]
    return r.strip()


def traiter(res):
    #♦res = res.replace(' ','')
    tab = res.split('\n')
    #tab = [x.strip() for x in tab]
    tab = [x for x in tab if x!='']
    src = index(tab,'Source:')
    Pys = index(tab,'Pays:') 
    return tab[src],tab[Pys],titre(tab,src)

traiter(res)           
        
t =pdf_reader.getPage(3).extractText()     
#traiter(t)

def Exec_pdf(chm,res,a):
    pdf_reader = PyPDF2.PdfFileReader(chm+"\\"+a)
    n = pdf_reader.getNumPages()
    for x in range(2,n):
        t =pdf_reader.getPage(x).extractText()     
        if 'Source:' in t.replace(' ','') and 'Pays:' in t.replace(' ',''):
            Source,Pays,Titre = traiter(t)
            #print(Titre)
            #print(Source)
            #print(Pays)
            #print('--------------------------')
            Source = Source[8:].strip()
            Pays = Pays[6:].strip()
            if Source[0]==':':
                Source = Source[1:]
            if Pays[0]==':':
                Pays = Pays[1:]
            dt = a.replace(' ','')[-14:-4]
            res.append([Titre,Source,Pays,a,dt])

chm = r'C:\Users\HP\Desktop\2023'
res = []
for x in range(1,3):
    rr = '0'+str(x) if len(str(x))==1 else str(x)
    for y in range(1,32):
        y = '0'+str(y) if len(str(y))==1 else str(y) 
        path = chm+'\\2023-'+rr+'\\2023'+rr+y
        print(path)
        try:
            files = os.listdir(path)
        except:
            files = None
        if files != None:
            for a in files:
                if 'RDP' in a:
                    Exec_pdf(path,res,a)
                    
df1 = pd.DataFrame(res, columns=["Titre",'Source','Pays','Nom du document','Date'])
df1.to_excel("Lien_CAB_2023.xlsx",index=False)                    
    
    