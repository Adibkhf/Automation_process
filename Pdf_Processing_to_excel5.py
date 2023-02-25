import pandas as pd
import os
import pdfx
import PyPDF2

df = pd.read_excel('tst.xlsx')

 
ll = 'GOLF - ANALYSE DES MENTIONS R.S 02.10.2020.pdf'


pdf = pdfx.PDFx(ll)
r = pdf.get_references()
rt = pdf.get_text()
rt = rt.split('/n')
rt = [x.strip() for x in rt]
rt = [x for x in rt if x!='']


rrr = rt.replace('\n','')
rrr = rrr.replace(' ','')



with open(ll, "rb") as file:
    pdf = PyPDF2.PdfFileReader(file)
    n = pdf.numPages
    txt = ''
    r = pdf.getPage(0)
    txt += r.extractText()
txt = txt.split(' ')
rt = txt
rt = [x.strip() for x in rt]
rt = [x for x in rt if x!='']
for x in rt:
    if x.startswith('https'):
        #print(x)
        if rt[rt.index(x)+1]!='Date':
            f = rt.index(x)+1
            l = x+rt[f]
            l = l.replace('\n','')

            print(l)
l = l.replace('\n','')

chm = r'C:\Users\HP\Desktop\2021\2021-'
res = []
for x in range(1,13):
    rr = '0'+str(x) if len(str(x))==1 else str(x)
    for y in range(1,32):
        y = '0'+str(y) if len(str(y))==1 else str(y)
        #print(chm+r+'\\2020'+r+y)
        path = chm+rr+'\\2021'+rr+y
        try:
            files = os.listdir(path)
        except:
            files = None
        if files != None:
            for a in files:
                if 'EXTRACTION MENTION' in a:
                    l = path+'\\'+a
                    #pdf = pdfx.PDFx(l)
                    #r = pdf.get_references_as_dict()
                    pathh = path[::-1]
                    dt = pathh[0:2][::-1]+'/'+pathh[2:4][::-1]+'/'+pathh[4:8][::-1]
                    try:
                        with open(l, "rb") as file:
                            n = pdf.numPages
                            pdf = PyPDF2.PdfFileReader(file)
                            txt = ''
                            r = pdf.getPage(0)
                            txt += r.extractText()
                        txt = txt.split(' ')
                        rt = txt
                        rt = [x.strip() for x in rt]
                        rt = [x for x in rt if x!='']
                        index = next((i for i, x in enumerate(rt) if x.startswith("https")), None)
                        print(index)
                        print(rt[index])
                        if rt[index+1]!='Date':
                            f = index+1
                            lll = rt[index]+rt[f]
                            lll = lll.replace('\n','')
                        else:
                            lll = rt[index]
                        res.append([dt,lll,a])
                    except:
                        pass

                elif 'EXTRACTION VIDEOS' in a:
                    l = path+'\\'+a
                    try:
                        pdf = pdfx.PDFx(l)
                        r = pdf.get_references_as_dict()
                        pathh = path[::-1]
                        dt = pathh[0:2][::-1]+'/'+pathh[2:4][::-1]+'/'+pathh[4:8][::-1]
                        for lien in r.values():
                            res.append([dt,max(lien, key=len),a])
                            #print(dt)
                            #print(lien)
                    except:
                        pass
                elif 'GOLF - Extraction- RS' in a:
                    l = path+'\\'+a
                    #pdf = pdfx.PDFx(l)
                    #r = pdf.get_references_as_dict()
                    pathh = path[::-1]
                    dt = pathh[0:2][::-1]+'/'+pathh[2:4][::-1]+'/'+pathh[4:8][::-1]
                    
                    
output = [sub_arr for i,sub_arr in enumerate(res) if sub_arr[1] not in [sub_arr[1] for j,sub_arr in enumerate(res) if i!=j]]
res = output
df1 = pd.DataFrame(res, columns=["Date","liens",'nom'])
df1.to_excel("2021.xlsx")


                
'''        
# Open the PDF file
with open("GOLF - EXTRACTION MENTION 04.04.2020.pdf", "rb") as file:
    pdf = PyPDF2.PdfFileReader(file)
    n = pdf.numPages
    c = 0
    txt = ''
    while c<n:     
        r = pdf.getPage(c)
        txt += r.extractText()
        c+=1
    txt = txt.split('\n')

txt = [x.strip() for x in txt]
txt = [x for x in txt if x !='']
txt = [x for x in txt if x.strip()[0:5] == 'https'][0]
'''

    