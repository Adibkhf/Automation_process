import PyPDF2
import os 
import pandas as pd 
chemain = "CAB - RDP SAR -  16.01.2023.pdf"

pdf_reader = PyPDF2.PdfFileReader(chemain)

n = pdf_reader.getNumPages()

date = chemain[-15:-4].strip()
 
def index(tab,txt):
    index = ''
    for i in range(len(tab)):
        if txt in tab[i].replace(' ',''):
            index = i
            break
    return index
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

def dict_Page(pdf_reader):
    d = {}
    n = pdf_reader.getNumPages()
    for x in range(2,n):
        t =pdf_reader.getPage(x).extractText()     
        if 'Source:' in t.replace(' ','') and 'Pays:' in t.replace(' ',''):
            Source,Pays,Titre1 = traiter(t)
            d[x] = Titre1
            #print(Source)
            #print(Pays)
        
    return d

def Autre_source(pdf_reader):
    rs = []
    n = pdf_reader.getNumPages()
    for x in range(2,n):
        t=pdf_reader.getPage(x).extractText()   
        rt = t.replace('\n','').replace(' ','')
        if 'Autressources:' in rt:
            rs.append(x)
    return rs
        
def Autre_source_index(pdf_reader):
    R = Autre_source(pdf_reader)
    D = dict_Page(pdf_reader)
    R.sort(reverse = True)
    D_autre = {}
    for x in R:
        a_s = x
        if x in D.keys():
            D_autre[a_s] = D[x]
        else:
            x = x-1
            while x not in D.keys():
                x = x-1
            D_autre[a_s] = D[x]
    return D_autre

def f_to_t(c):
    return [c[0], c[1] + c[2], c[3:]]

def Ajouter_lien(pdf_reader,res,x):
    t =pdf_reader.getPage(x).extractText()     
    #rt = t.replace('\n','').replace(' ','') 
    tab = t.split('\n')
    tab = [xa.strip() for xa in tab]
    tab = [xb for xb in tab if xb!='']
    if 'Sources  Pays  Langue' in tab:
        index_Source = tab.index('Sources  Pays  Langue')
        if index_Source+1!=len(tab):
            for x in range(index_Source+1,len(tab)):
                res.append(tab[x])
    elif 'Source  Pays  Langue' in tab:
        index_Source = tab.index('Source  Pays  Langue')
        if index_Source+1!=len(tab):
            for x in range(index_Source+1,len(tab)):
                res.append(tab[x])
        #•print(res)
    else:
        for a in tab:
            res.append(a)
            
res = []
def Exec_pdf(path,a):
    res = []
    pdf_reader = PyPDF2.PdfFileReader(path+'\\'+a)
    date = a[-15:-4].strip()
    #pdf_reader = PyPDF2.PdfFileReader(path)
    D_autre = Autre_source_index(pdf_reader)
    D = dict_Page(pdf_reader)
    D_list = list(D.keys())
    n =pdf_reader.getNumPages()
    D_autre_list = list(D_autre.keys())
    D_autre_list.sort()
    for x in D_autre_list:
        #x = 25 
        r = []
        er = x 
        ttitre = D_autre[x]
        Ajouter_lien(pdf_reader,r, er)
        print(er)
        while er not in D_list and er<n:
            Ajouter_lien(pdf_reader,r, er)
            er = er + 1 
            print(er)
        r = [x.split(' ') for x in r]
        for xa in r:
            xa.append(ttitre)
            xa.append(date)
            xa.append(a)
        res.append(r)
    
    return res



chemain = "CAB - RDP SAR -  03.01.2023.pdf"
a='CAB - RDP SAR -  03.01.2023.pdf'
path  = chemain
res = Exec_pdf(chemain, a)

chm = r'C:\Users\HP\Desktop\CAB2023'
Re = []
Paths = []
for x in range(1,3):
    rr = '0'+str(x) if len(str(x))==1 else str(x)
    for y in range(1,32):
        y = '0'+str(y) if len(str(y))==1 else str(y) 
        path = chm+'\\2023-'+rr+'\\2023'+rr+y
        print(path)
        Paths.append(path)

def trouve_sar(f):
    z = None
    for a in f:
        if 'RDP SAR' in a:
            z = a
    return z

Re = []
for p in Paths:
    try:
        files = os.listdir(p)
    except:
        files = None
    if files != None:
        a = trouve_sar(files)
        if a!=None:
            print(a)
            R = Exec_pdf(p,a)
            Re.append(R)
        
t = []
for x in Re:
    for y in x:
        t.append(y)
t = [xa for xa in t if 'CONFIDENTIEL' not in xa]
t = [xb for xb in t if '' not in xb]
t = [list(filter(lambda y: y != '', xc)) for xc in t]
RRR = []
for x in t:
    for y in x:
        RRR.append(y)
def f_to_t(c):
    return [c[0], c[1] + c[2], c[3],c[4],c[5],c[6]]
t = [xa for xa in RRR if 'CONFIDENTIEL' not in xa]
t = [list(filter(lambda y: y != '', xc)) for xc in t]
tr = []
for x in t:
    if len(x)==7:
        tr.append(f_to_t(x))
    else: 
        tr.append(x)
df1 = pd.DataFrame(tr, columns=['Sources','Pays','Langue','Titre','Date','Fichier','t','r'])
df1 = df1.drop("t", axis='columns')
df1 = df1.drop("r", axis='columns')
df1.to_excel("Autre_Source_CAB.xlsx",index=False)
res_f = []