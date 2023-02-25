# import the necessary packages
import pandas as pd 
import PyPDF2 
import tabula
import warnings
warnings.filterwarnings('ignore')

pdf = PyPDF2.PdfFileReader('tst.pdf')
pdf.getNumPages()


table = tabula.read_pdf('tst.PDF',pages = 1)
def clean(table,n,x):
    
    l = table[0]
    if len(l.columns) == 3:
        c2 = l.columns[1]
        c3 = l.columns[2]
        n_c = []
        for a,b in zip(l[c3],l[c2]):
            if isinstance(a, str):
                n_c.append(a)
            else:
                n_c.append(b)
        l = l.drop([l.columns[1],l.columns[2]],axis=1)
        l['new'] = n_c
    elif len(l.columns) == 4:
        c2 = l.columns[1]
        c3 = l.columns[2]
        c4 = l.columns[3]
        n_c = []
        for a,b,c in zip(l[c4],l[c2],l[c3]):
            if isinstance(a, str):
                n_c.append(a)
            elif isinstance(b, str) :
                n_c.append(b)
            else:
                n_c.append(c)    
        l = l.drop([l.columns[1],l.columns[2],l.columns[3]],axis=1)
        l['new'] = n_c

    l = l.dropna()
    if x == n :
        l= l.head(l.shape[0] - 2)
    return l

    
for x,y,z in zip(range(5),range(10),range(15)):
    print(x,y,z)
def txt_long(a,b):
    a = a.replace(' ','')
    code = a[0:6]
    date = a[6:10]
    date = date[0:2]+" "+date[2:]
    Libelle = a[10:len(a)-8]
    Valeur = a[-8:]
    Valeur = Valeur[0:2]+' '+Valeur[2:4]+' '+Valeur[4:]
    
    Debit_capitaux_crédit = b.replace(' ','')
    r = [code,date,Libelle,Valeur,Debit_capitaux_crédit]
    return r
    

def txt_normal(a,b):
    try:
        code = a[0:7].replace(' ','')
        date = a[7:13].replace(' ','')
        date = date[0:2]+" "+date[2:]
        Libelle = a[12:len(a)-10]
        if Libelle[0]!=' ':
            Libelle = Libelle[1:]
        Valeur = a[-10:]
        Debit_capitaux_crédit = b.replace(' ','')
        r = [code,date,Libelle,Valeur,Debit_capitaux_crédit]
        return r
    except:
        pass

'''
f = 'test.pdf'
for x in range(1,4):    
    table = tabula.read_pdf(f,pages = x)
    l = clean(table)
    C = l.columns
    for a,b in zip(l[C[0]],l[C[1]]):
        print(txt_normal(a, b))
'''
def extraire_excel(f,name):
    res = []
    pdf = PyPDF2.PdfFileReader(f)
    n = pdf.getNumPages()
    for x in range(1,n+1): 
        table = tabula.read_pdf(f,pages = x)
        l = clean(table,n,x)
        C = l.columns
        if x !=1: 
            a = l.iloc[0][0]
            b = l.iloc[0][1]
            res.append(txt_long(a, b))
            l = l.iloc[1: , :]
        for a,b in zip(l[C[0]],l[C[1]]):
            try:
                res.append(txt_normal(a, b))   
            except:
                pass
        df1 = pd.DataFrame(res, columns=["Code","Date","Libelle","Valeur", "Debit_capitaux_crédit"])
    
        df1.to_excel(name+'.xlsx',index=False)    

        
#extraire_excel('test.pdf')


table = tabula.read_pdf('JANVIER.PDF',pages = 3)
pdf = PyPDF2.PdfFileReader('JANVIER.PDF')
n = pdf.getNumPages()
 
l = clean(table, n, 3)
C = l.columns
a = l.iloc[0][0]
b = l.iloc[0][1]


import os
path = r'C:\Users\HP\Desktop\facture'
files = os.listdir(path)
for x in files:
    extraire_excel(x, x)

extraire_excel('Attijariwafa 3.PDF','Attijariwafa 3')
    

pdf = PyPDF2.PdfFileReader('Attijariwafa 2.pdf')
pdf.getNumPages()


table = tabula.read_pdf('Attijariwafa 2.pdf',pages = 1)


race = [0,5,6,4,7,8,9,10,11]

r = [int(x**3/x*2) for x in race if x>=6]










