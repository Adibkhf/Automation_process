import PyPDF2
import more_itertools as mit
import pdfx
import os
import pandas as pd

def index_mot(tab,txt):
    index = ''
    for i in range(len(tab)):
        if tab[i] == txt :
            index = i
            break
    return index



def split_into_tabs(tab):
    
    result = []
    sub_list = []
    for item in tab:
        if item == "Accéder à l'article" or item == "Accéder à l'article   Mehdi Bensaid":
            result.append(sub_list)
            sub_list = []
        else:
            sub_list.append(item)

    result.append(sub_list)
    return result

def Traiter_page(pdf_reader,p):
    t2=pdf_reader.getPage(p).extractText()   
    t2 = t2.split('\n')
    t2 = [x for x in t2 if x!=' ']
    t2 = [x.strip() for x in t2]
    t2 = [x for x in t2 if "CONFIDENTIEL" not in x]    
    #R = split_into_tabs(t2,"Accéder à l'article")
    return t2


def index(tab,txt):
    index = ''
    for i in range(len(tab)):
        if txt in tab[i]:
            index = i
            break
    return index
        
def Titre(tab):
    index_media = index(tab,'Média')
    r = tab[0:index_media]
    txt = ''
    for x in r:
        txt += x
    return txt.strip()


def Media(tab):
    r = ''
    for x in tab:
        if 'Média' in x:
            r = x
            break
    re = r.split(':')
    re = re[1].replace(' ','')
    return re.strip()

def Type(tab):
    r = ''
    for x in tab:
        if 'Type' in x:
            r = x
            break
    re = r.split(':')
    return re[1].strip()

def Tonalité(tab):
    rrr = ''
    for x in tab:
        if 'Tonalité' in x:
            rrr = x
            break
    re = rrr.split(':')
    re = re[1].strip()
    re = re.replace(' ','')
    if 'Négative' in re:
        return 'Négative'
    elif 'Neutre' in re:
        return 'Neutre'
    else:
        if'Positif' or 'Positive' in re:
            return 'Positive'

    
def Lien(pdf):
    pdf_core = pdfx.PDFx(pdf)
    r = pdf_core.get_references_as_dict()
    r =[x for x in r['url'] if x[0:4]!='file']
    r = [x for x in r if x[0:4]=='http'] 
    #Lien hypertext pdf    
    return r
    

#chemain = "RDP - MJCC_Mehdi Bensaid Langue Arabe - 07.02.2023.pdf"
#pdf_reader = PyPDF2.PdfFileReader(chemain)
#n = pdf_reader.getNumPages()
####

#t2=pdf_reader.getPage(5).extractText()
#t22 = t2.replace('\n','')
#t22 = t22.replace(' ','')

def Count_Mehdi_MJCC(pdf_reader):
    R_f = []
    M_B = None
    MJCC = None
    n = pdf_reader.getNumPages()
    for x in range(2,n):
        t2=pdf_reader.getPage(x).extractText()
        t22 = t2.replace('\n','')
        t22 = t22.replace(' ','')
        if "MehdiBensaid" == t22[-12:]:
            M_B = x 
        if 'MJCC' == t22[-4:]:
            MJCC = x
    M_B_text = ''
    MJCC_text = ''
    if M_B!=None and MJCC != None:
        if M_B<MJCC:
            for x in range(M_B,MJCC):
                M_B_text+=pdf_reader.getPage(x).extractText()
            for x in range(MJCC,n):
                MJCC_text+=pdf_reader.getPage(x).extractText()
        else:
            for x in range(MJCC,M_B):
                MJCC_text+=pdf_reader.getPage(x).extractText()
            for x in range(M_B,n):
                M_B_text+=pdf_reader.getPage(x).extractText()    
        M_B_count = M_B_text.count("Accéder à l'article")
        MJCC_count = MJCC_text.count("Accéder à l'article")    
        if M_B<MJCC:
            for x in range(M_B_count):
                R_f.append('Mehdi Bensaid')
            for x in range(MJCC_count):
                R_f.append('MJCC')
        else:
            for x in range(MJCC_count):
                R_f.append('MJCC')
            for x in range(M_B_count):
                R_f.append('Mehdi Bensaid')
    elif M_B==None and MJCC != None:
        for x in range(MJCC,n):
            MJCC_text += pdf_reader.getPage(x).extractText()
        MJCC_text = MJCC_text.replace(' ','')
        MJCC_count = MJCC_text.count("Accéder à l'article")
        for x in range(MJCC_count):
            R_f.append("MJCC")
    elif M_B!=None and MJCC == None:
        for x in range(M_B,n):
            M_B_text += pdf_reader.getPage(x).extractText()
        M_B_text = M_B_text.replace(' ','')
        M_B_count = M_B_text.count("Accéderàl'article")
        for x in range(M_B_count):
            R_f.append("Mehdi Bensaid")
    return R_f
        
    
def Exec(p,a):
    date = a[-15:-4].strip()

    chmm = p+'\\'+a
    pdf_reader = PyPDF2.PdfFileReader(chmm)
    tableu_c = Count_Mehdi_MJCC(pdf_reader)
    n = pdf_reader.getNumPages()
    res = []
    for p in range(2,n):
        res.append(Traiter_page(pdf_reader, p))
    r = []
    for x in res :
        for y in x:
            r.append(y)
    R = split_into_tabs(r)
    R = [x for x in R if len(x)>5]
    for lst in R:
        if lst[0] == 'MJCC' or lst[0] == 'Mehdi Bensaid':
            del lst[0]
    
    TR = []
    for aa in R:
        TR.append([Titre(aa),Media(aa),Type(aa),Tonalité(aa),a,date])

    for x,y in zip(TR,tableu_c):
        x.append(y)
        
    TTT = Lien(chmm)    
    for x in TR:
        for url in TTT:
            if x[1].split('.')[0] in url:
                x.append(url)
                break
    return TR
            

#T = Exec(r'C:\Users\HP\Desktop\2023-02\20230205','RDP - MJCC _ Mehdi Bensaid- Langues étrangères 05.02.2023.pdf')
#p = r'C:\Users\HP\Desktop\2023-02\20230212'
#a = 'RDP - MJCC _ Mehdi Bensaid- Langues étrangères 12.02.2023.pdf'
#chm = p+'\\'+a
#pdf_reader = PyPDF2.PdfFileReader(chm)
#l = Lien(chm)
#E = Exec(p,a)

chm = r'C:\Users\HP\Desktop\2023-02'
Re = []
Paths = []
for y in range(5,22):
    if len(str(y))==1:
        path = chm+'\\2023020'+str(y)
    else:
        path = chm+'\\202302'+str(y)
    print(path)
    Paths.append(path)    
    


Re = []
for p in Paths:
    try:
        files = os.listdir(p)
    except:
        files = None
    if files != None:
        for a in files:
            if 'RDP - MJCC' in a:
                print(p)
                Re+= Exec(p,a)
          
                
df1 = pd.DataFrame(Re, columns=['Titre','Source','Type','Tonalité','Fichier','Date','Thématique','Lien'])
def swap_columns(df, col1, col2):
    col_list = list(df.columns)
    x, y = col_list.index(col1), col_list.index(col2)
    col_list[y], col_list[x] = col_list[x], col_list[y]
    df = df[col_list]
    return df

df1 = swap_columns(df1,'Fichier','Lien')
df1 = swap_columns(df1,'Titre','Date')
df1 = swap_columns(df1,'Source','Titre')

df1.to_excel("MJCC_MB.xlsx",index=False)

