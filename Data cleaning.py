import pandas as pd
import spellchecker
import difflib
from langdetect import detect
# Read multiple sheets from the same Excel file
CAB_RS = pd.read_excel("CAB-2023.xlsx", sheet_name="CAB RS")
CAB_PE = pd.read_excel("CAB-2023.xlsx", sheet_name="CAB PE")
FR = pd.read_excel("CAB-2023.xlsx", sheet_name="FR")

#Count mention par date 
s = CAB_RS.resample('D', on='Date')['N. Mentions'].sum()
s.to_excel("Mention par Date.xlsx")

#Nombre de ligne CAB_PE
CAB_PE['Date'] = pd.to_datetime(CAB_PE['Date'])
counts = CAB_PE.groupby('Date').size().reset_index(name='counts')
counts.to_excel("Date_CAB_PE.xlsx",index=False)

GOLF_RS = pd.read_excel("GOLF-2023.xlsx", sheet_name="RS")
GOLF_PE = pd.read_excel("GOLF-2023.xlsx", sheet_name="PE")


spell = spellchecker.SpellChecker(language='fr')

#GOLF
def Nom_propre(D,n):
    r = []
    for x in D[n]:
        if type(x)==str:
            r.append(x[0].upper() + x[1:])
        else:
            r.append(None)
    return r
#p = Nom_propre(GOLF_PE,'Sujet C')
#GOLF_PE['Sujet C'] = p
def espace_delete(D,n):
    r = []
    for x in D[n]:
        if type(x)==str:
            r.append(x.strip())
        else:
            r.append(None)
    return r

def Orthographe(D,n):
    res = []
    for x in D[n]:
        if type(x)==str and detect(x)=='fr':
            r = []
            for a in x.split(' '):
                if len(spell.known([a]))==1:
                    r.append(' '.join(spell.known([a])))
                else:
                    r.append(a)
            res.append(' '.join(r))
        else:
            res.append(x)
    return res

def Correct_similar_sentence(D,n):
    r = D[n].sort_values().reset_index(drop=True)
    D[n] = D[n].sort_values().reset_index(drop=True)
    # Loop through each row of the column
    for i in range(1, len(D)):
        string1 = D.loc[i-1, n]
        string2 = D.loc[i, n]
        similarity  = difflib.SequenceMatcher(None, string1, string2).ratio()
        if similarity > 0.9:  # set your threshold here
            D.loc[i, n] = string1
    return r

def swap_columns(df, col1, col2):
    col_list = list(df.columns)
    x, y = col_list.index(col1), col_list.index(col2)
    col_list[y], col_list[x] = col_list[x], col_list[y]
    df = df[col_list]
    return df

#Delete space THEMATIQUE+SUJET(A+B+C)+TONALITé+Media+MENTION SAR
GOLF_PE = pd.read_excel("GOLF-2023.xlsx", sheet_name="PE")
Colonne_golf_pe=['Sujet A','Sujet B','Sujet C','Thématiques','MENTION SAR','Tonalité']
for x in Colonne_golf_pe:
    GOLF_PE[x] = espace_delete(GOLF_PE,x)
GOLF_RS = pd.read_excel("GOLF-2023.xlsx", sheet_name="RS")
Colonne_golf_rs=['Sujet A','Sujet B','Sujet C','Thématiques','Mention SAR','Tonalité']
for x in Colonne_golf_rs:
    GOLF_RS[x] = espace_delete(GOLF_RS,x)

GOLF_PE['Sujet C'] = Orthographe(GOLF_PE,'Sujet C')
GOLF_PE['Thématiques_bef'] = Correct_similar_sentence(GOLF_PE,'Thématiques')
GOLF_RS['new_thematique'] = Correct_similar_sentence(GOLF_RS,'Thématiques')


#Majuscule au debut de la phrase THEMATIQUE+SUJET(A+B+C)+TONALITé+Media+MENTION SAR
for x in Colonne_golf_pe:
    GOLF_PE[x] = Nom_propre(GOLF_PE,x)
for x in Colonne_golf_rs:
    GOLF_RS[x] = Nom_propre(GOLF_RS,x)
    
#Correction des erreurs THEMATIQUE+SUJET+TONALITé+Media
for x in Colonne_golf_pe:
    GOLF_PE[x] = Orthographe(GOLF_PE,x)
for x in Colonne_golf_rs:
    GOLF_RS[x] = Orthographe(GOLF_RS,x)

#Detect the same phrases
GOLF_RS.to_excel("golf_RS_net.xlsx",index=False)
GOLF_PE.to_excel("golf_PE_net.xlsx",index=False)


ra = []
for x in GOLF_PE['Thématiques']:
    ra.append(''.join([spell.correction(a) for a in x.split(' ')]))

for x,y in zip(GOLF_PE['Thématiques'],ra):
    print(x+':'+y)
    print('-------------------')

t = 'Ines Laklalech A Décroché Une Pleine Catégorie Sur Le Lpga Tour'
r = []
for x in t.split(' '):
    print(x)
    if len(spell.known([x]))==1:
        r.append(' '.join(spell.known([x])))
    else:
        r.append(x)
        
for x in r:
    print(x[0])

#FIFM
#Delete space THEMATIQUE+SUJET(A+B+C)+TONALITé+Media+MENTION SAR

#Majuscule au debut de la phrase THEMATIQUE+SUJET(A+B+C)+TONALITé+Media+MENTION SAR

#Correction des erreurs THEMATIQUE+SUJET+TONALITé+Media

#CAB
#Delete space THEMATIQUE+SUJET(A+B+C)+TONALITé+Media+MENTION SAR

#Majuscule au debut de la phrase THEMATIQUE+SUJET(A+B+C)+TONALITé+Media+MENTION SAR

#Correction des erreurs THEMATIQUE+SUJET+TONALITé+Media

Colonne_CAB_pe=['Sujet A','Sujet B','Sujet C','Thématique','Tonalité']
for x in Colonne_CAB_pe:
    CAB_PE[x] = espace_delete(CAB_PE,x)
CAB_RS = pd.read_excel("CAB-2023.xlsx", sheet_name="CAB RS")
Colonne_CAB_rs=['Sujet A','Sujet B','Sujet C','Thématiques','Tonalité']
for x in Colonne_CAB_rs:
    CAB_RS[x] = espace_delete(CAB_RS,x)

CAB_PE['Thématiques_bef'] = Correct_similar_sentence(CAB_PE,'Thématique')
CAB_RS['new_thematique'] = Correct_similar_sentence(CAB_RS,'Thématiques')


#Majuscule au debut de la phrase THEMATIQUE+SUJET(A+B+C)+TONALITé+Media+MENTION SAR
for x in Colonne_CAB_pe:
    CAB_PE[x] = Nom_propre(CAB_PE,x)
for x in Colonne_CAB_rs:
    CAB_RS[x] = Nom_propre(CAB_RS,x)
    
#Correction des erreurs THEMATIQUE+SUJET+TONALITé+Media
for x in Colonne_CAB_pe:
    CAB_PE[x] = Orthographe(CAB_PE,x)
for x in Colonne_CAB_rs:
    CAB_RS[x] = Orthographe(CAB_RS,x)

CAB_PE = swap_columns(CAB_PE,'Thématiques_bef','Thématique')
CAB_PE = swap_columns(CAB_PE,'Sujet A','Thématique')


CAB_RS = swap_columns(CAB_RS,'new_thematique','Thématiques')
CAB_RS = swap_columns(CAB_RS,'Sujet A','Thématiques')

#Detect the same phrases
CAB_RS.to_excel("CAB_RS_net.xlsx",index=False)
CAB_PE.to_excel("CAB_PE_net.xlsx",index=False)



