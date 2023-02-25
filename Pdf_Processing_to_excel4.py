import re 
import pandas as pd
import os
import pdfx

chm = r"CAB - ANALYSE DES MENTIONS R.S - 04.01.2023.pdf"
res = []
def traiter_pdf(res,chm,titre,date):
    pdf = pdfx.PDFx(chm)
    pdf_txt = pdf.get_text()
    rrr = pdf_txt.replace('\n','')
    rrr2 = rrr.replace(' ','')   
    match = re.search(r'\d+(?=Mention)', rrr2)
    Mentions = ''
    Mentions_SAR = ''
    if match:
        Mentions = match.group()
    match1 = re.search(r'\d+(?=S.A.RMoulayRachid)', rrr2)
    if match1:
        Mentions_SAR = match1.group()
    else:
        Mentions_SAR = '-'
    matchlallaoumkoultom = re.search(r'\d+(?=S.A.RLallaOumKeltoum)', rrr2)
    if matchlallaoumkoultom:
        Mentions_Sar_lalla = matchlallaoumkoultom.group()
    else:
        Mentions_Sar_lalla = '-'
        
    match_Moulay_ahmed  = re.search(r'\d+(?=S.AMoulayAhmed)', rrr2)
    if match_Moulay_ahmed:
        Mentions_Moulay_ahmed = match_Moulay_ahmed.group()
    else:
        Mentions_Moulay_ahmed = '-'
    match_Moulay_Abdeslam = re.search(r'\d+(?=S.AMoulayAbdeslam)',rrr2)
    if match_Moulay_Abdeslam:
        Mentions_Moulay_Abdeslam =match_Moulay_Abdeslam.group()
    else:
        Mentions_Moulay_Abdeslam = '-'

    res.append([date,titre,Mentions,Mentions_SAR,Mentions_Sar_lalla,Mentions_Moulay_ahmed,Mentions_Moulay_Abdeslam])


chm = r'C:\Users\HP\Desktop\CAB2023'
res = []
for x in range(1,3):
    rr = '0'+str(x) if len(str(x))==1 else str(x)
    for y in range(1,32):
        y = '0'+str(y) if len(str(y))==1 else str(y)
        path = chm+'\\2023-'+rr+'\\2023'+rr+y
        try:
            files = os.listdir(path)
        except:
            files = None
        if files != None:
            for a in files:
                if 'ANALYSE DES MENTIONS' in a:
                    l = path+'\\'+a
                    pathh = path[::-1]
                    dt = pathh[0:2][::-1]+'/'+pathh[2:4][::-1]+'/'+pathh[4:8][::-1]
                    try:
                        traiter_pdf(res,l,a,dt)
                    except:
                        print(dt)
            
#output = [sub_arr for i,sub_arr in enumerate(res) if sub_arr[1] not in [sub_arr[1] for j,sub_arr in enumerate(res) if i!=j]]
#res = output
df1 = pd.DataFrame(res, columns=["Date","Titre",'Mentions','Mentions_SAR','Mentions_Sar_lalla','Mentions_Moulay_ahmed','Mentions_Moulay_Abdeslam'])
df1.to_excel("Mentions_SAR_CAB_2023.xlsx",index=False)
