from pytube import YouTube
import pandas as pd
import os

# specify the URL of the video you want to download
url = 'https://www.youtube.com/watch?v=jNQXAC9IVRw'


# get the video with the highest resolution

def download(url,dst):
    yt = YouTube(url)

# get the video with the highest resolution
    stream = yt.streams.filter(progressive=True, file_extension='mp4').order_by('resolution').desc().first()

# download the video
    stream.download(dst)
    print(yt.title)


df = pd.read_excel('GOLF ytb 2020.xlsx')

# extract month from datetime column
df['month'] = df['Date'].dt.month

# group dataframe by month
monthly_groups = df.groupby('month')

# create a folder for each month and save data for that month to excel
for name, group in monthly_groups:
    os.makedirs(f"{name}", exist_ok=True)
    group.to_excel(f"{name}/pos_neutre.xlsx", index=False)

tonalite_group = df.groupby('Tonalité')
for name, group in tonalite_group:
    print(group)



chm = r'C:/Users/HP/Desktop/GOLF ytb 2020'
c = 0
for x in range(1,13):
    #print(chm+'/'+str(x)+'/'+'neg.xlsx')
    #print(chm+'/'+str(x)+'/'+'pos_neutre.xlsx')
    
    try:
        df = pd.read_excel(chm+'/'+str(x)+'/'+'pos_neutre.xlsx')
        os.makedirs("pos_neutre", exist_ok=True)
        for a in df["URL"]:
            try:
                download(a,chm+'/'+str(x)+'/'+'pos_neutre') 
                c = c+1
                print(c)
            except:
                pass
    except:
        pass
    try:
        df = pd.read_excel(chm+'/'+str(x)+'/'+'neg.xlsx')
        os.makedirs("neg", exist_ok=True)
        for a in df["URL"]:
            try:
                download(a,chm+'/'+str(x)+'/'+'neg') 
                c = c+1
                print(c)
            except:
                pass
    except:
        pass
    


cm = r'C:\Users\HP\Desktop\video\2020\10\pos_neutre'
df = pd.read_excel(cm+'\pos_neutre.xlsx')
for a in df["URL"]:
    try:
        download(a,cm)   
    except:
        print('lien_introuvable')

