import pandas as pd
import os
import shutil


Gesuchte_Person='Christian Ufrecht'
path_Exceldatei ="C:\\Users\\Christian\\Desktop\\TestArchive\\Übersicht Inhalt.xlsx"
path_BilderOrdner="C:\\Users\\Christian\\Desktop\\TestArchive"



Zielpfad = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')+'\\'+Gesuchte_Person
if not os.path.exists(Zielpfad ):
    os.makedirs(Zielpfad)

    


def get_filenames(path_Exceldatei,Gesuchte_Person):
    df = pd.read_excel(path_Exceldatei )
    Gesuchte_Bilder=[]
    for index, Personen in enumerate(df['Personen']):
        if type(Personen)==str:
            if Gesuchte_Person in Personen:
                   Gesuchte_Bilder.append(df['Bildname'][index])
    return Gesuchte_Bilder


Gesuchte_Bilder=get_filenames(path_Exceldatei,Gesuchte_Person)

# Get the list of all files in directory tree at given path
listOfFiles = list()
for (dirpath, dirnames, filenames) in os.walk(path_BilderOrdner):
    listOfFiles += [os.path.join(dirpath, file) for file in filenames]
    
for datei in Gesuchte_Bilder:
    for filename in listOfFiles:
        if '\\' + datei +  '.' in filename:
            shutil.copyfile(filename, Zielpfad+'\\'+datei+'.jpg')
        
        