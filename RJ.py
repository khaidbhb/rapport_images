#KHALID BOUHABBA
import glob
import os
from PIL import Image as pilimg
import xlsxwriter
import cv2
from openpyxl import load_workbook
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU, cm_to_EMU
from openpyxl.drawing.image import Image
from datetime import date

print("Merci de prendre en considération la nature des lettres (Majuscule/miniscule)\n")
e_name=str(input("Entrez le nom du fichier excel [avec extension, par exemple: rapport.xlxs]: "))
f_name=str(input("Entrez le nom de la feuille où vous voulez insérer les images: "))

x=0
y=50
j=0
percentage=0.161
img_height=2184
img_width=4608

def increment_char(c):
    return chr(ord(c) + 1) if c != 'Z' else 'AA'

def increment_str(s):
    lpart = s.rstrip('Z')
    if not lpart:  # s contains only 'Z'
        new_s = 'A' * (len(s) + 1)
    else:
        num_replacements = len(s) - len(lpart)
        new_s = lpart[:-1] + increment_char(lpart[-1])
        new_s += 'A' * num_replacements
    return new_s

def spaced_incr(s,n):
    for i in range(n):
        s=increment_str(s)
    return s

def reduce_size(image_path):
    image=pilimg.open(image_path)
    image.save(image_path,quality=60, optimize= True)




theFile = load_workbook(e_name)
worksheet = theFile[f_name]


images=glob.glob('images/*')
s=1
for image_ in images:
    try:
        #supprimer # ci-dessous si vous voulez réduire la taille/qualité des images
        #reduce_size(image_)

        if images.index(image_)%2==0:
            x =0
        else:
            x = img_width*percentage+2
        img = Image(image_)
        p2e = pixels_to_EMU
        w, h = img_width*percentage, img_height*percentage

        position = XDRPoint2D(p2e(x), p2e(y))
        size = XDRPositiveSize2D(p2e(w), p2e(h))
        img.anchor = AbsoluteAnchor(pos=position, ext=size)
        worksheet.add_image(img)
        j+=1
        if j%2==0:
            y+=img_height*percentage+2
            x=0
    except:
        print(f"Le fichier '{image_}' n'est pas une image! merci de le supprimer et de relancer le script")
        s=0
if s==1:
    theFile.save(e_name)
    print(f'\n le fichier {e_name} est modifié avec succès!')
else:
    print("\n modification non enregistrée, merci de vérifier le dossier des images.")
