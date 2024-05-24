# signacraft v1.0 - by chriffpy.
print("signacraft v1.0 - by Christopher Gertig - chriffpy")
print("Wird geladen...")
import os
import pandas as pd
import glob
import shutil
import re


# Pfad zum Arbeitsverzeichnis (Ordner in dem die Signaturdateien liegen)
# Beispiel: base_dir = 'E:\\Daten\\Signaturen'
base_dir = ''

print("Einlesen der Benutzer")
# Einlesen der Excel
excel_file = os.path.join(base_dir, 'Benutzer.xlsx')
df = pd.read_excel(excel_file)

print("Einlesen des Templates")
# Einlesen der HTML
template_file = os.path.join(base_dir, 'template.htm')
with open(template_file, 'r', encoding='utf-8') as file:
    html_template = file.read()

# Pfad zum eigenständigen Ordner "signature-Dateien", der kopiert werden soll
source_folder = os.path.join(base_dir, 'signature-Dateien')

# Pfad zum Ordner, in dem die Signaturen gespeichert werden sollen
# Beispiel: signatures_folder = 'E:\\Signatur'
signatures_folder = ''

print("Erstelle Signaturen für:")
# Erstellen der Signatur
for index, row in df.iterrows():
    # Erstellen der Ordner für den Benutzernamen, wenn er noch nicht existiert
    username_folder = os.path.join(signatures_folder, str(row['Benutzername']))
    if not os.path.exists(username_folder):
        os.makedirs(username_folder)
        

    # Kopieren des Ordners "signature-Dateien" in den Benutzerordner
    content_folder = os.path.join(username_folder, 'signature-Dateien')
    if os.path.exists(content_folder):
        shutil.rmtree(content_folder)
    shutil.copytree(source_folder, content_folder)
    print(username_folder)
  
    # Ersetzen der Platzhalter
    signature = html_template.replace('PERSONNAME', str(row['Name']))
    signature = signature.replace('POS', str(row['Position']))
    signature = signature.replace('EMAILADDR', str(row['E-Mail']))
    signature = signature.replace('DURCHWAHL', str(row['Durchwahl']))

    if 'Arbeitszeiten' in row and pd.notna(row['Arbeitszeiten']):
            lines = row['Arbeitszeiten'].split('\n')
            arbeitszeiten_html = ''.join([f'<p class="MsoNormal">{line}<o:p></o:p></P>' for line in lines])
            signature = signature.replace('Arbeitszeit', arbeitszeiten_html + "<p></p><p></p>")
            print("ARBEITSZEIT")
            
    else:
            signature = signature.replace('Arbeitszeit', '')

    # Abteilungsordner festlegen
    department_folder = os.path.join(base_dir, str(row['Abteilung']))

    # Ordner mit den Bildern
    image_folder = os.path.join(department_folder, 'Bilder')

    # Einlesen Excel-Datei mit den Links
    link_file = os.path.join(department_folder, 'Links.xlsx')
    links_df = pd.read_excel(link_file)

    # Umbenennen und Einbetten der Bilder
    image_files = sorted([f for f in glob.glob(os.path.join(image_folder, '*')) if not f.endswith('Thumbs.db')])

    link_index = 0  # Index für den Zugriff auf links_df
    for i, image_file in enumerate(image_files, start=6):  # Starte bei 6 für "image006"
        new_image_name = f'image{i:03d}.jpg'  # Neuer Dateiname im Format "imageXXX.jpg"

        # Kopieren und Umbenennen der Bilder in den Benutzerordner
        new_image_path = os.path.join(content_folder, new_image_name)
        shutil.copy(image_file, new_image_path)

        # Link in Variable speichern
        if link_index < len(links_df):
            link_row = links_df.iloc[link_index]
            link = link_row['Link'] if not pd.isna(link_row['Link']) else ''
        else:
            link = ''

        # Ersetzen des Bild-Links im HTML-Code
        image_tag = f'<p><span style=\'font-size:9.0pt;font-family:"Arial Narrow",sans-serif;color:black\'>&nbsp;<!--[if gte vml 1]><v:shape id="_x0000_i1032" type="#_x0000_t75" style=\'width:470.25pt;height:134.25pt\'><v:imagedata src="signature-Dateien/{new_image_name}" o:title="Banner"/></v:shape><![endif]--><![if !vml]><a href="{link}"><img border=0 width=627 height=179 src="signature-Dateien/{new_image_name}" v:shapes="_x0000_i1032"><![endif]></a></span><span style=\'font-size:9.0pt;font-family:"Arial Narrow",sans-serif;color:black\'><o:p></o:p></span></p>'

        # Ersetzen des Platzhalters im HTML-Code
        signature = signature.replace(f'BANNERIMAGE{i}', image_tag)

        link_index += 1  # den link_index für den nächsten Durchlauf erhöhen
        
    # Entfernen aller übrigen "BANNERIMAGE"-Platzhalter
    signature = re.sub('BANNERIMAGE\d+', '', signature)


    
    # Speichern der Signatur in einer HTML-Datei im entsprechenden Ordner
    with open(os.path.join(username_folder, 'signature.htm'), 'w', encoding='utf-8') as signature_file:
        signature_file.write(signature)

print("Signaturen wurden erfolgreich erstellt.")
