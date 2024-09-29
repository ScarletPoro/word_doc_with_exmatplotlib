from matplotlib import pyplot
from docx import Document
from docx.shared import Inches
import os

# Funzione per sostituire un placeholder con un'immagine
def replace_placeholder_with_image(doc, placeholder, image_path):
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            # Sostituisci il placeholder con l'immagine
            paragraph.clear()  # Cancella il testo del paragrafo
            run = paragraph.add_run()
            run.add_picture(image_path, width=Inches(5))  # Aggiunge il grafico come immagine

# Salva il grafico come immagine
def crea_grafico():
    # Imposta lo stile del grafico
    pyplot.style.use('dark_background')

    # Dati del grafico
    x = [25, 30, 32, 50, 75]
    y1 = [100, 82, 50, 33, 92]
    y2 = [90, 70, 150, 70, 22]

    # Creazione del grafico
    pyplot.plot(x, y1, label='italiani', color="red", marker='o', linewidth=2)
    pyplot.plot(x, y2, label='francesi', color="blue", linestyle='dotted', marker='.')

    # Aggiunta di titoli e legenda
    pyplot.title('nome del grafico')
    pyplot.xlabel('x - età')
    pyplot.ylabel('y - peso')
    pyplot.legend()
    pyplot.grid()

    # Salva il grafico come immagine PNG
    image_path = "grafico.png"
    pyplot.savefig(image_path)
    
    return image_path

# Percorso del file Word esistente
word_file_path = "documento_esistente.docx"

# Apri il documento esistente
doc = Document(word_file_path)

# Crea il grafico e ottieni il percorso dell'immagine
image_path = crea_grafico()

# Sostituisci il placeholder <<DIAGRAMMA1>> con il grafico
replace_placeholder_with_image(doc, '<<DIAGRAMMA1>>', image_path)
replace_placeholder_with_image(doc, '<<DIAGRAMMA2>>', image_path)


# Salva il documento aggiornato
doc.save("documento_aggiornato.docx")

# Pulisci l'immagine temporanea
os.remove(image_path)

print("Placeholder sostituito e documento salvato come 'documento_aggiornato.docx'")
