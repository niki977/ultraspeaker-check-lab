![The Ultraspeaker Check Lab](assets/img/check-lab-logo.png)

# The Ultraspeaker Check Lab

Controlla una presentazione PowerPoint (.pptx) o Keynote (.key) prima di consegnarla alla regia di un congresso: font, video e audio, immagini, leggibilità (contrasto con lo sfondo e con le foto, grandezza del testo per l’ultima fila, grafici per chi distingue male rosso e verde), formato e struttura, tempi, metadati EXIF delle foto. La scheda **Stile** aiuta a scegliere sfondo, colore del testo e font prima di creare le slide e crea un modello .pptx con queste scelte. Crea il report con punteggio «Pronto per la regia», l’esportazione PDF, il nome del file per la regia, la checklist del giorno e i kit di prova video e immagini da aprire in sala.

- **Versione online** (PowerPoint e Keynote): https://niki977.github.io/ultraspeaker-check-lab/
- **Add-in PowerPoint** (Windows, Mac, web): `ultraspeaker-check-lab-manifest.xml`
- Interfaccia e pagine in italiano, inglese, spagnolo, francese e tedesco.

L’analisi avviene interamente nel browser o nel pannello di PowerPoint: la presentazione non viene mai inviata a server esterni.

## Struttura

```
index.html            versione online (.pptx e .key)
addin/taskpane.html   pannello dell’add-in PowerPoint
addin/commands.html   file di funzioni richiesto dal manifest
ultraspeaker-check-lab-manifest.xml   manifest dell’add-in (sideload e Marketplace)
guide.html, support.html, privacy.html, terms.html   pagine in 5 lingue
assets/css/app.css    stile condiviso
assets/js/            motore di analisi (pptx, key, testi e contrasto, foto), scheda Stile, interfaccia, traduzioni
assets/kit/           modello e file di esempio dei kit di prova
assets/icons/         icone 16–512 px (pittogramma della E-checklist)
assets/img/           logo Check Lab (bianco e nero, SVG e PNG), Arena
assets/fonts/         Champagne & Limousines (Nymphont): woff2 per l’interfaccia, ttf per il PDF
assets/vendor/        JSZip 3.10.1 (MIT), jsPDF 2.5.2 (MIT)
test/                 presentazioni di prova con problemi voluti (generali e di leggibilità)
```

## Leggibilità: come si calcola

- Contrasto WCAG 2 tra testo e sfondo, con la luce della sala sommata all’immagine proiettata (buio 1%, penombra 5%, luci accese 12% del bianco). Minimo 4,5 : 1, 3 : 1 per titoli e testo grande.
- Testo su foto: la parte della foto sotto il testo viene letta nel browser; conta il 5% dell’area con meno contrasto, anche con una banda al 60% dietro il testo.
- Grandezza: le minuscole alte almeno 1/200 della distanza dell’ultima fila (AVIXA DISCAS), con l’altezza x del font e le bande nere se la slide non riempie lo schermo. In PowerPoint i punti sono su 540 pt di altezza, in Keynote sui pixel della slide.

## Immagine coordinata

Allineata al Brand Book di The Ultraspeaker: logo master con il descrittore CHECK LAB, in cui la E è una checklist a tre righe uguali; palette Stage Black; carattere Champagne & Limousines nell’interfaccia, nelle pagine e nel PDF; angoli 10 px, pulsante primario bianco, nessuna ombra, focus in blu #4F8DFF. Eccezione concordata: nel report Critico è rosso, Attenzione giallo, OK verde. Kit di prova: sfondo nero, Arena in basso a destra, testo in Arial per la massima compatibilità con i computer di regia.

© Nicholas D. Charles — The Ultraspeaker
