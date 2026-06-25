# Fleet Documents & Fines Manager - MVP GitHub

Add-in MyGeotab per test personale: archivio documenti/multe associato agli asset Geotab, OCR lato browser e salvataggio su GitHub.

## Cosa fa

- Carica PDF/immagine.
- Esegue OCR nel browser con Tesseract.js.
- Prova a estrarre targa, data, importo, scadenza e numero verbale.
- Confronta la targa con gli asset Geotab.
- Chiede conferma obbligatoria prima del salvataggio.
- Salva il file in `/documents/<tipo>/`.
- Aggiunge un nuovo record a `/data/documents.json` senza cancellare quelli già presenti.
- Mostra dashboard e archivio consultabile dall'add-in.

## Struttura repository

```text
/index.html
/app.js
/style.css
/addin_config_example.json
/data/documents.json
/documents/multe/
/documents/libretti/
/documents/assicurazioni/
/documents/altro/
```

## Importante su GitHub

Per scrivere nel repository serve un GitHub fine-grained token con permesso `Contents: Read and write` sul solo repository di test.

Non inserire mai il token nel codice. Nell'MVP il token viene incollato nella UI e usato solo in memoria.

## Perché un solo documents.json

Tutti i documenti sono nello stesso indice. Ogni record contiene `type`:

- `fine`
- `registration`
- `insurance`
- `leasing`
- `inspection`
- `fuel_card`
- `other`

Così il sistema può crescere senza cambiare architettura.

## Salvataggio append-safe

Quando salvi un documento:

1. viene caricato il PDF/immagine con nome univoco;
2. viene riletto l'ultimo `data/documents.json` da GitHub;
3. viene aggiunto il nuovo record all'array;
4. viene riscritto `documents.json` con tutti i record, vecchi + nuovo.

Quindi un nuovo caricamento non sovrascrive i precedenti.

## Limiti MVP

- GitHub non è uno storage documentale di produzione.
- OCR browser non è perfetto, soprattutto sui PDF scannerizzati male.
- Nessuna gestione utenti/permessi documentale avanzata.
- Possibili conflitti se due utenti salvano nello stesso momento.

Per produzione valutare SharePoint, OneDrive, Google Drive, S3, Supabase Storage o backend dedicato.
