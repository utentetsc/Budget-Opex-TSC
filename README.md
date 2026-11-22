# The Space — Budget manutenzione straordinaria

Web app per impostare i pesi di priorità, distribuire il budget tra i cinema e ripartire la quota di ciascuna sede secondo i consuntivi dell'anno scelto.

## Avvio in locale

1. Estrai lo ZIP e apri la cartella `the-space-budget-github`.
2. Installa Node.js 20.19 o superiore, oppure Node.js 22.12 o superiore.
3. Esegui `npm install`.
4. Esegui `npm run dev` e apri l'indirizzo mostrato nel terminale.

Per una build statica esegui `npm run build`; i file generati si trovano in `dist/`.

## Importazione e dati
Il progetto include i nomi dei 35 cinema presenti nel file e TSC Odeon Milano; il file originale e gli importi aziendali non sono inclusi.

Carica il file Excel o CSV dalla schermata iniziale. L'importazione avviene nel browser: il file non viene inviato a un server e non è incluso in questo ZIP. I dati del piano vengono salvati nel local storage del browser. Il file CSV esportato contiene le quote calcolate.

Il file di riferimento originale non è incluso per non inserire nel repository i consuntivi aziendali.
 
