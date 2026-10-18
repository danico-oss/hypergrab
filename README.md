# Hypergrab

**Cattura schermo dei test**

Applicazione desktop in Rust per l'acquisizione di evidenze durante la
qualificazione o la validazione di sistemi informatizzati. Importa un elenco di
test da un file CSV o XLSX, cattura la finestra dell'applicazione che resta in
primo piano, vi sovrimprime data e ora e salva il JPG con il nome del codice del
test.

## Funzionalità

- **Importazione elenco**: file CSV/TXT (separatore `;`, virgola o tabulazione,
  riconosciuto automaticamente) ed Excel (`.xlsx`, `.xlsm`, `.xls`, `.ods`,
  primo foglio). I file di testo sono letti come UTF-8 e, se non validi, come
  Windows-1252: i CSV esportati da Excel in italiano non perdono gli accenti.
- **Elenco dei test**: ogni riga mostra codice e descrizione; un clic la
  seleziona (la selezione è sempre e solo una cornice, mai un riempimento).
  All'apertura del file viene selezionato il primo test, così si può catturare
  subito.
- **Cattura**: il pulsante "Esegui cattura" o il tasto **F12** nascondono
  Hypergrab per qualche istante e fotografano la finestra che in quel momento è
  in primo piano. Se la finestra non è utilizzabile (per esempio è minimizzata)
  si cattura l'intero schermo virtuale, cioè tutti i monitor insieme.
- **Data e ora**: vengono disegnate in basso a destra sull'immagine, su un velo
  scuro che le tiene leggibili su qualunque sfondo. Il formato segue la lingua
  dell'interfaccia: `18/10/2026 14:03:22` in italiano e portoghese,
  `10/18/2026 14:03:22` in inglese. Il carattere è preso dai font di Windows
  (Segoe UI e ripieghi).
- **Anteprima e salvataggio**: l'anteprima mostra esattamente l'immagine che
  verrà scritta. "Salva JPG" scrive il file in JPEG (qualità 90) nella **stessa
  cartella del file di elenco**, con **nome uguale al codice del test**; se il
  file esiste già viene sovrascritto e l'anteprima lo segnala.
- **Interfaccia**: tema chiaro (predefinito) e scuro, stile compatto adatto
  anche agli schermi dei notebook, lingue italiano / inglese / portoghese
  brasiliano. Le preferenze (tema e lingua) sono ricordate in
  `hypergrab_settings.txt`, accanto all'eseguibile.

## Formato del file di elenco

Due colonne, in quest'ordine:

| Colonna | Contenuto | Note |
| --- | --- | --- |
| 1 | **Codice del test** | diventa il nome del file JPG |
| 2 | **Descrizione del test** | facoltativa |

La casella "La prima riga è un'intestazione" (attiva per impostazione
predefinita) scarta la prima riga: togli la spunta se il file contiene già solo
i dati. Le righe senza codice vengono ignorate.

Esempio CSV:

```csv
Codice;Descrizione
T-01;Verifica iniziale dell'ambiente
T-02;Controllo dei parametri di processo
T-03;Prova di accesso con utente non autorizzato
```

Un elenco di prova è in `esempi/elenco_test.csv`.

Nel file Excel basta lo stesso schema nel primo foglio; un codice memorizzato
come numero (per esempio `1003`) viene usato come `1003`, non come `1003.0`.

## Come si usa

1. "Apri file…" e scegli l'elenco (`esempi/elenco_test.csv` per una prova).
2. Porta in primo piano l'applicazione da documentare.
3. Seleziona il test nell'elenco e premi **F12** (o usa "Esegui cattura").
   Hypergrab si nasconde per circa mezzo secondo e poi torna con l'anteprima.
4. Controlla l'anteprima e premi "Salva JPG".

Il percorso di destinazione è indicato sotto i pulsanti dell'anteprima, così si
sa sempre dove finirà il file.

## Note tecniche

- La cattura usa le API di Windows: `GetForegroundWindow` per la finestra,
  copia dei pixel dallo schermo (`BitBlt` su un bitmap compatibile) e, se
  qualcosa non è disponibile, ripiego sull'intero schermo virtuale. Il canale
  alfa dei bitmap dello schermo non è significativo e viene ignorato.
- Il codice sanitizza il nome del file (`\ / : * ? " < > |`, caratteri di
  controllo, punti e spazi finali) e aggiunge un suffisso ai nomi riservati dal
  sistema (`CON`, `PRN`, `COM1`…).
- L'icona è disegnata a mano nel codice (`src/icon.rs`): il test
  `icon::tests::scrivi_file_ico` rigenera `assets/hypergrab.ico`, che `build.rs`
  incorpora nell'eseguibile con `rc.exe` insieme ai dati di versione. Il test
  `anteprime_png` produce anche un foglio di contatto con le varie misure su
  sfondo chiaro e scuro.
- Il profilo di rilascio è collegato alla CRT statica
  (`.cargo/config.toml`), quindi `hypergrab.exe` non richiede le Microsoft
  Visual C++ Redistributable sulla macchina di destinazione.

## Compilazione

```powershell
cargo test          # genera anche assets/hypergrab.ico
cargo run --release
```

Serve una toolchain Rust per il target `x86_64-pc-windows-msvc`; per
incorporare l'icona serve `rc.exe` dello Windows SDK (se manca, la build
prosegue senza icona, con un avviso).

## Licenza

GNU General Public License v3.0 —
<https://github.com/danico-oss/hypergrab>

Autore: Daniel Consonni. Un ringraziamento speciale a Geovanna Mendes de Souza,
le cui osservazioni sono state il punto di partenza per lo sviluppo di questa
applicazione.
