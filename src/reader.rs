//! Lettura dell'elenco dei test da un file CSV o XLSX.
//!
//! Formato atteso: **prima colonna** il codice del test, **seconda colonna**
//! la descrizione. Nei file CSV il separatore viene riconosciuto dalla prima
//! riga (`;`, virgola o tabulazione); i file Excel usano il primo foglio.

use calamine::{open_workbook_auto, Data, Reader};
use std::path::Path;

/// Una voce dell'elenco dei test.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TestItem {
    /// Codice del test: diventa il nome del file JPG.
    pub code: String,
    /// Descrizione del test (facoltativa).
    pub description: String,
}

/// Legge l'elenco dei test dal file indicato.
///
/// `salta_intestazione` va messo a `true` quando la prima riga contiene i titoli
/// delle colonne: in quel caso viene scartata.
pub fn leggi_elenco(path: &Path, salta_intestazione: bool) -> Result<Vec<TestItem>, String> {
    let estensione = path
        .extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_default();
    match estensione.as_str() {
        "xlsx" | "xlsm" | "xls" | "xlsb" | "ods" => leggi_foglio(path, salta_intestazione),
        _ => leggi_csv(path, salta_intestazione),
    }
}

/// Nome del foglio letto per primo, usato nei test e nei messaggi di stato.
#[cfg(test)]
fn nome_foglio(path: &Path) -> Option<String> {
    let wb = open_workbook_auto(path).ok()?;
    wb.sheet_names().first().cloned()
}

// ---------------------------------------------------------------------------
// CSV
// ---------------------------------------------------------------------------

/// Legge un file di testo separato da `;`, virgola o tabulazione.
fn leggi_csv(path: &Path, salta_intestazione: bool) -> Result<Vec<TestItem>, String> {
    let bytes = std::fs::read(path).map_err(|e| e.to_string())?;
    let testo = decodifica(&bytes);
    let delimitatore = separatore(&testo);

    let mut lettore = csv::ReaderBuilder::new()
        .delimiter(delimitatore)
        .has_headers(false)
        .flexible(true)
        .from_reader(testo.as_bytes());

    let mut voci = Vec::new();
    for (indice, record) in lettore.records().enumerate() {
        let record = record.map_err(|e| e.to_string())?;
        if indice == 0 && salta_intestazione {
            continue;
        }
        let code = record.get(0).unwrap_or("").trim();
        if code.is_empty() {
            continue;
        }
        let description = record.get(1).unwrap_or("").trim();
        voci.push(TestItem {
            code: code.to_string(),
            description: description.to_string(),
        });
    }
    Ok(voci)
}

/// Riconosce il separatore dalla prima riga: vince quello piu' frequente.
/// In caso di parita' (o di nessun separatore) si usa la virgola.
pub fn separatore(testo: &str) -> u8 {
    let prima = testo.lines().next().unwrap_or("");
    let conta = |c: char| prima.matches(c).count();
    let (punto_e_virgola, virgola, tab) = (conta(';'), conta(','), conta('\t'));
    if tab > virgola && tab > punto_e_virgola {
        b'\t'
    } else if punto_e_virgola > virgola {
        b';'
    } else {
        b','
    }
}

/// Interpreta i byte del file come testo: UTF-8 (con o senza BOM) e, se non e'
/// UTF-8 valido, come Windows-1252 (il caso tipico dei CSV esportati da Excel
/// in italiano).
pub fn decodifica(bytes: &[u8]) -> String {
    let bytes = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes);
    match std::str::from_utf8(bytes) {
        Ok(s) => s.to_string(),
        Err(_) => bytes.iter().map(|&b| da_cp1252(b)).collect(),
    }
}

/// Converte un byte secondo la tabella Windows-1252.
fn da_cp1252(b: u8) -> char {
    // Solo l'intervallo 0x80..0x9F differisce da Latin-1: sono i caratteri che
    // Excel usa al posto dei caratteri di controllo.
    const ALTO: [char; 32] = [
        '\u{20AC}', '\u{81}', '\u{201A}', '\u{0192}', '\u{201E}', '\u{2026}', '\u{2020}',
        '\u{2021}', '\u{02C6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{8D}',
        '\u{017D}', '\u{8F}', '\u{90}', '\u{2018}', '\u{2019}', '\u{201C}', '\u{201D}',
        '\u{2022}', '\u{2013}', '\u{2014}', '\u{02DC}', '\u{2122}', '\u{0161}', '\u{203A}',
        '\u{0153}', '\u{9D}', '\u{017E}', '\u{0178}',
    ];
    if (0x80..0xA0).contains(&b) {
        ALTO[(b - 0x80) as usize]
    } else {
        b as char
    }
}

// ---------------------------------------------------------------------------
// XLSX
// ---------------------------------------------------------------------------

/// Legge il primo foglio di un file Excel.
fn leggi_foglio(path: &Path, salta_intestazione: bool) -> Result<Vec<TestItem>, String> {
    let mut cartella = open_workbook_auto(path).map_err(|e| e.to_string())?;
    let nome = cartella
        .sheet_names()
        .first()
        .cloned()
        .ok_or_else(|| "il file non contiene fogli".to_string())?;
    let intervallo = cartella
        .worksheet_range(&nome)
        .map_err(|e| e.to_string())?;

    let mut voci = Vec::new();
    for (indice, riga) in intervallo.rows().enumerate() {
        if indice == 0 && salta_intestazione {
            continue;
        }
        let code = testo_cella(riga.first());
        if code.is_empty() {
            continue;
        }
        voci.push(TestItem {
            code,
            description: testo_cella(riga.get(1)),
        });
    }
    Ok(voci)
}

/// Testo di una cella: i numeri interi perdono il separatore decimale, cosi'
/// un codice memorizzato come numero resta "1001" e non "1001.0".
fn testo_cella(cella: Option<&Data>) -> String {
    match cella {
        None | Some(Data::Empty) => String::new(),
        Some(Data::String(s)) => s.trim().to_string(),
        Some(Data::Float(f)) if f.fract() == 0.0 => format!("{}", *f as i64),
        Some(altra) => altra.to_string().trim().to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    /// Percorso temporaneo unico per i file di prova.
    fn temporaneo(nome: &str) -> PathBuf {
        let dir = std::env::temp_dir();
        let unico = format!(
            "hypergrab_{}_{}_{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos(),
            nome
        );
        dir.join(unico)
    }

    #[test]
    fn riconosce_il_separatore() {
        assert_eq!(separatore("Codice;Descrizione\nT-1;Prima"), b';');
        assert_eq!(separatore("Codice,Descrizione\nT-1,Prima"), b',');
        assert_eq!(separatore("Codice\tDescrizione\nT-1\tPrima"), b'\t');
        // Nessun separatore: virgola predefinita.
        assert_eq!(separatore("Codice\nT-1"), b',');
        // Parita' fra punto e virgola e virgola: virgola.
        assert_eq!(separatore("a;b,c"), b',');
        // Conta solo la prima riga.
        assert_eq!(separatore("a;b\nc,d,e"), b';');
    }

    #[test]
    fn decodifica_utf8_e_windows_1252() {
        assert_eq!(decodifica("Ciao àèì".as_bytes()), "Ciao àèì");
        // BOM in testa rimosso.
        let mut con_bom = vec![0xEF, 0xBB, 0xBF];
        con_bom.extend_from_slice("Codice".as_bytes());
        assert_eq!(decodifica(&con_bom), "Codice");
        // Windows-1252: 0xE7 = 'ç', 0xE3 = 'ã'.
        assert_eq!(decodifica(&[b'C', 0xE7, b'a']), "Cça");
        assert_eq!(decodifica(&[0xE3]), "ã");
        // 0x80 in Windows-1252 e' l'euro.
        assert_eq!(decodifica(&[0x80]), "€");
    }

    #[test]
    fn legge_csv_con_intestazione() {
        let path = temporaneo("elenco.csv");
        std::fs::write(
            &path,
            "Codice;Descrizione\nT-01;Verifica iniziale\nT-02;Verifica finale\n",
        )
        .unwrap();

        let voci = leggi_elenco(&path, true).unwrap();
        assert_eq!(voci.len(), 2);
        assert_eq!(voci[0].code, "T-01");
        assert_eq!(voci[0].description, "Verifica iniziale");
        assert_eq!(voci[1].code, "T-02");

        // Senza intestazione la prima riga e' un dato come gli altri.
        let voci = leggi_elenco(&path, false).unwrap();
        assert_eq!(voci.len(), 3);
        assert_eq!(voci[0].code, "Codice");
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn legge_csv_con_campi_fra_virgolette() {
        let path = temporaneo("virgolette.csv");
        std::fs::write(
            &path,
            "code,description\nT-1,\"Descrizione; con separatore\"\nT-2,\"Virgolette \"\"interne\"\"\"\n",
        )
        .unwrap();
        let voci = leggi_elenco(&path, true).unwrap();
        assert_eq!(voci.len(), 2);
        assert_eq!(voci[0].description, "Descrizione; con separatore");
        assert_eq!(voci[1].description, "Virgolette \"interne\"");
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn salta_le_righe_senza_codice() {
        let path = temporaneo("vuote.csv");
        std::fs::write(&path, "C;D\nT-1;ok\n;descrizione senza codice\n\nT-2;ok\n").unwrap();
        let voci = leggi_elenco(&path, true).unwrap();
        assert_eq!(voci.len(), 2);
        assert_eq!(voci[1].code, "T-2");
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn legge_un_foglio_xlsx() {
        let path = temporaneo("elenco.xlsx");
        std::fs::write(&path, xlsx_di_prova()).unwrap();

        let voci = leggi_elenco(&path, true).unwrap();
        assert_eq!(voci.len(), 3);
        assert_eq!(voci[0].code, "T-01");
        assert_eq!(voci[0].description, "Prima prova");
        assert_eq!(voci[1].code, "T-02");
        assert_eq!(voci[1].description, "Seconda, prova");
        // Codice memorizzato come numero: niente ".0".
        assert_eq!(voci[2].code, "1003");
        assert_eq!(voci[2].description, "Codice numerico");

        assert_eq!(nome_foglio(&path).as_deref(), Some("Elenco"));
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn file_inesistente_da_errore() {
        let path = temporaneo("non_esiste.csv");
        assert!(leggi_elenco(&path, true).is_err());
    }

    /// Costruisce un XLSX minimo in memoria: intestazione + tre righe, con
    /// una cella numerica, cosi' lettura e conversione sono verificate senza
    /// dipendere da file esterni.
    fn xlsx_di_prova() -> Vec<u8> {
        let righe = [
            ("Codice", "Descrizione", None),
            ("T-01", "Prima prova", None),
            ("T-02", "Seconda, prova", None),
            ("1003", "Codice numerico", Some(1003.0)),
        ];
        let mut righe_xml = String::new();
        for (i, (a, b, numero)) in righe.iter().enumerate() {
            let r = i + 1;
            righe_xml.push_str(&format!("<row r=\"{r}\">"));
            match numero {
                Some(n) => righe_xml.push_str(&format!("<c r=\"A{r}\"><v>{n}</v></c>")),
                None => righe_xml.push_str(&format!(
                    "<c r=\"A{r}\" t=\"inlineStr\"><is><t>{}</t></is></c>",
                    escape(a)
                )),
            }
            righe_xml.push_str(&format!(
                "<c r=\"B{r}\" t=\"inlineStr\"><is><t>{}</t></is></c>",
                escape(b)
            ));
            righe_xml.push_str("</row>");
        }

        let parti: [(&str, String); 5] = [
            (
                "[Content_Types].xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#
                    .to_string(),
            ),
            (
                "_rels/.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#
                    .to_string(),
            ),
            (
                "xl/workbook.xml",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Elenco" sheetId="1" r:id="rId1"/></sheets>
</workbook>"#
                    .to_string(),
            ),
            (
                "xl/_rels/workbook.xml.rels",
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#
                    .to_string(),
            ),
            (
                "xl/worksheets/sheet1.xml",
                format!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{righe_xml}</sheetData></worksheet>"#
                ),
            ),
        ];

        let mut buffer = std::io::Cursor::new(Vec::new());
        {
            let mut zip = zip::ZipWriter::new(&mut buffer);
            let opzioni = zip::write::SimpleFileOptions::default();
            for (nome, contenuto) in parti {
                use std::io::Write;
                zip.start_file(nome, opzioni).unwrap();
                zip.write_all(contenuto.as_bytes()).unwrap();
            }
            zip.finish().unwrap();
        }
        buffer.into_inner()
    }

    fn escape(testo: &str) -> String {
        testo
            .replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
    }
}
