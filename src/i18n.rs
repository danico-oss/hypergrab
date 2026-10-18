//! Internazionalizzazione dell'interfaccia.
//!
//! Le stringhe italiane sono la lingua di riferimento: nei punti dell'interfaccia
//! si passa la stringa italiana a [`t`], che restituisce la traduzione nella lingua
//! attiva. Se una traduzione manca, viene mostrato l'italiano (fallback sicuro).
//!
//! Le lingue supportate sono italiano, portoghese brasiliano e inglese.

/// Lingua dell'interfaccia.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum Language {
    It,
    PtBr,
    En,
}

impl Default for Language {
    fn default() -> Self {
        Language::It
    }
}

impl Language {
    /// Tutte le lingue supportate, nell'ordine mostrato nella barra alta.
    pub const ALL: [Language; 3] = [Language::It, Language::PtBr, Language::En];

    /// Codice salvato nel file delle preferenze.
    pub fn code(self) -> &'static str {
        match self {
            Language::It => "it",
            Language::PtBr => "pt-BR",
            Language::En => "en",
        }
    }

    /// Codice breve mostrato nei pulsanti della barra alta.
    pub fn short(self) -> &'static str {
        match self {
            Language::It => "IT",
            Language::PtBr => "PT",
            Language::En => "EN",
        }
    }

    /// Interpreta un codice lingua (tollerante a maiuscole e varianti).
    pub fn from_code(code: &str) -> Option<Language> {
        match code.trim().to_lowercase().as_str() {
            "it" | "ita" | "italiano" | "italian" => Some(Language::It),
            "pt" | "pt-br" | "pt_br" | "ptbr" | "portugues" | "português" | "portuguese" => {
                Some(Language::PtBr)
            }
            "en" | "eng" | "english" | "inglese" | "inglês" => Some(Language::En),
            _ => None,
        }
    }
}

/// Traduce una stringa italiana nella lingua attiva.
pub fn t(lang: Language, it: &'static str) -> &'static str {
    match lang {
        Language::It => it,
        Language::En => lookup(EN, it).unwrap_or(it),
        Language::PtBr => lookup(PT, it).unwrap_or(it),
    }
}

/// Sceglie fra la forma singolare e plurale in base al numero.
pub fn plural(lang: Language, n: usize, one: &'static str, many: &'static str) -> &'static str {
    t(lang, if n == 1 { one } else { many })
}

fn lookup(table: &'static [(&'static str, &'static str)], key: &str) -> Option<&'static str> {
    table.iter().find(|(k, _)| *k == key).map(|(_, v)| *v)
}

static EN: &[(&str, &str)] = &[
    ("Cattura schermo dei test", "Test screenshot capture"),
    ("Importa elenco test", "Import test list"),
    ("Apri file…", "Open file…"),
    ("Apri elenco test", "Open test list"),
    ("La prima riga è un'intestazione", "The first row is a header"),
    ("Elenco test", "Test list"),
    ("Nessun file importato", "No file imported"),
    ("Apri un CSV o un XLSX per vedere i test.", "Open a CSV or an XLSX to list the tests."),
    ("Nessun test nel file.", "No tests in the file."),
    ("Seleziona un test dall'elenco prima di catturare.", "Select a test from the list before capturing."),
    ("Esegui cattura", "Run capture"),
    ("oppure premi F12", "or press F12"),
    ("Anteprima", "Preview"),
    ("Nessuna anteprima: esegui una cattura.", "No preview: run a capture."),
    ("Salva JPG", "Save JPG"),
    ("Pronto.", "Ready."),
    ("Errore di lettura:", "Read error:"),
    ("Elenco importato da:", "List imported from:"),
    ("test", "test"),
    ("tests", "tests"),
    ("Cattura in corso: la finestra torna a breve…", "Capture in progress: the window will be back shortly…"),
    ("Test selezionato:", "Selected test:"),
    ("schermo intero", "the whole screen"),
    ("Cattura completata:", "Capture completed:"),
    ("Cattura fallita.", "Capture failed."),
    ("Nessun test selezionato.", "No test selected."),
    ("Nessun file importato.", "No file imported."),
    ("Salvato:", "Saved:"),
    ("Errore nel salvataggio:", "Save error:"),
    ("Il file esistente verrà sovrascritto.", "The existing file will be overwritten."),
    ("Informazioni", "About"),
    ("Lingua", "Language"),
    ("Cambia tema", "Switch theme"),
    (
        "Hypergrab è uno strumento di cattura di schermo pensato per accompagnare il processo di acquisizione di evidenze durante la qualificazione o la validazione di sistemi informatizzati.",
        "Hypergrab is a screen capture tool designed to support the process of gathering evidence during the qualification or validation of computerized systems.",
    ),
    (
        "Permette di importare un elenco di test da un file CSV o XLSX (prima colonna: codice del test, seconda colonna: descrizione) e di catturare la finestra dell'applicazione che resta in primo piano nascondendo temporaneamente Hypergrab.",
        "It lets you import a list of tests from a CSV or XLSX file (first column: test code, second column: test description) and capture the window of the application that stays in the foreground while Hypergrab hides itself.",
    ),
    (
        "All'immagine viene aggiunta in basso a destra la data e l'ora di acquisizione; il file JPG viene salvato nella cartella del file di elenco, con nome uguale al codice del test.",
        "The acquisition date and time are added to the bottom right of the image; the JPG file is saved in the folder of the list file, named after the test code.",
    ),
    (
        "Sviluppato in Rust con Egui come sfida personale. Distribuito con licenza GNU GPL v3. Autore: Daniel Consonni. Un ringraziamento speciale a Geovanna Mendes de Souza, le cui osservazioni sono state il punto di partenza per lo sviluppo di questa applicazione.",
        "Developed in Rust with Egui as a personal challenge. Distributed under the GNU GPL v3 license. Author: Daniel Consonni. A special thank you to Geovanna Mendes de Souza, whose observations were the starting point for the development of this application.",
    ),
];

static PT: &[(&str, &str)] = &[
    ("Cattura schermo dei test", "Captura de tela dos testes"),
    ("Importa elenco test", "Importar lista de testes"),
    ("Apri file…", "Abrir arquivo…"),
    ("Apri elenco test", "Abrir lista de testes"),
    ("La prima riga è un'intestazione", "A primeira linha é um cabeçalho"),
    ("Elenco test", "Lista de testes"),
    ("Nessun file importato", "Nenhum arquivo importado"),
    ("Apri un CSV o un XLSX per vedere i test.", "Abra um CSV ou um XLSX para ver os testes."),
    ("Nessun test nel file.", "Nenhum teste no arquivo."),
    ("Seleziona un test dall'elenco prima di catturare.", "Selecione um test na lista antes de capturar."),
    ("Esegui cattura", "Executar captura"),
    ("oppure premi F12", "ou pressione F12"),
    ("Anteprima", "Pré-visualização"),
    ("Nessuna anteprima: esegui una cattura.", "Nenhuma pré-visualização: execute uma captura."),
    ("Salva JPG", "Salvar JPG"),
    ("Pronto.", "Pronto."),
    ("Errore di lettura:", "Errore de leitura:"),
    ("Elenco importato da:", "Lista importada de:"),
    ("test", "teste"),
    ("tests", "testes"),
    ("Cattura in corso: la finestra torna a breve…", "Captura em andamento: a janela volta em instantes…"),
    ("Test selezionato:", "Teste selecionado:"),
    ("schermo intero", "a tela inteira"),
    ("Cattura completata:", "Captura concluída:"),
    ("Cattura fallita.", "Captura falhou."),
    ("Nessun test selezionato.", "Nenhum teste selecionado."),
    ("Nessun file importato.", "Nenhum arquivo importado."),
    ("Salvato:", "Salvo:"),
    ("Errore nel salvataggio:", "Errore ao salvar:"),
    ("Il file esistente verrà sovrascritto.", "O arquivo existente será sobrescrito."),
    ("Informazioni", "Informações"),
    ("Lingua", "Idioma"),
    ("Cambia tema", "Alternar tema"),
    (
        "Hypergrab è uno strumento di cattura di schermo pensato per accompagnare il processo di acquisizione di evidenze durante la qualificazione o la validazione di sistemi informatizzati.",
        "Hypergrab é uma ferramenta de captura de tela pensada para acompanhar o processo de aquisição de evidências durante a qualificação ou validação de sistemas informatizados.",
    ),
    (
        "Permette di importare un elenco di test da un file CSV o XLSX (prima colonna: codice del test, seconda colonna: descrizione) e di catturare la finestra dell'applicazione che resta in primo piano nascondendo temporaneamente Hypergrab.",
        "Permite importar uma lista de testes de um arquivo CSV ou XLSX (primeira coluna: código do teste, segunda coluna: descrição) e capturar a janela da aplicação que permanece em primeiro plano ocultando temporariamente o Hypergrab.",
    ),
    (
        "All'immagine viene aggiunta in basso a destra la data e l'ora di acquisizione; il file JPG viene salvato nella cartella del file di elenco, con nome uguale al codice del test.",
        "À imagem é adicionada, no canto inferior direito, a data e a hora da aquisição; o arquivo JPG é salvo na pasta do arquivo de lista, com nome igual ao código do teste.",
    ),
    (
        "Sviluppato in Rust con Egui come sfida personale. Distribuito con licenza GNU GPL v3. Autore: Daniel Consonni. Un ringraziamento speciale a Geovanna Mendes de Souza, le cui osservazioni sono state il punto di partenza per lo sviluppo di questa applicazione.",
        "Desenvolvido em Rust com Egui como desafio pessoal. Distribuído sob a licença GNU GPL v3. Autor: Daniel Consonni. Um agradecimento especial a Geovanna Mendes de Souza, cujas observações foram o ponto de partida para o desenvolvimento desta aplicação.",
    ),
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn italiano_e_identita() {
        assert_eq!(t(Language::It, "Salva JPG"), "Salva JPG");
    }

    #[test]
    fn traduce_in_inglese_e_portoghese() {
        assert_eq!(t(Language::En, "Salva JPG"), "Save JPG");
        assert_eq!(t(Language::PtBr, "Salva JPG"), "Salvar JPG");
    }

    #[test]
    fn chiave_sconosciuta_resta_italiana() {
        assert_eq!(t(Language::En, "chiave inesistente"), "chiave inesistente");
    }

    #[test]
    fn codici_lingua() {
        assert_eq!(Language::from_code("pt-BR"), Some(Language::PtBr));
        assert_eq!(Language::from_code("EN"), Some(Language::En));
        assert_eq!(Language::from_code("xx"), None);
        assert_eq!(Language::default(), Language::It);
        assert_eq!(Language::It.short(), "IT");
        assert_eq!(Language::PtBr.short(), "PT");
        assert_eq!(Language::En.short(), "EN");
    }

    #[test]
    fn plurale() {
        assert_eq!(plural(Language::En, 1, "test", "tests"), "test");
        assert_eq!(plural(Language::En, 5, "test", "tests"), "tests");
        assert_eq!(plural(Language::PtBr, 1, "test", "tests"), "teste");
        assert_eq!(plural(Language::PtBr, 5, "test", "tests"), "testes");
        assert_eq!(plural(Language::It, 5, "test", "tests"), "tests");
    }

    #[test]
    fn nessuna_chiave_duplicata() {
        for table in [EN, PT] {
            for (i, (k, _)) in table.iter().enumerate() {
                for (k2, _) in &table[i + 1..] {
                    assert_ne!(k, k2, "chiave duplicata nella tabella: {k}");
                }
            }
        }
    }

    #[test]
    fn stesse_chiavi_in_tutte_le_lingue() {
        for (k, _) in EN {
            assert!(
                PT.iter().any(|(k2, _)| k2 == k),
                "chiave presente in EN ma non in PT: {k}"
            );
        }
        for (k, _) in PT {
            assert!(
                EN.iter().any(|(k2, _)| k2 == k),
                "chiave presente in PT ma non in EN: {k}"
            );
        }
    }

    #[test]
    fn en_e_diverso_dall_italiano_sul_contenuto_lungo() {
        // Le pagine "Informazioni" hanno chiavi in comune: la traduzione EN
        // deve essere effettivamente diversa dall'italiano.
        let it = "Hypergrab è uno strumento di cattura di schermo pensato per accompagnare il processo di acquisizione di evidenze durante la qualificazione o la validazione di sistemi informatizzati.";
        assert_ne!(t(Language::En, it), it);
        assert_ne!(t(Language::PtBr, it), it);
    }
}
