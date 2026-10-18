//! Hypergrab: screen capture of tests.
//!
//! The application imports a list of tests from a CSV or XLSX file (the first
//! column holds the code, the second the description), lets you select a test
//! and, with the "Esegui cattura" button or the F12 key, hides itself to
//! photograph the window that stays in the foreground. Date and time are
//! merged into the image at the bottom right, then the JPG is saved in the
//! folder of the list file under the name of the test code.
//!
//! Starting the program with a file path as an argument opens the list right
//! away; the same happens when a file is dragged onto the window.

// No console: this is an application with a graphical interface.
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod capture;
mod finestra;
mod i18n;
mod icon;
mod reader;
mod settings;
mod theme;

use eframe::egui;
use i18n::{t, Language};
use reader::TestItem;
use std::path::PathBuf;
use std::sync::mpsc::{Receiver, TryRecvError};
use std::time::{Duration, Instant};
use theme::Icon;

/// Delay between the moment the window hides itself and the capture: it lets
/// the underlying application come back to the foreground.
const ATTESA_CATTURA: Duration = Duration::from_millis(650);

/// Additional delay, after the previous one, so that the foreground really
/// passes to another application: without it, if Windows did not give the
/// foreground to the underlying one at once, Hypergrab would be photographed.
const ATTESA_PRIMO_PIANO: Duration = Duration::from_millis(1500);

/// How long, after the capture, we keep asking for the window to become
/// visible again: the first attempt may be lost.
const RIPRISTINO: Duration = Duration::from_millis(400);

enum Stato {
    /// Window visible, no capture in progress.
    Pronto,
    /// Window hidden and capture in progress on a separate thread.
    InCorso,
}

/// Outcome of the capture performed on a separate thread.
enum EsitoCattura {
    Cattura {
        immagine: image::RgbImage,
        finestra: String,
        data_ora: String,
    },
    Errore(String),
}

/// Captured image, ready for the preview and for saving.
struct Anteprima {
    immagine: image::RgbImage,
    texture: egui::TextureHandle,
    finestra: String,
    data_ora: String,
}

struct Hypergrab {
    lang: Language,
    dark_mode: bool,
    tema_applicato: bool,

    // --- Test list ---
    percorso: Option<PathBuf>,
    intestazione: bool,
    test: Vec<TestItem>,
    selezionato: Option<usize>,

    // --- Capture ---
    stato: Stato,
    /// Hypergrab window, to hide it and show it again without depending on the
    /// drawing loop (while it is hidden it receives no WM_PAINT).
    finestra: Option<finestra::Finestra>,
    /// Capture in progress on a separate thread.
    ricevitore: Option<Receiver<EsitoCattura>>,
    /// While it holds, every pass asks for the window to be shown again.
    ripristino_fino: Option<Instant>,
    anteprima: Option<Anteprima>,
    salvato: Option<PathBuf>,

    messaggio: String,
    messaggio_errore: bool,
    info_aperta: bool,
}

impl Hypergrab {
    /// New application; if `elenco_iniziale` is given (command line argument)
    /// the list is opened right away.
    fn new(lang: Language, dark_mode: bool, elenco_iniziale: Option<PathBuf>) -> Self {
        let mut app = Self {
            lang,
            dark_mode,
            tema_applicato: false,
            percorso: None,
            intestazione: true,
            test: Vec::new(),
            selezionato: None,
            stato: Stato::Pronto,
            finestra: None,
            ricevitore: None,
            ripristino_fino: None,
            anteprima: None,
            salvato: None,
            messaggio: t(lang, "Pronto.").to_string(),
            messaggio_errore: false,
            info_aperta: false,
        };
        if let Some(percorso) = elenco_iniziale {
            app.apri_elenco(percorso);
        }
        app
    }

    fn messaggio(&mut self, testo: String, errore: bool) {
        self.messaggio = testo;
        self.messaggio_errore = errore;
    }

    // -----------------------------------------------------------------------
    // Test list
    // -----------------------------------------------------------------------

    /// Opens a test list and shows it. In case of error it reports it without
    /// losing the list already loaded.
    fn apri_elenco(&mut self, percorso: PathBuf) {
        match reader::leggi_elenco(&percorso, self.intestazione) {
            Ok(test) => {
                let nome = percorso
                    .file_name()
                    .map(|n| n.to_string_lossy().to_string())
                    .unwrap_or_default();
                if test.is_empty() {
                    self.messaggio(
                        format!(
                            "{} {} {}",
                            t(self.lang, "Errore di lettura:"),
                            nome,
                            t(self.lang, "Nessun test nel file.")
                        ),
                        true,
                    );
                } else {
                    let n = test.len();
                    let sostantivo =
                        i18n::plural(self.lang, n, "test", "tests");
                    self.messaggio(
                        format!(
                            "{} {nome} — {n} {sostantivo}.",
                            t(self.lang, "Elenco importato da:")
                        ),
                        false,
                    );
                }
                self.percorso = Some(percorso);
                // First test selected: the capture can start right away.
                self.selezionato = if test.is_empty() { None } else { Some(0) };
                self.salvato = None;
                self.test = test;
            }
            Err(e) => {
                self.messaggio(
                    format!("{} {e}", t(self.lang, "Errore di lettura:")),
                    true,
                );
            }
        }
    }

    /// Re-reads the file already open (for example after changing the header
    /// option).
    fn rileggi_elenco(&mut self) {
        if let Some(percorso) = self.percorso.clone() {
            let codice = self.test.get(self.selezionato.unwrap_or(0)).map(|t| t.code.clone());
            self.apri_elenco(percorso);
            if let Some(codice) = codice {
                if let Some(i) = self.test.iter().position(|t| t.code == codice) {
                    self.selezionato = Some(i);
                }
            }
        }
    }

    fn test_selezionato(&self) -> Option<&TestItem> {
        self.selezionato.and_then(|i| self.test.get(i))
    }

    /// Folder where the JPG is saved: the one of the list file.
    fn cartella_salvataggio(&self) -> Option<PathBuf> {
        let percorso = self.percorso.as_ref()?;
        match percorso.parent() {
            Some(cartella) if !cartella.as_os_str().is_empty() => Some(cartella.to_path_buf()),
            _ => Some(PathBuf::from(".")),
        }
    }

    /// Path of the JPG that would be written right now.
    fn destinazione(&self) -> Option<PathBuf> {
        let cartella = self.cartella_salvataggio()?;
        let codice = &self.test_selezionato()?.code;
        Some(capture::percorso_salvataggio(&cartella, codice))
    }

    // -----------------------------------------------------------------------
    // Capture
    // -----------------------------------------------------------------------

    /// Hides the window and starts the capture on a separate thread.
    ///
    /// The capture cannot happen here: while the window is hidden Windows
    /// sends no WM_PAINT and the drawing loop stops, so not even the return
    /// of the window can depend on a later pass.
    fn avvia_cattura(&mut self, ctx: &egui::Context) {
        if self.test_selezionato().is_none() {
            self.messaggio(
                t(self.lang, "Nessun test selezionato.").to_string(),
                true,
            );
            return;
        }
        let Some(finestra) = finestra::Finestra::attiva() else {
            self.messaggio(t(self.lang, "Cattura fallita.").to_string(), true);
            return;
        };

        self.messaggio(
            t(self.lang, "Cattura in corso: la finestra torna a breve…").to_string(),
            false,
        );
        self.stato = Stato::InCorso;
        self.finestra = Some(finestra);

        let (invia, ricevi) = std::sync::mpsc::channel();
        self.ricevitore = Some(ricevi);
        let lang = self.lang;
        let ctx = ctx.clone();
        std::thread::spawn(move || {
            // The window disappears: the application to be documented comes
            // back to the foreground.
            finestra.nascondi();
            std::thread::sleep(ATTESA_CATTURA);
            // If the foreground has not passed to the underlying application
            // yet, we wait a little longer: the capture must not be about
            // Hypergrab.
            let scadenza = Instant::now() + ATTESA_PRIMO_PIANO;
            while Instant::now() < scadenza && !capture::primo_piano_pronto() {
                std::thread::sleep(Duration::from_millis(50));
            }
            let esito = cattura_e_componi(lang);
            // The window becomes visible again even if the capture failed.
            finestra.mostra();
            let _ = invia.send(esito);
            ctx.request_repaint();
        });
    }

    /// Shows the outcome of the capture in the preview.
    fn applica_esito(&mut self, ctx: &egui::Context, esito: EsitoCattura) {
        match esito {
            EsitoCattura::Cattura {
                immagine,
                finestra,
                data_ora,
            } => {
                let texture = ctx.load_texture(
                    "anteprima",
                    immagine_in_colori(&immagine),
                    egui::TextureOptions::LINEAR,
                );
                let finestra = if finestra.is_empty() {
                    t(self.lang, "schermo intero").to_string()
                } else {
                    finestra
                };
                self.anteprima = Some(Anteprima {
                    immagine,
                    texture,
                    finestra: finestra.clone(),
                    data_ora,
                });
                self.salvato = None;
                self.messaggio(
                    format!("{} {finestra}", t(self.lang, "Cattura completata:")),
                    false,
                );
            }
            EsitoCattura::Errore(e) => {
                self.messaggio(
                    format!("{} {e}", t(self.lang, "Cattura fallita.")),
                    true,
                );
            }
        }
    }

    /// Writes the JPG into the folder of the list file.
    fn salva(&mut self) {
        let Some(percorso) = self.destinazione() else {
            self.messaggio(
                t(self.lang, "Nessun file importato.").to_string(),
                true,
            );
            return;
        };
        let Some(anteprima) = self.anteprima.as_ref() else {
            return;
        };
        match capture::salva_jpg(&anteprima.immagine, &percorso) {
            Ok(()) => {
                self.messaggio(
                    format!("{} {}", t(self.lang, "Salvato:"), percorso.display()),
                    false,
                );
                self.salvato = Some(percorso);
            }
            Err(e) => {
                self.messaggio(
                    format!("{} {e}", t(self.lang, "Errore nel salvataggio:")),
                    true,
                );
            }
        }
    }

    // -----------------------------------------------------------------------
    // Interface
    // -----------------------------------------------------------------------

    fn barra_alta(&mut self, ui: &mut egui::Ui) {
        let dark = self.dark_mode;
        ui.horizontal(|ui| {
            // Application symbol: the capture frame.
            let (rect, _) = ui.allocate_exact_size(egui::vec2(26.0, 26.0), egui::Sense::hover());
            theme::draw_icon(ui.painter(), Icon::Capture, rect, theme::accent(dark), 1.8);

            // Name and subtitle, in solid black in the light theme.
            ui.vertical(|ui| {
                ui.spacing_mut().item_spacing.y = 0.0;
                ui.label(
                    egui::RichText::new("Hypergrab")
                        .size(16.0)
                        .strong()
                        .color(theme::text_strong(dark)),
                );
                ui.label(
                    egui::RichText::new(t(self.lang, "Cattura schermo dei test"))
                        .size(10.5)
                        .color(theme::muted(dark)),
                );
            });

            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                let icona = if dark { Icon::Sun } else { Icon::Moon };
                if theme::small_icon_button(ui, icona, ui.visuals().text_color(), 22.0)
                    .on_hover_text(t(self.lang, "Cambia tema"))
                    .clicked()
                {
                    self.dark_mode = !self.dark_mode;
                    ctx_theme(ui.ctx(), self.dark_mode);
                    settings::save(&settings::Settings {
                        dark_mode: self.dark_mode,
                        lang: self.lang,
                    });
                }
                if theme::small_icon_button(ui, Icon::Info, ui.visuals().text_color(), 22.0)
                    .on_hover_text(t(self.lang, "Informazioni"))
                    .clicked()
                {
                    self.info_aperta = true;
                }
                ui.add_space(4.0);
                for lingua in Language::ALL {
                    let attiva = lingua == self.lang;
                    if theme::pill_button(ui, None, lingua.short(), attiva).clicked() && !attiva {
                        self.lang = lingua;
                        self.aggiorna_lingua_messaggi(lingua);
                        settings::save(&settings::Settings {
                            dark_mode: self.dark_mode,
                            lang: lingua,
                        });
                    }
                }
                ui.add_space(4.0);
                ui.label(
                    egui::RichText::new(t(self.lang, "Lingua"))
                        .size(10.5)
                        .color(theme::muted(dark)),
                );
            });
        });
    }

    /// Aligns the messages already shown to the language just chosen.
    fn aggiorna_lingua_messaggi(&mut self, lingua: Language) {
        if self.messaggio == t(Language::It, "Pronto.") {
            self.messaggio = t(lingua, "Pronto.").to_string();
        }
    }

    /// Left column: import and test list.
    fn colonna_elenco(&mut self, ui: &mut egui::Ui, disponibile: egui::Vec2) {
        let dark = self.dark_mode;
        ui.allocate_ui(disponibile, |ui| {
            ui.vertical(|ui| {
                // --- Import ---
                theme::card(ui, |ui| {
                    ui.set_width(ui.available_width());
                    theme::section_title(ui, t(self.lang, "Importa elenco test"));
                    ui.add_space(2.0);
                    if theme::action_button(
                        ui,
                        Icon::Folder,
                        t(self.lang, "Apri file…"),
                    )
                    .clicked()
                    {
                        if let Some(percorso) = rfd::FileDialog::new()
                            .set_title(t(self.lang, "Apri elenco test"))
                            .add_filter("CSV", &["csv", "txt"])
                            .add_filter("Excel", &["xlsx", "xlsm", "xls", "ods"])
                            .pick_file()
                        {
                            self.apri_elenco(percorso);
                        }
                    }
                    ui.add_space(3.0);
                    let path_testo = match &self.percorso {
                        Some(p) => p.display().to_string(),
                        None => t(self.lang, "Nessun file importato").to_string(),
                    };
                    ui.label(
                        egui::RichText::new(tronca(ui, &path_testo, ui.available_width(), 10.5))
                            .size(10.5)
                            .color(theme::muted(dark)),
                    );
                    let prima = ui.checkbox(
                        &mut self.intestazione,
                        t(self.lang, "La prima riga è un'intestazione"),
                    );
                    if prima.changed() {
                        self.rileggi_elenco();
                    }
                });

                ui.add_space(6.0);

                // --- List ---
                theme::card(ui, |ui| {
                    ui.set_width(ui.available_width());
                    theme::section_title(ui, t(self.lang, "Elenco test"));
                    ui.add_space(2.0);
                    if self.test.is_empty() {
                        theme::hint(ui, t(self.lang, "Apri un CSV o un XLSX per vedere i test."));
                        return;
                    }
                    let altezza = ui.available_height().max(120.0);
                    let mut scelta: Option<usize> = None;
                    egui::ScrollArea::vertical()
                        .max_height(altezza)
                        .auto_shrink([false, false])
                        .show(ui, |ui| {
                            ui.set_width(ui.available_width());
                            for (i, item) in self.test.iter().enumerate() {
                                let selezionato = self.selezionato == Some(i);
                                if riga_test(ui, item, selezionato, dark).clicked() {
                                    scelta = Some(i);
                                }
                            }
                        });
                    if let Some(i) = scelta {
                        if self.selezionato != Some(i) {
                            // The test changes: the destination file changes as well.
                            self.salvato = None;
                        }
                        self.selezionato = Some(i);
                    }
                });
            });
        });
    }

    /// Right column: capture command, preview and saving.
    fn colonna_cattura(&mut self, ui: &mut egui::Ui, ctx: &egui::Context) {
        let dark = self.dark_mode;
        let pronto = matches!(self.stato, Stato::Pronto);
        let ha_test = self.test_selezionato().is_some();
        let mut richiesta = false;
        let mut salvataggio = false;

        ui.vertical(|ui| {
            theme::card(ui, |ui| {
                ui.set_width(ui.available_width());
                ui.horizontal(|ui| {
                    ui.add_enabled_ui(pronto && ha_test, |ui| {
                        if theme::big_button(ui, Icon::Capture, t(self.lang, "Esegui cattura"))
                            .clicked()
                        {
                            richiesta = true;
                        }
                    });
                    ui.add_space(6.0);
                    ui.vertical(|ui| {
                        theme::hint(ui, t(self.lang, "oppure premi F12"));
                        match self.test_selezionato() {
                            Some(item) => {
                                let testo = format!(
                                    "{} {}",
                                    t(self.lang, "Test selezionato:"),
                                    item.code
                                );
                                ui.label(
                                    egui::RichText::new(testo)
                                        .size(11.5)
                                        .strong()
                                        .color(theme::text_strong(dark)),
                                );
                                ui.label(
                                    egui::RichText::new(tronca(
                                        ui,
                                        &item.description,
                                        ui.available_width(),
                                        10.5,
                                    ))
                                    .size(10.5)
                                    .color(theme::muted(dark)),
                                );
                            }
                            None => {
                                ui.label(
                                    egui::RichText::new(t(
                                        self.lang,
                                        "Seleziona un test dall'elenco prima di catturare.",
                                    ))
                                    .size(11.0)
                                    .color(theme::muted(dark)),
                                );
                            }
                        }
                    });
                });
            });

            ui.add_space(6.0);

            theme::card(ui, |ui| {
                ui.set_width(ui.available_width());
                theme::section_title(ui, t(self.lang, "Anteprima"));
                ui.add_space(2.0);

                match self.anteprima.as_ref() {
                    Some(anteprima) => {
                        let (w, h) = anteprima.immagine.dimensions();
                        let spazio = ui.available_size() - egui::vec2(0.0, 30.0);
                        let scala = (spazio.x / w as f32)
                            .min(spazio.y / h as f32)
                            .min(1.0)
                            .max(0.05);
                        ui.add(egui::Image::new((
                            anteprima.texture.id(),
                            egui::vec2(w as f32 * scala, h as f32 * scala),
                        )));
                        ui.add_space(4.0);
                        let info = format!(
                            "{} {} — {}",
                            t(self.lang, "Cattura completata:"),
                            anteprima.finestra,
                            anteprima.data_ora,
                        );
                        theme::hint(ui, &tronca(ui, &info, ui.available_width(), 10.5));
                    }
                    None => {
                        theme::hint(ui, t(self.lang, "Nessuna anteprima: esegui una cattura."));
                    }
                }

                ui.add_space(6.0);
                ui.horizontal(|ui| {
                    ui.add_enabled_ui(self.anteprima.is_some() && ha_test, |ui| {
                        if theme::action_button(
                            ui,
                            Icon::Save,
                            t(self.lang, "Salva JPG"),
                        )
                        .clicked()
                        {
                            salvataggio = true;
                        }
                    });
                    ui.add_space(6.0);
                    if let Some(destinazione) = self.destinazione() {
                        // The path already written in this session is shown
                        // in the highlight color, without any overwrite
                        // warning.
                        let appena_salvato = self.salvato.as_deref() == Some(destinazione.as_path());
                        let esistente = destinazione.exists();
                        let testo = tronca(
                            ui,
                            &destinazione.display().to_string(),
                            ui.available_width() - 8.0,
                            10.5,
                        );
                        if appena_salvato {
                            ui.label(
                                egui::RichText::new(testo)
                                    .size(10.5)
                                    .color(theme::accent(dark)),
                            );
                        } else {
                            theme::hint(ui, &testo);
                            if esistente {
                                let avviso =
                                    t(self.lang, "Il file esistente verrà sovrascritto.");
                                ui.label(
                                    egui::RichText::new(avviso)
                                        .size(10.5)
                                        .color(theme::WARNING),
                                );
                            }
                        }
                    }
                });
            });
        });

        if richiesta {
            self.avvia_cattura(ctx);
        }
        if salvataggio {
            self.salva();
        }
    }

    fn barra_stato(&self, ui: &mut egui::Ui) {
        let dark = self.dark_mode;
        ui.horizontal(|ui| {
            let colore = if self.messaggio_errore {
                theme::DANGER
            } else {
                ui.visuals().text_color()
            };
            ui.label(
                egui::RichText::new(tronca(ui, &self.messaggio, ui.available_width() - 220.0, 11.0))
                    .size(11.0)
                    .color(colore),
            );
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                if let Some(percorso) = &self.percorso {
                    let nome = percorso
                        .file_name()
                        .map(|n| n.to_string_lossy().to_string())
                        .unwrap_or_default();
                    ui.label(
                        egui::RichText::new(nome)
                            .size(10.5)
                            .color(theme::muted(dark)),
                    );
                }
                if !self.test.is_empty() {
                    let n = self.test.len();
                    let sostantivo = i18n::plural(self.lang, n, "test", "tests");
                    ui.label(
                        egui::RichText::new(format!("{n} {sostantivo}"))
                            .size(10.5)
                            .color(theme::muted(dark)),
                    );
                }
            });
        });
    }

    fn finestra_info(&mut self, ctx: &egui::Context) {
        let mut aperta = self.info_aperta;
        egui::Window::new(t(self.lang, "Informazioni"))
            .open(&mut aperta)
            .collapsible(false)
            .resizable(false)
            .default_width(420.0)
            .anchor(egui::Align2::CENTER_CENTER, egui::vec2(0.0, 0.0))
            .show(ctx, |ui| {
                ui.set_max_width(400.0);
                for chiave in [
                    "Hypergrab è uno strumento di cattura di schermo pensato per accompagnare il processo di acquisizione di evidenze durante la qualificazione o la validazione di sistemi informatizzati.",
                    "Permette di importare un elenco di test da un file CSV o XLSX (prima colonna: codice del test, seconda colonna: descrizione) e di catturare la finestra dell'applicazione che resta in primo piano nascondendo temporaneamente Hypergrab.",
                    "All'immagine viene aggiunta in basso a destra la data e l'ora di acquisizione; il file JPG viene salvato nella cartella del file di elenco, con nome uguale al codice del test.",
                    "Sviluppato in Rust con Egui come sfida personale. Distribuito con licenza GNU GPL v3. Autore: Daniel Consonni. Un ringraziamento speciale a Geovanna Mendes de Souza, le cui osservazioni sono state il punto di partenza per lo sviluppo di questa applicazione.",
                ] {
                    ui.label(egui::RichText::new(t(self.lang, chiave)).size(11.5));
                    ui.add_space(5.0);
                }
                ui.separator();
                ui.hyperlink_to(
                    "github.com/danico-oss/hypergrab",
                    "https://github.com/danico-oss/hypergrab",
                );
            });
        self.info_aperta = aperta;
    }
}

impl eframe::App for Hypergrab {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        // Compact scale, also suitable for laptop screens.
        ctx.set_pixels_per_point(1.0);

        if !self.tema_applicato {
            theme::apply(ctx, self.dark_mode);
            self.tema_applicato = true;
        }

        // --- Outcome of the capture running on a separate thread ---
        let esito = self
            .ricevitore
            .as_ref()
            .map(|ricevitore| ricevitore.try_recv());
        match esito {
            Some(Ok(esito)) => {
                self.ricevitore = None;
                self.stato = Stato::Pronto;
                // The window has already been shown again by the capture thread:
                // we insist a little longer and request the foreground.
                self.ripristino_fino = Some(Instant::now() + RIPRISTINO);
                ctx.send_viewport_cmd(egui::ViewportCommand::Focus);
                self.applica_esito(ctx, esito);
            }
            Some(Err(TryRecvError::Empty)) => {
                ctx.request_repaint_after(Duration::from_millis(30));
            }
            Some(Err(TryRecvError::Disconnected)) => {
                // The capture thread died without producing an outcome.
                self.ricevitore = None;
                self.stato = Stato::Pronto;
                self.ripristino_fino = Some(Instant::now() + RIPRISTINO);
                self.messaggio(t(self.lang, "Cattura fallita.").to_string(), true);
            }
            None => {}
        }

        // --- Restoring the window after the capture ---
        // If the first attempt were lost the window would stay hidden: it is
        // shown again on every pass for a short while.
        if let Some(fine) = self.ripristino_fino {
            if let Some(finestra) = self.finestra {
                finestra.mostra();
            }
            if Instant::now() >= fine {
                self.ripristino_fino = None;
                self.finestra = None;
            } else {
                ctx.request_repaint_after(Duration::from_millis(40));
            }
        }

        // --- F12 key ---
        // Ignored while the window is coming back to view: a key left held
        // down must not start another capture right away.
        let f12 = ctx.input(|i| i.key_pressed(egui::Key::F12));
        if f12 && matches!(self.stato, Stato::Pronto) && self.ripristino_fino.is_none() {
            self.avvia_cattura(ctx);
        }

        // --- File dragged onto the window ---
        let trascinato = ctx.input(|i| i.raw.dropped_files.iter().find_map(|f| f.path.clone()));
        if let Some(percorso) = trascinato {
            self.apri_elenco(percorso);
        }

        let dark = self.dark_mode;

        // --- Top bar (theme background) ---
        let frame_alto = egui::Frame::new()
            .fill(theme::ribbon_bg(dark))
            .inner_margin(egui::Margin {
                left: 8,
                right: 8,
                top: 4,
                bottom: 4,
            });
        egui::TopBottomPanel::top("barra_alta")
            .frame(frame_alto)
            .show(ctx, |ui| {
                self.barra_alta(ui);
            });

        // --- Status bar ---
        let frame_basso = egui::Frame::new()
            .fill(theme::ribbon_bg(dark))
            .inner_margin(egui::Margin {
                left: 8,
                right: 8,
                top: 2,
                bottom: 2,
            });
        egui::TopBottomPanel::bottom("barra_stato")
            .frame(frame_basso)
            .show(ctx, |ui| {
                self.barra_stato(ui);
            });

        // --- Central area: list on the left, capture on the right ---
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.add_space(4.0);
            let disponibile = ui.available_size();
            let larghezza_elenco = (disponibile.x * 0.42).clamp(260.0, 420.0);
            ui.horizontal_top(|ui| {
                self.colonna_elenco(
                    ui,
                    egui::vec2(larghezza_elenco, disponibile.y - 6.0),
                );
                ui.add_space(6.0);
                // The right column takes all the remaining width.
                let rimasto = ui.available_width();
                ui.allocate_ui(egui::vec2(rimasto, disponibile.y - 6.0), |ui| {
                    self.colonna_cattura(ui, ctx);
                });
            });
        });

        self.finestra_info(ctx);
    }

    /// Background color of the window: it follows the theme, so that any
    /// transparent areas do not show the default dark grey of eframe.
    fn clear_color(&self, _visuals: &egui::Visuals) -> [f32; 4] {
        theme::panel_bg(self.dark_mode).to_normalized_gamma_f32()
    }
}

/// Applies the theme and keeps track of the application (used by the theme
/// button in the top bar).
fn ctx_theme(ctx: &egui::Context, dark: bool) {
    theme::apply(ctx, dark);
}

/// Captures the window that stays in the foreground and prints date and time on it.
///
/// It runs on a separate thread: the Hypergrab window is hidden, so at that
/// moment the drawing loop of the application must do nothing.
fn cattura_e_componi(lang: Language) -> EsitoCattura {
    let mut cattura = match capture::cattura_finestra_primo_piano() {
        Ok(cattura) => cattura,
        Err(e) => return EsitoCattura::Errore(e),
    };
    let data_ora = capture::testo_data_ora(lang, chrono::Local::now());
    if let Err(e) = capture::compone_data_ora(&mut cattura.immagine, &data_ora) {
        return EsitoCattura::Errore(e);
    }
    EsitoCattura::Cattura {
        immagine: cattura.immagine,
        finestra: cattura.finestra,
        data_ora,
    }
}

/// Converts the captured image into the egui format.
fn immagine_in_colori(immagine: &image::RgbImage) -> egui::ColorImage {
    egui::ColorImage::from_rgb(
        [immagine.width() as usize, immagine.height() as usize],
        immagine.as_raw(),
    )
}

/// One row of the test list: code and description, selection as a frame.
fn riga_test(ui: &mut egui::Ui, item: &TestItem, selezionato: bool, dark: bool) -> egui::Response {
    let altezza = 38.0;
    let (rect, risposta) =
        ui.allocate_exact_size(egui::vec2(ui.available_width(), altezza), egui::Sense::click());

    if risposta.hovered() {
        ui.painter()
            .rect_filled(rect, egui::CornerRadius::same(3), theme::hover_bg(dark));
    }

    let larghezza_testo = rect.width() - 16.0;
    ui.painter().text(
        egui::pos2(rect.left() + 8.0, rect.top() + 7.0),
        egui::Align2::LEFT_TOP,
        tronca(ui, &item.code, larghezza_testo, 12.5),
        egui::FontId::proportional(12.5),
        theme::text_strong(dark),
    );
    ui.painter().text(
        egui::pos2(rect.left() + 8.0, rect.top() + 22.0),
        egui::Align2::LEFT_TOP,
        tronca(ui, &item.description, larghezza_testo, 10.5),
        egui::FontId::proportional(10.5),
        theme::muted(dark),
    );

    // The selection is always and only a frame.
    if selezionato {
        theme::selection_frame(ui.painter(), rect, egui::CornerRadius::same(3), dark);
    }
    if risposta.hovered() {
        ui.ctx().set_cursor_icon(egui::CursorIcon::PointingHand);
    }
    risposta
}

/// Shortens a text with dots if it does not fit the available width.
fn tronca(ui: &egui::Ui, testo: &str, larghezza: f32, dimensione: f32) -> String {
    if larghezza <= 0.0 {
        return String::new();
    }
    let font = egui::FontId::proportional(dimensione);
    let misura = |s: &str| {
        ui.fonts(|f| {
            f.layout_no_wrap(s.to_owned(), font.clone(), egui::Color32::BLACK)
                .size()
                .x
        })
    };
    if misura(testo) <= larghezza {
        return testo.to_string();
    }
    let mut accorciato: String = testo.to_string();
    while !accorciato.is_empty() && misura(&format!("{accorciato}…")) > larghezza {
        accorciato.pop();
    }
    format!("{accorciato}…")
}

fn main() -> eframe::Result<()> {
    let preferenze = settings::load();
    // Path possibly given on the command line: it is opened right away.
    let elenco_iniziale = std::env::args_os()
        .nth(1)
        .map(PathBuf::from)
        .filter(|p| p.is_file());
    let sottotitolo = t(preferenze.lang, "Cattura schermo dei test");
    let titolo = format!("Hypergrab — {sottotitolo}");

    let opzioni = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size((980.0, 620.0))
            .with_min_inner_size((720.0, 480.0))
            .with_icon(icon::egui_icon(64))
            .with_title(&titolo),
        ..eframe::NativeOptions::default()
    };

    eframe::run_native(
        &titolo,
        opzioni,
        Box::new(move |_| {
            Ok(Box::new(Hypergrab::new(
                preferenze.lang,
                preferenze.dark_mode,
                elenco_iniziale,
            )))
        }),
    )
}
