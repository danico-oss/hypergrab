//! Hypergrab: cattura di schermo dei test.
//!
//! L'applicazione importa un elenco di test da un file CSV o XLSX (prima
//! colonna il codice, seconda colonna la descrizione), permette di selezionare
//! un test e, con il pulsante "Esegui cattura" o il tasto F12, si nasconde per
//! fotografare la finestra che resta in primo piano. All'immagine vengono uniti
//! data e ora in basso a destra, poi il JPG viene salvato nella cartella del
//! file di elenco con il nome del codice del test.

mod capture;
mod i18n;
mod icon;
mod reader;
mod settings;
mod theme;

use eframe::egui;
use i18n::{t, Language};
use reader::TestItem;
use std::path::PathBuf;
use std::time::{Duration, Instant};
use theme::Icon;

/// Attesa fra il momento in cui la finestra si nasconde e la cattura: serve a
/// lasciare che l'applicazione sottostante torni in primo piano.
const ATTESA_CATTURA: Duration = Duration::from_millis(650);

enum Stato {
    /// Finestra visibile, nessuna cattura in corso.
    Pronto,
    /// Finestra nascosta: la cattura avviene allo scadere dell'attesa.
    InAttesa { scadenza: Instant },
}

/// Immagine catturata, pronta per l'anteprima e per il salvataggio.
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

    // --- Elenco dei test ---
    percorso: Option<PathBuf>,
    intestazione: bool,
    test: Vec<TestItem>,
    selezionato: Option<usize>,

    // --- Cattura ---
    stato: Stato,
    anteprima: Option<Anteprima>,
    salvato: Option<PathBuf>,

    messaggio: String,
    messaggio_errore: bool,
    info_aperta: bool,
}

impl Hypergrab {
    fn new(lang: Language, dark_mode: bool) -> Self {
        Self {
            lang,
            dark_mode,
            tema_applicato: false,
            percorso: None,
            intestazione: true,
            test: Vec::new(),
            selezionato: None,
            stato: Stato::Pronto,
            anteprima: None,
            salvato: None,
            messaggio: t(lang, "Pronto.").to_string(),
            messaggio_errore: false,
            info_aperta: false,
        }
    }

    fn messaggio(&mut self, testo: String, errore: bool) {
        self.messaggio = testo;
        self.messaggio_errore = errore;
    }

    // -----------------------------------------------------------------------
    // Elenco dei test
    // -----------------------------------------------------------------------

    /// Apre un elenco di test e lo mostra. In caso di errore lo segnala senza
    /// perdere l'elenco gia' caricato.
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
                // Primo test selezionato: si puo' catturare subito.
                self.selezionato = if test.is_empty() { None } else { Some(0) };
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

    /// Rilegge il file gia' aperto (per esempio dopo aver cambiato l'opzione
    /// sull'intestazione).
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

    /// Cartella in cui salvare il JPG: quella del file di elenco.
    fn cartella_salvataggio(&self) -> Option<PathBuf> {
        let percorso = self.percorso.as_ref()?;
        match percorso.parent() {
            Some(cartella) if !cartella.as_os_str().is_empty() => Some(cartella.to_path_buf()),
            _ => Some(PathBuf::from(".")),
        }
    }

    /// Percorso del JPG che verrebbe scritto adesso.
    fn destinazione(&self) -> Option<PathBuf> {
        let cartella = self.cartella_salvataggio()?;
        let codice = &self.test_selezionato()?.code;
        Some(capture::percorso_salvataggio(&cartella, codice))
    }

    // -----------------------------------------------------------------------
    // Cattura
    // -----------------------------------------------------------------------

    /// Nasconde la finestra e programma la cattura.
    fn avvia_cattura(&mut self, ctx: &egui::Context) {
        if self.test_selezionato().is_none() {
            self.messaggio(
                t(self.lang, "Nessun test selezionato.").to_string(),
                true,
            );
            return;
        }
        self.messaggio(
            t(self.lang, "Cattura in corso: la finestra torna a breve…").to_string(),
            false,
        );
        self.stato = Stato::InAttesa {
            scadenza: Instant::now() + ATTESA_CATTURA,
        };
        ctx.send_viewport_cmd(egui::ViewportCommand::Visible(false));

        // Con la finestra nascosta i ridisegni potrebbero diradarsi: un
        // risveglio programmato garantisce che la cattura avvenga comunque.
        let ctx = ctx.clone();
        std::thread::spawn(move || {
            std::thread::sleep(ATTESA_CATTURA);
            ctx.request_repaint();
        });
    }

    /// Esegue la cattura, compone data e ora e riporta la finestra in vista.
    fn esegui_cattura(&mut self, ctx: &egui::Context) {
        let esito = capture::cattura_finestra_primo_piano();
        ctx.send_viewport_cmd(egui::ViewportCommand::Visible(true));
        ctx.send_viewport_cmd(egui::ViewportCommand::Focus);
        self.stato = Stato::Pronto;

        let mut cattura = match esito {
            Ok(cattura) => cattura,
            Err(e) => {
                self.messaggio(
                    format!("{} {e}", t(self.lang, "Cattura fallita.")),
                    true,
                );
                return;
            }
        };

        let data_ora = capture::testo_data_ora(self.lang, chrono::Local::now());
        if let Err(e) = capture::compone_data_ora(&mut cattura.immagine, &data_ora) {
            self.messaggio(
                format!("{} {e}", t(self.lang, "Cattura fallita.")),
                true,
            );
            return;
        }

        let texture = ctx.load_texture(
            "anteprima",
            immagine_in_colori(&cattura.immagine),
            egui::TextureOptions::LINEAR,
        );
        let finestra = if cattura.finestra.is_empty() {
            t(self.lang, "schermo intero").to_string()
        } else {
            cattura.finestra
        };
        self.anteprima = Some(Anteprima {
            immagine: cattura.immagine,
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

    /// Scrive il JPG nella cartella del file di elenco.
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
    // Interfaccia
    // -----------------------------------------------------------------------

    fn barra_alta(&mut self, ui: &mut egui::Ui) {
        let dark = self.dark_mode;
        ui.horizontal(|ui| {
            // Simbolo dell'applicazione: la cornice di cattura.
            let (rect, _) = ui.allocate_exact_size(egui::vec2(26.0, 26.0), egui::Sense::hover());
            theme::draw_icon(ui.painter(), Icon::Capture, rect, theme::accent(dark), 1.8);

            // Nome e sottotitolo, in nero pieno nel tema chiaro.
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

    /// Allinea i messaggi gia' mostrati alla lingua appena scelta.
    fn aggiorna_lingua_messaggi(&mut self, lingua: Language) {
        if self.messaggio == t(Language::It, "Pronto.") {
            self.messaggio = t(lingua, "Pronto.").to_string();
        }
    }

    /// Colonna di sinistra: importazione ed elenco dei test.
    fn colonna_elenco(&mut self, ui: &mut egui::Ui, disponibile: egui::Vec2) {
        let dark = self.dark_mode;
        ui.allocate_ui(disponibile, |ui| {
            ui.vertical(|ui| {
                // --- Importazione ---
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

                // --- Elenco ---
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
                        self.selezionato = Some(i);
                    }
                });
            });
        });
    }

    /// Colonna di destra: comando di cattura, anteprima e salvataggio.
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
                            "{} {} — {}{}",
                            t(self.lang, "Cattura completata:"),
                            anteprima.finestra,
                            anteprima.data_ora,
                            match &self.salvato {
                                Some(_) => "",
                                None => "",
                            }
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
                        let esistente = destinazione.exists();
                        let testo = destinazione.display().to_string();
                        theme::hint(ui, &tronca(ui, &testo, ui.available_width() - 8.0, 10.5));
                        if esistente {
                            let avviso = t(self.lang, "Il file esistente verrà sovrascritto.");
                            ui.label(
                                egui::RichText::new(avviso)
                                    .size(10.5)
                                    .color(theme::WARNING),
                            );
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
        // Scala compatta, adatta anche agli schermi dei notebook.
        ctx.set_pixels_per_point(1.0);

        if !self.tema_applicato {
            theme::apply(ctx, self.dark_mode);
            self.tema_applicato = true;
        }

        // --- Cattura programmata ---
        if let Stato::InAttesa { scadenza } = self.stato {
            if Instant::now() >= scadenza {
                self.esegui_cattura(ctx);
            } else {
                ctx.request_repaint_after(Duration::from_millis(30));
            }
        }

        // --- Tasto F12 ---
        let f12 = ctx.input(|i| i.key_pressed(egui::Key::F12));
        if f12 && matches!(self.stato, Stato::Pronto) {
            self.avvia_cattura(ctx);
        }

        let dark = self.dark_mode;

        // --- Barra alta (sfondo del tema) ---
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

        // --- Barra di stato ---
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

        // --- Area centrale: elenco a sinistra, cattura a destra ---
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
                // La colonna di destra prende tutta la larghezza rimasta.
                let rimasto = ui.available_width();
                ui.allocate_ui(egui::vec2(rimasto, disponibile.y - 6.0), |ui| {
                    self.colonna_cattura(ui, ctx);
                });
            });
        });

        self.finestra_info(ctx);
    }

    /// Colore di fondo della finestra: segue il tema, cosi' eventuali aree
    /// trasparenti non mostrano il grigio scuro predefinito di eframe.
    fn clear_color(&self, _visuals: &egui::Visuals) -> [f32; 4] {
        theme::panel_bg(self.dark_mode).to_normalized_gamma_f32()
    }
}

/// Applica il tema e tiene traccia dell'applicazione (usata dal pulsante del
/// tema nella barra alta).
fn ctx_theme(ctx: &egui::Context, dark: bool) {
    theme::apply(ctx, dark);
}

/// Converte l'immagine catturata nel formato di egui.
fn immagine_in_colori(immagine: &image::RgbImage) -> egui::ColorImage {
    egui::ColorImage::from_rgb(
        [immagine.width() as usize, immagine.height() as usize],
        immagine.as_raw(),
    )
}

/// Una riga dell'elenco dei test: codice e descrizione, selezione a cornice.
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

    // La selezione e' sempre e solo una cornice.
    if selezionato {
        theme::selection_frame(ui.painter(), rect, egui::CornerRadius::same(3), dark);
    }
    if risposta.hovered() {
        ui.ctx().set_cursor_icon(egui::CursorIcon::PointingHand);
    }
    risposta
}

/// Accorcia un testo con i puntini se non entra nella larghezza disponibile.
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
            )))
        }),
    )
}
