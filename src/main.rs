#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use calamine::{Reader, Xlsx, open_workbook};
use eframe::egui;
use image::DynamicImage;
use std::path::PathBuf;
use std::time::{Duration, Instant};
use xcap::Monitor;

struct TestItem {
    codice: String,
    descrizione: String,
}

#[derive(PartialEq)]
enum AppState {
    Idle,
    RequestMinimize,
    WaitingForMinimize(Instant),
    CaptureAndSave,
    Restore,
}

struct MyApp {
    excel_path: Option<PathBuf>,
    items: Vec<TestItem>,
    selected_index: Option<usize>,
    status_message: String,
    state: AppState,
}

impl Default for MyApp {
    fn default() -> Self {
        Self {
            excel_path: None,
            items: Vec::new(),
            selected_index: None,
            status_message: "Select an Excel file to start".to_string(),
            state: AppState::Idle,
        }
    }
}

impl MyApp {
    fn load_excel(&mut self, path: PathBuf) {
        let mut workbook: Xlsx<_> = match open_workbook(&path) {
            Ok(w) => w,
            Err(e) => {
                self.status_message = format!("Error: {}", e);
                return;
            }
        };

        if let Some(Ok(range)) = workbook.worksheet_range_at(0) {
            self.items.clear();
            for row in range.rows().skip(1) {
                let codice = row.get(0).map(|c| c.to_string()).unwrap_or_default();
                let descrizione = row.get(1).map(|c| c.to_string()).unwrap_or_default();
                if !codice.is_empty() {
                    self.items.push(TestItem {
                        codice,
                        descrizione,
                    });
                }
            }
            self.excel_path = Some(path);
            self.status_message = format!("Loaded {} tests", self.items.len());
        }
    }

    fn execute_capture_logic(&mut self) {
        let idx = self.selected_index.expect("Select a test");
        let item = &self.items[idx];
        let dir = self.excel_path.as_ref().unwrap().parent().unwrap();

        // Calcolo nome file sequenziale
        let mut final_path = dir.join(format!("{}.jpg", item.codice));
        let mut counter = 1;
        while final_path.exists() {
            final_path = dir.join(format!("{}_{}.jpg", item.codice, counter));
            counter += 1;
        }

        // Cattura schermo
        let monitors = Monitor::all().unwrap();
        if let Some(monitor) = monitors.first() {
            match monitor.capture_image() {
                Ok(image) => {
                    let dynamic_img = DynamicImage::ImageRgba8(image);
                    match dynamic_img.to_rgb8().save(&final_path) {
                        Ok(_) => {
                            self.status_message = format!(
                                "✅ Saved: {}",
                                final_path.file_name().unwrap().to_string_lossy()
                            )
                        }
                        Err(e) => self.status_message = format!("❌ Save error: {}", e),
                    }
                }
                Err(e) => self.status_message = format!("❌ Capture error: {}", e),
            }
        }
    }
}

impl eframe::App for MyApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        // --- GESTIONE LOGICA STATI ---
        match self.state {
            AppState::RequestMinimize => {
                ctx.send_viewport_cmd(egui::ViewportCommand::Minimized(true));
                self.state = AppState::WaitingForMinimize(Instant::now());
            }
            AppState::WaitingForMinimize(start_time) => {
                // Attesa 400ms per permettere alla finestra di sparire del tutto
                if start_time.elapsed() >= Duration::from_millis(400) {
                    self.state = AppState::CaptureAndSave;
                }
                ctx.request_repaint();
            }
            AppState::CaptureAndSave => {
                // Esegue la cattura e il salvataggio mentre è ancora minimizzata
                self.execute_capture_logic();
                self.state = AppState::Restore;
                ctx.request_repaint();
            }
            AppState::Restore => {
                // Solo ora ripristiniamo la finestra
                ctx.send_viewport_cmd(egui::ViewportCommand::Minimized(false));
                ctx.send_viewport_cmd(egui::ViewportCommand::Focus);
                self.state = AppState::Idle;
            }
            AppState::Idle => {}
        }

        // --- INTERFACCIA UTENTE ---
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("📸 Excel Test Capture");

            if ui.button("📂 Load Excel file").clicked() {
                if let Some(path) = rfd::FileDialog::new()
                    .add_filter("Excel", &["xlsx"])
                    .pick_file()
                {
                    self.load_excel(path);
                }
            }

            ui.separator();

            if self.items.is_empty() {
                ui.label("Load a file to see the tests.");
            } else {
                egui::ScrollArea::vertical()
                    .max_height(350.0)
                    .show(ui, |ui| {
                        for (i, item) in self.items.iter().enumerate() {
                            let is_selected = self.selected_index == Some(i);
                            if ui
                                .selectable_label(
                                    is_selected,
                                    format!("[{}] {}", item.codice, item.descrizione),
                                )
                                .clicked()
                            {
                                self.selected_index = Some(i);
                            }
                        }
                    });
            }

            ui.separator();

            if let Some(_) = self.selected_index {
                // Disabilita il bottone se siamo già in fase di cattura
                let enabled = self.state == AppState::Idle;
                ui.add_enabled_ui(enabled, |ui| {
                    if ui
                        .add_sized(
                            [ui.available_width(), 40.0],
                            egui::Button::new("📸 CAPTURE"),
                        )
                        .clicked()
                    {
                        self.state = AppState::RequestMinimize;
                    }
                });
            }

            ui.add_space(10.0);
            ui.label(egui::RichText::new(&self.status_message).color(egui::Color32::LIGHT_BLUE));
        });
    }
}

fn main() -> eframe::Result {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([500.0, 300.0])
            .with_title("Test Capture Tool"),
        ..Default::default()
    };

    eframe::run_native(
        "Excel Screen Capture",
        options,
        Box::new(|_cc| Ok(Box::new(MyApp::default()))),
    )
}
