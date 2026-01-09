#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use calamine::{Reader, Xlsx, open_workbook};
use iced::keyboard;
use iced::widget::{button, column, container, image, row, scrollable, text};
use iced::{Border, Element, Event, Length, Settings, Size, Task, Theme, window};
use std::path::PathBuf;
use std::time::Duration;
use xcap::Monitor;

pub fn main() -> iced::Result {
    MyApp::run(Settings::default())
}

#[derive(Debug, Clone)]
enum Message {
    OpenFile,
    FileSelected(Option<PathBuf>),
    SelectItem(usize),
    StartCapture,
    TickCapture,
    CaptureFinished(Result<PathBuf, String>),
    OpenLastCapture,
    OpenFolder,
    SetView(View),
    KeyPressed(keyboard::Key),
    Init(window::Id),
}

#[derive(PartialEq, Debug, Clone, Copy)]
enum View {
    Main,
    Info,
}

#[derive(PartialEq, Clone)]
enum AppState {
    Idle,
    Waiting,
}

struct TestItem {
    code: String,
    description: String,
}

struct MyApp {
    window_id: Option<window::Id>,
    excel_path: Option<PathBuf>,
    last_capture_path: Option<PathBuf>,
    items: Vec<TestItem>,
    selected_index: Option<usize>,
    status_message: String,
    state: AppState,
    current_view: View,
}

impl MyApp {
    fn new() -> (Self, Task<Message>) {
        (
            Self {
                window_id: None,
                excel_path: None,
                last_capture_path: None,
                items: Vec::new(),
                selected_index: None,
                status_message: String::from("System ready. Load an Excel file to begin."),
                state: AppState::Idle,
                current_view: View::Main,
            },
            Task::none(),
        )
    }

    fn update(&mut self, message: Message) -> Task<Message> {
        match message {
            Message::Init(id) => {
                if self.window_id.is_none() {
                    self.window_id = Some(id);
                }
            }
            Message::KeyPressed(key) => {
                if let keyboard::Key::Named(keyboard::key::Named::F12) = key {
                    if self.selected_index.is_some() && self.state == AppState::Idle {
                        return self.update(Message::StartCapture);
                    }
                }
            }
            Message::StartCapture => {
                if let Some(id) = self.window_id {
                    self.state = AppState::Waiting;
                    self.status_message = String::from("Action: Minimizing and capturing...");
                    return Task::batch(vec![
                        window::minimize(id, true),
                        Task::perform(
                            async { tokio::time::sleep(Duration::from_millis(1500)).await },
                            |_| Message::TickCapture,
                        ),
                    ]);
                } else {
                    self.status_message = String::from("Error: Window ID missing.");
                }
            }
            Message::TickCapture => {
                if let (Some(idx), Some(path)) = (self.selected_index, &self.excel_path) {
                    let code = self.items[idx].code.clone();
                    let path = path.clone();
                    return Task::perform(
                        async move { Self::async_capture(path, code).await },
                        Message::CaptureFinished,
                    );
                }
            }
            Message::CaptureFinished(result) => {
                self.state = AppState::Idle;
                match result {
                    Ok(path) => {
                        self.status_message =
                            format!("File saved: {:?}", path.file_name().unwrap());
                        self.last_capture_path = Some(path);
                    }
                    Err(e) => self.status_message = format!("Error: {}", e),
                }
                if let Some(id) = self.window_id {
                    return window::minimize(id, false);
                }
            }
            Message::OpenLastCapture => {
                if let Some(ref path) = self.last_capture_path {
                    let p = path.clone();
                    #[cfg(target_os = "windows")]
                    let _ = std::process::Command::new("explorer").arg(p).spawn();
                    #[cfg(target_os = "linux")]
                    let _ = std::process::Command::new("xdg-open").arg(p).spawn();
                }
            }
            Message::OpenFile => {
                return Task::perform(
                    async {
                        rfd::FileDialog::new()
                            .add_filter("Excel", &["xlsx"])
                            .pick_file()
                    },
                    Message::FileSelected,
                );
            }
            Message::FileSelected(Some(path)) => self.load_excel(path),
            Message::SelectItem(index) => self.selected_index = Some(index),
            Message::OpenFolder => {
                if let Some(ref path) = self.excel_path {
                    if let Some(dir) = path.parent() {
                        #[cfg(target_os = "windows")]
                        let _ = std::process::Command::new("explorer").arg(dir).spawn();
                        #[cfg(target_os = "linux")]
                        let _ = std::process::Command::new("xdg-open").arg(dir).spawn();
                    }
                }
            }
            Message::SetView(v) => self.current_view = v,
            _ => {}
        }
        Task::none()
    }

    fn view(&self) -> Element<'_, Message> {
        let nav = container(
            row![
                button(text("DASHBOARD").size(12))
                    .on_press(Message::SetView(View::Main))
                    .padding([8, 20]),
                button(text("SYSTEM INFO").size(12))
                    .on_press(Message::SetView(View::Info))
                    .padding([8, 20]),
            ]
            .spacing(10),
        )
        .width(Length::Fill);

        let content: Element<'_, Message> = match self.current_view {
            View::Main => {
                let mut list_col = column![].spacing(2).width(Length::Fill);

                for (i, item) in self.items.iter().enumerate() {
                    let is_selected = self.selected_index == Some(i);
                    list_col = list_col.push(
                        button(
                            row![
                                text(item.code.clone()).width(Length::Fixed(100.0)),
                                text(item.description.clone()).size(14),
                            ]
                            .spacing(20),
                        )
                        .width(Length::Fill)
                        .padding(12)
                        .on_press(Message::SelectItem(i))
                        .style(if is_selected {
                            button::primary
                        } else {
                            button::text
                        }),
                    );
                }

                let scroll_list = container(scrollable(list_col))
                    .height(300)
                    .padding(5)
                    .style(|theme: &Theme| {
                        let palette = theme.extended_palette();
                        container::Style {
                            background: Some(palette.background.weak.color.into()),
                            border: Border {
                                color: palette.background.strong.color,
                                width: 1.0,
                                radius: 4.0.into(),
                            },
                            ..Default::default()
                        }
                    });

                let mut main_view_col = column![
                    row![
                        button(text("Import Excel").size(14))
                            .on_press(Message::OpenFile)
                            .padding([10, 20]),
                        button(text("Open Directory").size(14))
                            .on_press(Message::OpenFolder)
                            .padding([10, 20]),
                    ]
                    .spacing(10),
                    scroll_list,
                ]
                .spacing(20);

                if let Some(ref path) = self.last_capture_path {
                    main_view_col = main_view_col.push(
                        column![
                            text("LAST CAPTURE (Click to open):").size(11),
                            button(
                                image(path.clone())
                                    .width(Length::Fixed(160.0))
                                    .height(Length::Fixed(90.0))
                            )
                            .on_press(Message::OpenLastCapture)
                            .style(button::text)
                        ]
                        .spacing(5),
                    );
                }

                main_view_col
                    .push(
                        button(text("EXECUTE CAPTURE (F12)").size(16))
                            .width(Length::Fill)
                            .padding(15)
                            .on_press(Message::StartCapture)
                            .style(button::primary),
                    )
                    .into()
            }
            View::Info => container(
                column![
                    text("Screen Capture Manager").size(24),
                    text("This software was developed to assist in the process of obtaining")
                        .size(12),
                    text(
                        "evidence during the qualification and validation of computerized systems."
                    )
                    .size(12),
                    text("\n").size(16),
                    text("Architecture: Rust + Iced 0.13").size(16),
                    text("Version: 0.99.0-STABLE").size(14),
                    text("Platform: Windows & Linux compatible").size(14),
                    text("\n").size(16),
                    text("https://github.com/danico-oss/hypergrab")
                        .size(14)
                        .color([0.3, 0.5, 1.0]),
                    text("distributed under the GPL v. 3.0 license.").size(12),
                    text("\n").size(16),
                    text("\n").size(16),
                    text("A special thank you to Geovanna Mendes de Souza, whose observations")
                        .size(12),
                    text("were the starting point for the development of this application.")
                        .size(12),
                ]
                .spacing(15),
            )
            .padding(40)
            .into(),
        };

        let mut main_column = column![nav, content].spacing(25);

        if self.current_view == View::Main {
            let status_footer = container(
                row![
                    text("STATUS:").size(12),
                    text(self.status_message.clone()).size(12),
                ]
                .spacing(10),
            )
            .width(Length::Fill)
            .padding(12)
            .style(|theme: &Theme| {
                let palette = theme.extended_palette();
                container::Style {
                    background: Some(palette.background.strong.color.into()),
                    ..Default::default()
                }
            });

            main_column = main_column.push(status_footer);
        }

        container(main_column).padding(25).into()
    }

    fn subscription(&self) -> iced::Subscription<Message> {
        iced::event::listen_with(|event, _status, id| match event {
            Event::Keyboard(iced::keyboard::Event::KeyPressed { key, .. }) => {
                Some(Message::KeyPressed(key))
            }
            _ => Some(Message::Init(id)),
        })
    }

    fn run(_settings: Settings) -> iced::Result {
        iced::application(Self::title_static, Self::update, Self::view)
            .window(window::Settings {
                size: Size::new(550.0, 750.0),
                resizable: false,
                ..Default::default()
            })
            .subscription(Self::subscription)
            .theme(|_| Theme::Dark)
            .run_with(Self::new)
    }

    fn title_static(_state: &Self) -> String {
        "Screen Capture Manager".to_string()
    }

    fn load_excel(&mut self, path: PathBuf) {
        if let Ok(mut workbook) = open_workbook::<Xlsx<_>, _>(&path) {
            if let Some(Ok(range)) = workbook.worksheet_range_at(0) {
                self.items = range
                    .rows()
                    .skip(1)
                    .filter_map(|r| {
                        let code = r.get(0)?.to_string().trim().to_string();
                        let desc = r.get(1)?.to_string();
                        if code.is_empty() {
                            None
                        } else {
                            Some(TestItem {
                                code,
                                description: desc,
                            })
                        }
                    })
                    .collect();
                self.excel_path = Some(path);
                self.status_message = format!("Loaded {} records.", self.items.len());
            }
        }
    }

    async fn async_capture(excel_path: PathBuf, item_code: String) -> Result<PathBuf, String> {
        let dir = excel_path.parent().unwrap_or(&excel_path).to_path_buf();
        let safe_name = item_code.replace(|c: char| !c.is_alphanumeric(), "_");
        let mut final_path = dir.join(format!("{}.png", safe_name));
        let mut counter = 1;

        while final_path.exists() {
            final_path = dir.join(format!("{}_{}.png", safe_name, counter));
            counter += 1;
        }

        let path_for_thread = final_path.clone();
        tokio::task::spawn_blocking(move || {
            let monitors = Monitor::all().map_err(|e| e.to_string())?;
            let monitor = monitors
                .iter()
                .find(|m| m.x() == 0 && m.y() == 0)
                .unwrap_or(monitors.first().ok_or("No display found.")?);

            let image = monitor.capture_image().map_err(|e| e.to_string())?;
            image.save(&path_for_thread).map_err(|e| e.to_string())?;
            Ok(path_for_thread)
        })
        .await
        .map_err(|e| e.to_string())?
    }
}
