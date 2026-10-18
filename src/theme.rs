//! Tema grafico in stile "ArtCraft / GridCraft" (aspetto Excel-like) e piccole
//! icone vettoriali disegnate a mano, così l'app non dipende da font di icone.

use eframe::egui::{
    self, Align2, Color32, CornerRadius, FontId, Rect, RichText, Stroke, Vec2,
};

// ---------------------------------------------------------------------------
// Palette
// ---------------------------------------------------------------------------

/// Verde "Excel" usato come accento nel tema chiaro.
pub const ACCENT_LIGHT: Color32 = Color32::from_rgb(0x21, 0x73, 0x46);
/// Verde più luminoso, leggibile sul tema scuro.
pub const ACCENT_DARK: Color32 = Color32::from_rgb(0x35, 0xB1, 0x7A);

pub const DANGER: Color32 = Color32::from_rgb(0xC0, 0x39, 0x2B);
pub const WARNING: Color32 = Color32::from_rgb(0xE8, 0x8A, 0x00);

pub fn accent(dark: bool) -> Color32 {
    if dark {
        ACCENT_DARK
    } else {
        ACCENT_LIGHT
    }
}

pub fn panel_bg(dark: bool) -> Color32 {
    if dark {
        Color32::from_rgb(0x1E, 0x1E, 0x1E)
    } else {
        Color32::from_rgb(0xF3, 0xF3, 0xF3)
    }
}

pub fn ribbon_bg(dark: bool) -> Color32 {
    if dark {
        Color32::from_rgb(0x2B, 0x2B, 0x2B)
    } else {
        Color32::WHITE
    }
}

pub fn card_bg(dark: bool) -> Color32 {
    if dark {
        Color32::from_rgb(0x26, 0x26, 0x27)
    } else {
        Color32::from_rgb(0xFB, 0xFB, 0xFB)
    }
}

pub fn border(dark: bool) -> Color32 {
    if dark {
        Color32::from_rgb(0x3C, 0x3C, 0x3C)
    } else {
        Color32::from_rgb(0xD6, 0xD6, 0xD6)
    }
}

pub fn muted(dark: bool) -> Color32 {
    if dark {
        Color32::from_rgb(0xA8, 0xA8, 0xA8)
    } else {
        Color32::from_rgb(0x5F, 0x63, 0x68)
    }
}

/// Colore del testo in primo piano: nero pieno in tema chiaro, bianco in tema scuro.
pub fn text_strong(dark: bool) -> Color32 {
    if dark {
        Color32::from_rgb(0xF2, 0xF2, 0xF2)
    } else {
        Color32::from_rgb(0x00, 0x00, 0x00)
    }
}

pub fn hover_bg(dark: bool) -> Color32 {
    if dark {
        Color32::from_rgb(0x3A, 0x3A, 0x3A)
    } else {
        Color32::from_rgb(0xE4, 0xEF, 0xE8)
    }
}

// ---------------------------------------------------------------------------
// Applicazione dello stile
// ---------------------------------------------------------------------------

pub fn apply(ctx: &egui::Context, dark: bool) {
    let mut visuals = if dark {
        egui::Visuals::dark()
    } else {
        egui::Visuals::light()
    };
    let a = accent(dark);

    visuals.panel_fill = panel_bg(dark);
    visuals.window_fill = card_bg(dark);
    visuals.extreme_bg_color = if dark {
        Color32::from_rgb(0x15, 0x15, 0x15)
    } else {
        Color32::WHITE
    };
    visuals.faint_bg_color = if dark {
        Color32::from_rgb(0x24, 0x24, 0x24)
    } else {
        Color32::from_rgb(0xF0, 0xF0, 0xF0)
    };
    visuals.selection.bg_fill = a;
    visuals.selection.stroke = Stroke::new(1.0_f32, Color32::WHITE);
    visuals.hyperlink_color = a;
    visuals.window_corner_radius = CornerRadius::same(6);
    visuals.menu_corner_radius = CornerRadius::same(6);
    visuals.widgets.noninteractive.bg_stroke = Stroke::new(1.0_f32, border(dark));
    visuals.widgets.noninteractive.fg_stroke = Stroke::new(1.0_f32, visuals.text_color());
    for w in [
        &mut visuals.widgets.inactive,
        &mut visuals.widgets.hovered,
        &mut visuals.widgets.active,
        &mut visuals.widgets.open,
    ] {
        w.corner_radius = CornerRadius::same(4);
    }
    visuals.widgets.inactive.weak_bg_fill = if dark {
        Color32::from_rgb(0x33, 0x33, 0x33)
    } else {
        Color32::from_rgb(0xEC, 0xEC, 0xEC)
    };
    visuals.widgets.hovered.weak_bg_fill = hover_bg(dark);
    visuals.widgets.hovered.bg_stroke = Stroke::new(1.0_f32, a.gamma_multiply(0.75));
    visuals.widgets.active.bg_fill = a.gamma_multiply(0.9);
    visuals.widgets.active.fg_stroke = Stroke::new(1.0_f32, Color32::WHITE);
    visuals.widgets.active.bg_stroke = Stroke::new(1.0_f32, a);

    let mut style = (*ctx.style()).clone();
    style.visuals = visuals;
    style.spacing.item_spacing = Vec2::new(6.0, 4.0);
    style.spacing.button_padding = Vec2::new(6.0, 3.0);
    style.spacing.interact_size.y = 20.0;
    style.spacing.menu_margin = egui::Margin::same(4);
    // Barre di scorrimento "solide" (non flottanti): riservano spazio al loro
    // interno, così non coprono l'ultima riga di tabelle ed elenchi.
    // Con lo stile predefinito di egui le barre flottano sopra il contenuto.
    let mut scroll = egui::style::ScrollStyle::solid();
    scroll.bar_width = 8.0;
    scroll.bar_inner_margin = 4.0;
    scroll.bar_outer_margin = 2.0;
    style.spacing.scroll = scroll;
    // Testi più compatti: riducono l'ingombro complessivo dell'interfaccia.
    style.text_styles = [
        (egui::TextStyle::Heading, FontId::proportional(16.0)),
        (egui::TextStyle::Body, FontId::proportional(12.5)),
        (egui::TextStyle::Monospace, FontId::monospace(11.5)),
        (egui::TextStyle::Button, FontId::proportional(12.5)),
        (egui::TextStyle::Small, FontId::proportional(10.5)),
    ]
    .into();
    ctx.set_style(style);
}

// ---------------------------------------------------------------------------
// Icone vettoriali
// ---------------------------------------------------------------------------

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum Icon {
    Save,
    Folder,
    Sun,
    Moon,
    Info,
    Capture,
}

/// Converte coordinate normalizzate (0..1) del rettangolo dell'icona in punti schermo.
fn n(rect: Rect, x: f32, y: f32) -> egui::Pos2 {
    egui::pos2(rect.left() + rect.width() * x, rect.top() + rect.height() * y)
}

/// Disegna un'icona vettoriale dentro `rect` (nessuna dipendenza da font).
pub fn draw_icon(painter: &egui::Painter, icon: Icon, rect: Rect, color: Color32, w: f32) {
    let s = Stroke::new(w, color);
    let line = |a: (f32, f32), b: (f32, f32)| {
        painter.line_segment([n(rect, a.0, a.1), n(rect, b.0, b.1)], s);
    };

    match icon {
        Icon::Save => {
            // Corpo "floppy"
            painter.rect_stroke(
                Rect::from_min_max(n(rect, 0.16, 0.16), n(rect, 0.84, 0.84)),
                CornerRadius::same(1),
                s,
                egui::StrokeKind::Middle,
            );
            // Etichetta in alto
            painter.rect_filled(
                Rect::from_min_max(n(rect, 0.32, 0.16), n(rect, 0.68, 0.40)),
                CornerRadius::ZERO,
                color,
            );
            // Sportello in basso
            painter.rect_stroke(
                Rect::from_min_max(n(rect, 0.30, 0.55), n(rect, 0.70, 0.84)),
                CornerRadius::ZERO,
                s,
                egui::StrokeKind::Middle,
            );
        }
        Icon::Folder => {
            painter.add(egui::Shape::closed_line(
                vec![
                    n(rect, 0.12, 0.78),
                    n(rect, 0.12, 0.22),
                    n(rect, 0.40, 0.22),
                    n(rect, 0.47, 0.34),
                    n(rect, 0.88, 0.34),
                    n(rect, 0.88, 0.78),
                ],
                s,
            ));
        }
        Icon::Sun => {
            painter.circle_filled(n(rect, 0.5, 0.5), rect.width() * 0.20, color);
            for i in 0..8 {
                let a = std::f32::consts::TAU * (i as f32) / 8.0;
                let (dx, dy) = (a.cos(), a.sin());
                line(
                    (0.5 + dx * 0.30, 0.5 + dy * 0.30),
                    (0.5 + dx * 0.44, 0.5 + dy * 0.44),
                );
            }
        }
        Icon::Moon => {
            // Falce: arco del cerchio esterno + arco del cerchio interno (spostato).
            let (cx, cy, r) = (0.5f32, 0.5f32, 0.36f32);
            let (bx, by, br) = (0.62f32, 0.5f32, 0.30f32);
            let d = ((bx - cx).powi(2) + (by - cy).powi(2)).sqrt();
            let a = (d * d + r * r - br * br) / (2.0 * d);
            let h = (r * r - a * a).max(0.0).sqrt();
            let ix = cx + a * (bx - cx) / d;
            let iy = cy + a * (by - cy) / d;
            let p1 = (ix, iy - h);
            let p2 = (ix, iy + h);
            let ang = |px: f32, py: f32, ox: f32, oy: f32| (py - oy).atan2(px - ox);
            let a1 = ang(p1.0, p1.1, cx, cy);
            let a2 = ang(p2.0, p2.1, cx, cy);
            let b1 = ang(p2.0, p2.1, bx, by);
            let b2 = ang(p1.0, p1.1, bx, by);

            let mut pts: Vec<egui::Pos2> = Vec::new();
            // arco esterno passando dal lato sinistro (angoli decrescenti da a1)
            let mut t = a1;
            while t > a2 - std::f32::consts::TAU {
                pts.push(n(rect, cx + r * t.cos(), cy + r * t.sin()));
                t -= 0.15;
            }
            pts.push(n(rect, p2.0, p2.1));
            // arco interno passando dal lato sinistro del cerchio interno
            let mut t = b1;
            let mut guard = 0;
            while (t - b2).abs() > 0.05 && guard < 200 {
                pts.push(n(rect, bx + br * t.cos(), by + br * t.sin()));
                t += 0.15;
                guard += 1;
            }
            painter.add(egui::Shape::convex_polygon(pts, color, Stroke::NONE));
        }
        Icon::Info => {
            painter.circle_stroke(n(rect, 0.5, 0.5), rect.width() * 0.36, s);
            line((0.5, 0.44), (0.5, 0.72));
            painter.circle_filled(n(rect, 0.5, 0.30), w * 0.7, color);
        }
        Icon::Capture => {
            // Cornice di cattura: quattro angoli aperti con il diaframma al centro.
            for (orizzontale, verticale) in [
                ((0.12, 0.14), (0.36, 0.14)),
                ((0.12, 0.14), (0.12, 0.38)),
                ((0.64, 0.14), (0.88, 0.14)),
                ((0.88, 0.14), (0.88, 0.38)),
                ((0.64, 0.86), (0.88, 0.86)),
                ((0.88, 0.62), (0.88, 0.86)),
                ((0.12, 0.86), (0.36, 0.86)),
                ((0.12, 0.62), (0.12, 0.86)),
            ] {
                line(orizzontale, verticale);
            }
            painter.circle_filled(n(rect, 0.5, 0.5), rect.width() * 0.15, color);
        }
    }
}

// ---------------------------------------------------------------------------
// Widget riutilizzabili
// ---------------------------------------------------------------------------

/// Pulsante "grande" della barra multifunzione: icona sopra, etichetta sotto.
pub fn big_button(ui: &mut egui::Ui, icon: Icon, label: &str) -> egui::Response {
    big_button_state(ui, icon, label, false)
}

/// Come [`big_button`], ma con stato "attivo" (toggle) evidenziato.
pub fn big_button_state(
    ui: &mut egui::Ui,
    icon: Icon,
    label: &str,
    active: bool,
) -> egui::Response {
    let dark = ui.visuals().dark_mode;
    let a = accent(dark);
    // La larghezza segue l'etichetta: le traduzioni piu' lunghe ("Executar
    // captura") non devono uscire dal pulsante.
    let testo_w = ui.fonts(|f| {
        f.layout_no_wrap(
            label.to_owned(),
            FontId::proportional(10.0),
            Color32::WHITE,
        )
        .size()
        .x
    });
    let size = Vec2::new((testo_w + 18.0).max(48.0), 38.0);
    let (rect, resp) = ui.allocate_exact_size(size, egui::Sense::click());
    let painter = ui.painter();

    // Lo stato attivo ("Punti" e "Legenda" della barra comandi) è indicato
    // **solo** dalla cornice, come ogni altra selezione del programma; il
    // riempimento è riservato ai riscontri transitorii del puntatore.
    let bg = if resp.is_pointer_button_down_on() {
        a.gamma_multiply(0.30)
    } else if resp.hovered() {
        hover_bg(dark)
    } else {
        Color32::TRANSPARENT
    };
    if bg != Color32::TRANSPARENT {
        painter.rect_filled(rect, CornerRadius::same(5), bg);
    }
    if active {
        selection_frame(painter, rect, CornerRadius::same(5), dark);
    } else if resp.hovered() {
        painter.rect_stroke(
            rect,
            CornerRadius::same(5),
            Stroke::new(1.0_f32, a.gamma_multiply(0.55)),
            egui::StrokeKind::Inside,
        );
    }

    let icon_rect = Rect::from_center_size(
        egui::pos2(rect.center().x, rect.top() + 13.0),
        Vec2::splat(17.0),
    );
    let fg = if resp.hovered() || active {
        a
    } else {
        ui.visuals().text_color()
    };
    draw_icon(painter, icon, icon_rect, fg, 1.5);
    painter.text(
        egui::pos2(rect.center().x, rect.bottom() - 9.0),
        Align2::CENTER_CENTER,
        label,
        FontId::proportional(10.0),
        ui.visuals().text_color(),
    );

    if resp.hovered() {
        ui.ctx().set_cursor_icon(egui::CursorIcon::PointingHand);
    }
    resp
}

/// Pulsante compatto a "pillola": icona opzionale + testo, con stato attivo.
pub fn pill_button(
    ui: &mut egui::Ui,
    icon: Option<Icon>,
    label: &str,
    active: bool,
) -> egui::Response {
    let dark = ui.visuals().dark_mode;
    let a = accent(dark);
    let text_w = ui
        .fonts(|f| {
            f.layout_no_wrap(
                label.to_owned(),
                FontId::proportional(11.0),
                Color32::WHITE,
            )
            .size()
            .x
        });
    let width = text_w + 16.0 + if icon.is_some() { 15.0 } else { 0.0 };
    let (rect, resp) = ui.allocate_exact_size(Vec2::new(width, 20.0), egui::Sense::click());
    let painter = ui.painter();

    // La selezione (linguette di Configurazioni, pulsanti attivi del grafico)
    // è **solo** una cornice: nessun riempimento, altrimenti il testo andrebbe
    // forzato in bianco sul tema chiaro e smarrebbe il proprio colore.
    let bg = if resp.is_pointer_button_down_on() {
        a.gamma_multiply(0.35)
    } else if resp.hovered() {
        hover_bg(dark)
    } else {
        Color32::TRANSPARENT
    };
    if bg != Color32::TRANSPARENT {
        painter.rect_filled(rect, CornerRadius::same(4), bg);
    }
    painter.rect_stroke(
        rect,
        CornerRadius::same(4),
        Stroke::new(
            1.4_f32,
            if active {
                a
            } else {
                border(dark)
            },
        ),
        egui::StrokeKind::Inside,
    );

    let fg = if active {
        a
    } else if resp.hovered() {
        a
    } else {
        ui.visuals().text_color()
    };
    let mut text_left = rect.left() + 8.0;
    if let Some(ic) = icon {
        let icon_rect = Rect::from_center_size(
            egui::pos2(rect.left() + 12.0, rect.center().y),
            Vec2::splat(12.0),
        );
        draw_icon(painter, ic, icon_rect, fg, 1.3);
        text_left += 14.0;
    }
    painter.text(
        egui::pos2(text_left, rect.center().y),
        Align2::LEFT_CENTER,
        label,
        FontId::proportional(11.0),
        fg,
    );
    if resp.hovered() {
        ui.ctx().set_cursor_icon(egui::CursorIcon::PointingHand);
    }
    resp
}

/// "Card" bianca con bordo, usata per raggruppare i contenuti di una pagina.
pub fn card<R>(ui: &mut egui::Ui, add: impl FnOnce(&mut egui::Ui) -> R) -> R {
    let dark = ui.visuals().dark_mode;
    egui::Frame::new()
        .fill(card_bg(dark))
        .stroke(Stroke::new(1.0_f32, border(dark)))
        .corner_radius(CornerRadius::same(6))
        .inner_margin(egui::Margin::same(9))
        .show(ui, add)
        .inner
}

/// Titolo di sezione con barretta di accento a sinistra.
pub fn section_title(ui: &mut egui::Ui, title: &str) {
    let dark = ui.visuals().dark_mode;
    let a = accent(dark);
    ui.horizontal(|ui| {
        let (rect, _) = ui.allocate_exact_size(Vec2::new(3.0, 15.0), egui::Sense::hover());
        ui.painter().rect_filled(rect, CornerRadius::same(2), a);
        ui.add_space(2.0);
        ui.label(
            RichText::new(title)
                .size(13.0)
                .strong()
                // Nero pieno nel tema chiaro: il colore di testo di default di
                // egui in tema chiaro è un grigio (80), non nero.
                .color(text_strong(dark)),
        );
    });
}

/// Etichetta secondaria (grigio, piccola).
pub fn hint(ui: &mut egui::Ui, text: &str) {
    let dark = ui.visuals().dark_mode;
    ui.label(RichText::new(text).size(10.5).color(muted(dark)));
}

/// Cornice di selezione: è l'**unico** modo in cui il programma indica una
/// selezione (riga dell'elenco progetti, voce della lista sensori, linguetta
/// di scheda attiva, pulsante attivo). La selezione è sempre e solo un
/// contorno, mai un riempimento: così il testo sottostante mantiene il proprio
/// colore e resta leggibile, e tutte le schede si riconoscono allo stesso modo.
pub fn selection_frame(
    painter: &egui::Painter,
    rect: egui::Rect,
    corner: egui::CornerRadius,
    dark: bool,
) {
    painter.rect_stroke(
        rect.expand(2.0),
        corner,
        Stroke::new(1.4_f32, accent(dark)),
        egui::StrokeKind::Inside,
    );
}

/// Piccolo pulsante quadrato con un'icona vettoriale, pensato per le righe di
/// lista dove serve un'azione di scorrimento (dimensioni e riscontri di
/// puntatore simili a `Ui::small_button`).
pub fn small_icon_button(
    ui: &mut egui::Ui,
    icon: Icon,
    color: Color32,
    size: f32,
) -> egui::Response {
    let dark = ui.visuals().dark_mode;
    let (rect, resp) = ui.allocate_exact_size(Vec2::splat(size), egui::Sense::click());
    if resp.hovered() || resp.is_pointer_button_down_on() {
        ui.painter()
            .rect_filled(rect, CornerRadius::same(3), hover_bg(dark));
    }
    draw_icon(ui.painter(), icon, rect, color, 1.4);
    if resp.hovered() {
        ui.ctx().set_cursor_icon(egui::CursorIcon::PointingHand);
    }
    resp
}

/// Pulsante d'azione con icona + testo, a tinta piena (il comando principale
/// della scheda in cui si trova).
pub fn action_button(
    ui: &mut egui::Ui,
    icon: Icon,
    label: &str,
) -> egui::Response {
    let dark = ui.visuals().dark_mode;
    let a = accent(dark);
    let text_w = ui.fonts(|f| {
        f.layout_no_wrap(label.to_owned(), FontId::proportional(11.5), Color32::WHITE)
            .size()
            .x
    });
    let size = Vec2::new(text_w + 34.0, 23.0);
    let (rect, resp) = ui.allocate_exact_size(size, egui::Sense::click());
    let painter = ui.painter();

    let (bg, fg, border_color) = (a, Color32::WHITE, a);

    let bg = if resp.is_pointer_button_down_on() {
        bg.gamma_multiply(0.85)
    } else if resp.hovered() {
        bg.gamma_multiply(1.15)
    } else {
        bg
    };

    painter.rect_filled(rect, CornerRadius::same(4), bg);
    painter.rect_stroke(
        rect,
        CornerRadius::same(4),
        Stroke::new(1.0_f32, border_color),
        egui::StrokeKind::Inside,
    );

    let icon_rect = Rect::from_center_size(
        egui::pos2(rect.left() + 12.0, rect.center().y),
        Vec2::splat(12.0),
    );
    draw_icon(painter, icon, icon_rect, fg, 1.4);
    painter.text(
        egui::pos2(rect.left() + 22.0, rect.center().y),
        Align2::LEFT_CENTER,
        label,
        FontId::proportional(11.5),
        fg,
    );

    if resp.hovered() {
        ui.ctx().set_cursor_icon(egui::CursorIcon::PointingHand);
    }
    resp
}
