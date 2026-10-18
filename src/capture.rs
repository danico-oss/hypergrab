//! Cattura della finestra in primo piano e composizione della data e dell'ora.
//!
//! La cattura usa le API di Windows: si prende la finestra che in quel momento
//! sta in primo piano (Hypergrab si e' appena nascosto) e se ne copiano i pixel
//! dallo schermo. Se la finestra non e' utilizzabile (minimizzata, chiusa fra
//! il clic e la cattura) si ripiega sull'intero schermo virtuale.

use chrono::{DateTime, Local};
use image::{Rgb, RgbImage};
use std::io::Write;
use std::path::{Path, PathBuf};

use crate::i18n::Language;

/// Risultato di una cattura.
pub struct Cattura {
    /// Pixel catturati.
    pub immagine: RgbImage,
    /// Titolo della finestra catturata (vuoto se si e' ripiegati sullo schermo).
    pub finestra: String,
}

// ---------------------------------------------------------------------------
// Data e ora
// ---------------------------------------------------------------------------

/// Formato della data e dell'ora per la lingua attiva.
pub fn formato_data_ora(lang: Language) -> &'static str {
    match lang {
        Language::En => "%m/%d/%Y %H:%M:%S",
        Language::It | Language::PtBr => "%d/%m/%Y %H:%M:%S",
    }
}

/// Data e ora come verranno impresse sull'immagine.
pub fn testo_data_ora(lang: Language, quando: DateTime<Local>) -> String {
    quando.format(formato_data_ora(lang)).to_string()
}

// ---------------------------------------------------------------------------
// Cattura
// ---------------------------------------------------------------------------

#[cfg(not(windows))]
pub fn cattura_finestra_primo_piano() -> Result<Cattura, String> {
    Err("la cattura dello schermo e' disponibile solo su Windows".to_string())
}

#[cfg(windows)]
pub fn cattura_finestra_primo_piano() -> Result<Cattura, String> {
    use windows_sys::Win32::UI::WindowsAndMessaging::{
        GetForegroundWindow, GetSystemMetrics, GetWindowRect, SM_CXVIRTUALSCREEN,
        SM_CYVIRTUALSCREEN, SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN,
    };

    unsafe {
        // Finestra in primo piano: e' l'applicazione che rimane visibile.
        let finestra = GetForegroundWindow();
        if utilizzabile(finestra) {
            let mut rect = windows_sys::Win32::Foundation::RECT {
                left: 0,
                top: 0,
                right: 0,
                bottom: 0,
            };
            if GetWindowRect(finestra, &mut rect) != 0 {
                let (larghezza, altezza) = (rect.right - rect.left, rect.bottom - rect.top);
                if larghezza > 0 && altezza > 0 {
                    if let Some(immagine) =
                        copia_schermo(rect.left, rect.top, larghezza, altezza)
                    {
                        return Ok(Cattura {
                            immagine,
                            finestra: titolo_finestra(finestra),
                        });
                    }
                }
            }
        }

        // Ripiego: intero schermo virtuale (tutti i monitor insieme).
        let (x, y) = (
            GetSystemMetrics(SM_XVIRTUALSCREEN),
            GetSystemMetrics(SM_YVIRTUALSCREEN),
        );
        let (larghezza, altezza) = (
            GetSystemMetrics(SM_CXVIRTUALSCREEN),
            GetSystemMetrics(SM_CYVIRTUALSCREEN),
        );
        if larghezza <= 0 || altezza <= 0 {
            return Err("schermo non rilevato".to_string());
        }
        copia_schermo(x, y, larghezza, altezza)
            .map(|immagine| Cattura {
                immagine,
                finestra: String::new(),
            })
            .ok_or_else(|| "cattura non riuscita".to_string())
    }
}

/// Vero se la finestra si può fotografare: esiste, è visibile, non è ridotta a
/// icona e non è una finestra di questa applicazione.
///
/// Hypergrab si nasconde prima di catturare, ma Windows può impiegare un attimo
/// a dare il primo piano all'applicazione sottostante: finché non è così la
/// cattura riguarderebbe Hypergrab stessa.
#[cfg(windows)]
unsafe fn utilizzabile(finestra: windows_sys::Win32::Foundation::HWND) -> bool {
    use windows_sys::Win32::System::Threading::GetCurrentProcessId;
    use windows_sys::Win32::UI::WindowsAndMessaging::{
        GetWindowThreadProcessId, IsIconic, IsWindowVisible,
    };

    if finestra.is_null() {
        return false;
    }
    let mut processo = 0u32;
    GetWindowThreadProcessId(finestra, &mut processo);
    if processo == GetCurrentProcessId() {
        return false;
    }
    IsWindowVisible(finestra) != 0 && IsIconic(finestra) == 0
}

/// Vero se in primo piano c'è già l'applicazione da documentare.
#[cfg(windows)]
pub fn primo_piano_pronto() -> bool {
    unsafe { utilizzabile(windows_sys::Win32::UI::WindowsAndMessaging::GetForegroundWindow()) }
}

#[cfg(not(windows))]
pub fn primo_piano_pronto() -> bool {
    false
}

/// Copia dallo schermo il rettangolo indicato.
#[cfg(windows)]
unsafe fn copia_schermo(x: i32, y: i32, larghezza: i32, altezza: i32) -> Option<RgbImage> {
    use windows_sys::Win32::Graphics::Gdi::{
        BitBlt, CreateCompatibleBitmap, CreateCompatibleDC, DeleteDC, DeleteObject, GetDC, GetDIBits,
        ReleaseDC, SelectObject, BITMAPINFO, BITMAPINFOHEADER, BI_RGB, DIB_RGB_COLORS, HGDIOBJ,
        SRCCOPY,
    };

    let schermo = GetDC(std::ptr::null_mut());
    if schermo.is_null() {
        return None;
    }
    let memoria = CreateCompatibleDC(schermo);
    let bitmap = CreateCompatibleBitmap(schermo, larghezza, altezza);
    let mut risultato = None;

    if !memoria.is_null() && !bitmap.is_null() {
        let vecchio = SelectObject(memoria, bitmap as HGDIOBJ);
        if BitBlt(
            memoria, 0, 0, larghezza, altezza, schermo, x, y, SRCCOPY,
        ) != 0
        {
            let mut info: BITMAPINFO = std::mem::zeroed();
            info.bmiHeader.biSize = std::mem::size_of::<BITMAPINFOHEADER>() as u32;
            info.bmiHeader.biWidth = larghezza;
            // Altezza negativa: le righe arrivano dall'alto verso il basso.
            info.bmiHeader.biHeight = -altezza;
            info.bmiHeader.biPlanes = 1;
            info.bmiHeader.biBitCount = 32;
            info.bmiHeader.biCompression = BI_RGB;

            let mut pixel = vec![0u8; (larghezza * altezza * 4) as usize];
            let righe = GetDIBits(
                memoria,
                bitmap,
                0,
                altezza as u32,
                pixel.as_mut_ptr() as *mut core::ffi::c_void,
                &mut info,
                DIB_RGB_COLORS,
            );
            if righe != 0 {
                // Da BGRA a RGB: il canale alfa dei bitmap dello schermo non
                // e' significativo.
                let mut rgb = Vec::with_capacity(pixel.len() / 4 * 3);
                for c in pixel.chunks_exact(4) {
                    rgb.extend_from_slice(&[c[2], c[1], c[0]]);
                }
                risultato = RgbImage::from_raw(larghezza as u32, altezza as u32, rgb);
            }
        }
        SelectObject(memoria, vecchio);
        DeleteObject(bitmap as HGDIOBJ);
    }
    if !memoria.is_null() {
        DeleteDC(memoria);
    }
    ReleaseDC(std::ptr::null_mut(), schermo);
    risultato
}

/// Titolo della finestra, per i messaggi di stato.
#[cfg(windows)]
unsafe fn titolo_finestra(finestra: windows_sys::Win32::Foundation::HWND) -> String {
    use windows_sys::Win32::UI::WindowsAndMessaging::GetWindowTextW;
    let mut buffer = [0u16; 512];
    let n = GetWindowTextW(finestra, buffer.as_mut_ptr(), buffer.len() as i32);
    if n <= 0 {
        return String::new();
    }
    String::from_utf16_lossy(&buffer[..n as usize])
}

// ---------------------------------------------------------------------------
// Composizione della data e dell'ora sull'immagine
// ---------------------------------------------------------------------------

/// Arancione brillante con cui vengono impresse data e ora sull'immagine.
const ARANCIONE_DATA_ORA: Rgb<u8> = Rgb([255, 165, 0]);

/// Colore del contorno attorno alla scritta.
const CONTORNO_DATA_ORA: Rgb<u8> = Rgb([0, 0, 0]);

/// Disegna data e ora in basso a destra, in arancione brillante e senza sfondo.
///
/// Il carattere viene preso dai font di sistema. Attorno alla scritta viene
/// disegnato un contorno scuro pieno, che la stacca sia dai fondi chiari sia da
/// quelli scuri senza bisogno di una targhetta colorata.
pub fn compone_data_ora(immagine: &mut RgbImage, testo: &str) -> Result<(), String> {
    let font = carica_font()?;
    let (larghezza, altezza) = immagine.dimensions();
    let lato = larghezza.min(altezza) as f32;
    let scala = ab_glyph::PxScale::from((lato * 0.05).clamp(15.0, 72.0));
    let (testo_w, testo_h) = imageproc::drawing::text_size(scala, &font, testo);

    let margine = (lato * 0.02).clamp(8.0, 30.0) as i32;
    let x = larghezza as i32 - margine - testo_w as i32;
    let y = altezza as i32 - margine - testo_h as i32;

    // Contorno: il testo scuro viene ripetuto a ogni pixel di scostamento
    // intorno alla posizione, così attorno alle lettere resta un alone pieno.
    let raggio = ((scala.y * 0.06).round() as i32).clamp(1, 3);
    for dy in -raggio..=raggio {
        for dx in -raggio..=raggio {
            if dx == 0 && dy == 0 {
                continue;
            }
            imageproc::drawing::draw_text_mut(
                immagine,
                CONTORNO_DATA_ORA,
                x + dx,
                y + dy,
                scala,
                &font,
                testo,
            );
        }
    }

    // Arancione brillante, disegnato due volte: i pixel di bordo non restano
    // smorzati dall'antialiasing e la scritta riesce piena.
    for _ in 0..2 {
        imageproc::drawing::draw_text_mut(immagine, ARANCIONE_DATA_ORA, x, y, scala, &font, testo);
    }
    Ok(())
}

/// Cerca un carattere di sistema utilizzabile per il testo sull'immagine.
fn carica_font() -> Result<ab_glyph::FontVec, String> {
    // Segoe UI c'e' su ogni Windows moderno; gli altri sono ripieghi.
    const CANDIDATI: [&str; 6] = [
        "segoeui.ttf",
        "arial.ttf",
        "tahoma.ttf",
        "verdana.ttf",
        "calibri.ttf",
        "DejaVuSans.ttf",
    ];
    let mut percorsi: Vec<PathBuf> = Vec::new();
    for radice in ["C:\\Windows\\Fonts", "C:\\WINNT\\Fonts"] {
        for nome in CANDIDATI {
            percorsi.push(Path::new(radice).join(nome));
        }
    }
    if let Ok(cartella) = std::env::var("LOCALAPPDATA") {
        for nome in CANDIDATI {
            percorsi.push(
                Path::new(&cartella)
                    .join("Microsoft")
                    .join("Windows")
                    .join("Fonts")
                    .join(nome),
            );
        }
    }
    for percorso in percorsi {
        if let Ok(dati) = std::fs::read(&percorso) {
            if let Ok(font) = ab_glyph::FontVec::try_from_vec(dati) {
                return Ok(font);
            }
        }
    }
    Err("nessun carattere di sistema trovato per scrivere data e ora".to_string())
}

// ---------------------------------------------------------------------------
// Salvataggio
// ---------------------------------------------------------------------------

/// Nome del file JPG per un codice di test.
pub fn nome_file(codice: &str) -> String {
    format!("{}.jpg", sanitizza(codice))
}

/// Percorso del file JPG: stessa cartella del file di elenco, nome = codice.
pub fn percorso_salvataggio(cartella: &Path, codice: &str) -> PathBuf {
    cartella.join(nome_file(codice))
}

/// Rende un codice utilizzabile come nome di file su Windows.
pub fn sanitizza(codice: &str) -> String {
    const VIETATI: &str = "\\/:*?\"<>|";
    let mut nome: String = codice
        .chars()
        .map(|c| {
            if VIETATI.contains(c) || (c as u32) < 32 {
                '_'
            } else {
                c
            }
        })
        .collect();
    // Windows non ammette spazi e punti finali.
    nome = nome.trim().trim_end_matches(|c| c == '.' || c == ' ').to_string();
    if nome.is_empty() {
        return "test".to_string();
    }
    // Nomi riservati dal sistema (CON, PRN, ...): basta un suffisso.
    const RISERVATI: [&str; 22] = [
        "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7",
        "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    ];
    if RISERVATI.contains(&nome.to_uppercase().as_str()) {
        nome.push('_');
    }
    nome
}

/// Scrive l'immagine in JPEG (qualita' 90): il file esistente viene
/// sovrascritto.
pub fn salva_jpg(immagine: &RgbImage, percorso: &Path) -> Result<(), String> {
    let file = std::fs::File::create(percorso).map_err(|e| e.to_string())?;
    let mut scrittura = std::io::BufWriter::new(file);
    let mut codificatore =
        image::codecs::jpeg::JpegEncoder::new_with_quality(&mut scrittura, 90);
    codificatore
        .encode(
            immagine.as_raw(),
            immagine.width(),
            immagine.height(),
            image::ExtendedColorType::Rgb8,
        )
        .map_err(|e| e.to_string())?;
    scrittura.flush().map_err(|e| e.to_string())?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::TimeZone;

    fn quando_fisso() -> DateTime<Local> {
        Local.with_ymd_and_hms(2026, 10, 18, 14, 3, 22).unwrap()
    }

    /// Una finestra che non esiste non si può fotografare.
    #[cfg(windows)]
    #[test]
    fn finestra_inesistente_non_e_utilizzabile() {
        assert!(!unsafe { utilizzabile(std::ptr::null_mut()) });
    }

    #[test]
    fn data_e_ora_secondo_la_lingua() {
        assert_eq!(
            testo_data_ora(Language::It, quando_fisso()),
            "18/10/2026 14:03:22"
        );
        assert_eq!(
            testo_data_ora(Language::PtBr, quando_fisso()),
            "18/10/2026 14:03:22"
        );
        assert_eq!(
            testo_data_ora(Language::En, quando_fisso()),
            "10/18/2026 14:03:22"
        );
    }

    #[test]
    fn sanitizza_i_codici() {
        assert_eq!(sanitizza("T-01"), "T-01");
        assert_eq!(sanitizza("  T-01  "), "T-01");
        assert_eq!(sanitizza("a/b\\c"), "a_b_c");
        assert_eq!(sanitizza("a:b*c?d\"e<f>g|h"), "a_b_c_d_e_f_g_h");
        assert_eq!(sanitizza("prova."), "prova");
        // Codice vuoto o fatto di soli caratteri vietati: nome di ripiego.
        assert_eq!(sanitizza(""), "test");
        assert_eq!(sanitizza("///"), "___");
        assert_eq!(sanitizza("   "), "test");
        // Nomi riservati dal sistema.
        assert_eq!(sanitizza("CON"), "CON_");
        assert_eq!(sanitizza("com1"), "com1_");
    }

    #[test]
    fn nome_e_percorso_del_file() {
        assert_eq!(nome_file("T-01"), "T-01.jpg");
        assert_eq!(
            percorso_salvataggio(Path::new("C:\\prove"), "T-01"),
            PathBuf::from("C:\\prove\\T-01.jpg")
        );
        assert_eq!(
            percorso_salvataggio(Path::new("C:\\prove"), "a/b"),
            PathBuf::from("C:\\prove\\a_b.jpg")
        );
    }

    #[test]
    fn salva_e_rilegge_un_jpg() {
        let mut immagine = RgbImage::new(64, 40);
        for (x, _y, p) in immagine.enumerate_pixels_mut() {
            *p = Rgb([(x * 4) as u8, 40, 200]);
        }
        let path = std::env::temp_dir().join(format!(
            "hypergrab_jpg_{}_{}.jpg",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        salva_jpg(&immagine, &path).unwrap();

        let letto = image::open(&path).unwrap();
        assert_eq!(letto.width(), 64);
        assert_eq!(letto.height(), 40);
        // Il file deve essere un JPEG vero: la firma inizia con FF D8.
        let bytes = std::fs::read(&path).unwrap();
        assert_eq!(&bytes[0..2], &[0xFF, 0xD8]);
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn compone_data_e_ora_in_basso_a_destra() {
        // Il test usa il carattere di sistema: se non c'e' (sistema non
        // Windows), viene saltato.
        let Ok(_) = carica_font() else {
            eprintln!("carattere di sistema assente: verifica saltata");
            return;
        };

        let mut immagine = RgbImage::from_pixel(800, 600, Rgb([255u8, 255u8, 255u8]));
        let prima = immagine.clone();
        compone_data_ora(&mut immagine, "18/10/2026 14:03:22").unwrap();

        // In alto a sinistra non cambia nulla: la scritta sta in basso a
        // destra.
        for y in 0..300 {
            for x in 0..400 {
                assert_eq!(immagine.get_pixel(x, y), prima.get_pixel(x, y));
            }
        }

        let (mut arancioni, mut contorno) = (0, 0);
        let (mut min_x, mut max_x, mut min_y, mut max_y) = (u32::MAX, 0, u32::MAX, 0);
        for y in 300..600 {
            for x in 400..800 {
                let nuovo = immagine.get_pixel(x, y);
                if nuovo == prima.get_pixel(x, y) {
                    continue;
                }
                (min_x, max_x) = (min_x.min(x), max_x.max(x));
                (min_y, max_y) = (min_y.min(y), max_y.max(y));
                // Arancione brillante...
                if nuovo.0[0] > 200 && nuovo.0[1] > 80 && nuovo.0[1] < 230 && nuovo.0[2] < 100 {
                    arancioni += 1;
                }
                // ...con il contorno scuro che lo rende leggibile.
                if nuovo.0.iter().all(|&canale| canale < 80) {
                    contorno += 1;
                }
            }
        }
        assert!(
            arancioni > 200,
            "data e ora non in arancione brillante: {arancioni} pixel"
        );
        assert!(contorno > 100, "manca il contorno scuro: {contorno} pixel");

        // Attorno alla scritta non viene disegnato nulla: né un velo né una
        // targhetta. I pixel a ridosso del rettangolo della scritta devono
        // essere rimasti intatti.
        let margine = 4;
        for y in min_y.saturating_sub(margine)..=(max_y + margine).min(599) {
            for x in min_x.saturating_sub(margine)..=(max_x + margine).min(799) {
                if x >= min_x && x <= max_x && y >= min_y && y <= max_y {
                    continue;
                }
                assert_eq!(
                    immagine.get_pixel(x, y),
                    prima.get_pixel(x, y),
                    "c'e' qualcosa attorno alla scritta in ({x}, {y})"
                );
            }
        }
    }

    /// Verifica la catena di copia dallo schermo (GetDC/BitBlt/GetDIBits)
    /// senza dipendere dalla finestra in primo piano: copia un angolo di
    /// schermo. In un ambiente senza schermo il test viene saltato.
    #[test]
    #[cfg(windows)]
    fn copia_una_porzione_di_schermo() {
        let Some(immagine) = (unsafe { copia_schermo(0, 0, 64, 64) }) else {
            eprintln!("schermo non disponibile: verifica saltata");
            return;
        };
        assert_eq!(immagine.dimensions(), (64, 64));
    }
}
