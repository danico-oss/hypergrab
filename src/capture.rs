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
        GetForegroundWindow, GetSystemMetrics, GetWindowRect, IsIconic, SM_CXVIRTUALSCREEN,
        SM_CYVIRTUALSCREEN, SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN,
    };

    unsafe {
        // Finestra in primo piano: e' l'applicazione che rimane visibile.
        let finestra = GetForegroundWindow();
        if !finestra.is_null() && IsIconic(finestra) == 0 {
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

/// Disegna data e ora in basso a destra, su un velo scuro per restare leggibili
/// su qualunque sfondo. Il carattere viene preso dai font di sistema.
pub fn compone_data_ora(immagine: &mut RgbImage, testo: &str) -> Result<(), String> {
    let font = carica_font()?;
    let (larghezza, altezza) = immagine.dimensions();
    let lato = larghezza.min(altezza) as f32;
    let scala = ab_glyph::PxScale::from((lato * 0.045).clamp(13.0, 64.0));
    let (testo_w, testo_h) = imageproc::drawing::text_size(scala, &font, testo);

    let margine = (lato * 0.02).clamp(6.0, 30.0) as i32;
    let riempimento = (scala.y * 0.35).max(4.0) as i32;
    let x1 = larghezza as i32 - margine;
    let y1 = altezza as i32 - margine;
    let x0 = x1 - testo_w as i32 - 2 * riempimento;
    let y0 = y1 - testo_h as i32 - 2 * riempimento;

    velo(immagine, x0, y0, x1, y1, 0.55);
    imageproc::drawing::draw_text_mut(
        immagine,
        Rgb([255u8, 255u8, 255u8]),
        x0 + riempimento,
        y0 + riempimento,
        scala,
        &font,
        testo,
    );
    Ok(())
}

/// Scurisce leggermente il rettangolo indicato (velo sotto il testo).
fn velo(immagine: &mut RgbImage, x0: i32, y0: i32, x1: i32, y1: i32, alfa: f32) {
    let (w, h) = (immagine.width() as i32, immagine.height() as i32);
    for y in y0.max(0)..y1.min(h) {
        for x in x0.max(0)..x1.min(w) {
            let p = immagine.get_pixel_mut(x as u32, y as u32);
            for canale in 0..3 {
                p.0[canale] = (p.0[canale] as f32 * (1.0 - alfa)).round() as u8;
            }
        }
    }
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

        // In alto a sinistra non cambia nulla: la targhetta sta in basso a
        // destra.
        for y in 0..300 {
            for x in 0..400 {
                assert_eq!(immagine.get_pixel(x, y), prima.get_pixel(x, y));
            }
        }

        // In basso a destra il velo scurisce molti pixel...
        let scuriti = (300..600)
            .flat_map(|y| (400..800).map(move |x| (x, y)))
            .filter(|&(x, y)| {
                immagine.get_pixel(x, y).0[0] < prima.get_pixel(x, y).0[0]
            })
            .count();
        assert!(scuriti > 2000, "velo troppo piccolo: {scuriti} pixel");

        // ...e su quel velo restano i pixel bianchi del testo.
        let testo = (300..600)
            .flat_map(|y| (400..800).map(move |x| (x, y)))
            .any(|(x, y)| immagine.get_pixel(x, y).0 == [255, 255, 255]);
        assert!(testo, "il testo della data non e' stato disegnato");
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
