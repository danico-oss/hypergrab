//! Foreground window capture and composition of the date and time.
//!
//! The capture uses the Windows APIs: it takes the window that at that moment
//! is in the foreground (Hypergrab has just hidden itself) and copies its
//! pixels from the screen. If the window is not usable (minimized, closed
//! between the click and the capture) it falls back to the whole virtual screen.

use chrono::{DateTime, Local};
use image::{Rgb, RgbImage};
use std::io::Write;
use std::path::{Path, PathBuf};

use crate::i18n::Language;

/// Result of a capture.
pub struct Cattura {
    /// Captured pixels.
    pub immagine: RgbImage,
    /// Title of the captured window (empty if it fell back to the screen).
    pub finestra: String,
}

// ---------------------------------------------------------------------------
// Date and time
// ---------------------------------------------------------------------------

/// Date and time format for the active language.
pub fn formato_data_ora(lang: Language) -> &'static str {
    match lang {
        Language::En => "%m/%d/%Y %H:%M:%S",
        Language::It | Language::PtBr => "%d/%m/%Y %H:%M:%S",
    }
}

/// Date and time as they will be stamped onto the image.
pub fn testo_data_ora(lang: Language, quando: DateTime<Local>) -> String {
    quando.format(formato_data_ora(lang)).to_string()
}

// ---------------------------------------------------------------------------
// Capture
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
        // Foreground window: it is the application that remains visible.
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

        // Fallback: whole virtual screen (all the monitors together).
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

/// True if the window can be photographed: it exists, it is visible, it is not
/// reduced to an icon and it is not a window of this application.
///
/// Hypergrab hides itself before capturing, but Windows may take a moment
/// to give the foreground to the underlying application: until that happens
/// the capture would concern Hypergrab itself.
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

/// True if the application to be documented is already in the foreground.
#[cfg(windows)]
pub fn primo_piano_pronto() -> bool {
    unsafe { utilizzabile(windows_sys::Win32::UI::WindowsAndMessaging::GetForegroundWindow()) }
}

#[cfg(not(windows))]
pub fn primo_piano_pronto() -> bool {
    false
}

/// Copies the indicated rectangle from the screen.
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
            // Negative height: the rows arrive from top to bottom.
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
                // From BGRA to RGB: the alpha channel of the screen bitmaps is
                // not significant.
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

/// Window title, for the status messages.
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
// Composition of the date and time onto the image
// ---------------------------------------------------------------------------

/// Bright orange with which the date and time are stamped onto the image.
const ARANCIONE_DATA_ORA: Rgb<u8> = Rgb([255, 165, 0]);

/// Color of the outline around the text.
const CONTORNO_DATA_ORA: Rgb<u8> = Rgb([0, 0, 0]);

/// Draws the date and time at the bottom right, in bright orange with no background.
///
/// The font is taken from the system fonts. Around the text a solid dark
/// outline is drawn, which sets it off both from light and from dark
/// backgrounds without the need for a colored label.
pub fn compone_data_ora(immagine: &mut RgbImage, testo: &str) -> Result<(), String> {
    let font = carica_font()?;
    let (larghezza, altezza) = immagine.dimensions();
    let lato = larghezza.min(altezza) as f32;
    let scala = ab_glyph::PxScale::from((lato * 0.05).clamp(15.0, 72.0));
    let (testo_w, testo_h) = imageproc::drawing::text_size(scala, &font, testo);

    let margine = (lato * 0.02).clamp(8.0, 30.0) as i32;
    let x = larghezza as i32 - margine - testo_w as i32;
    let y = altezza as i32 - margine - testo_h as i32;

    // Outline: the dark text is repeated at every offset pixel
    // around the position, so that a solid halo stays around the letters.
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

    // Bright orange, drawn twice: the edge pixels do not stay
    // dampened by the antialiasing and the text comes out solid.
    for _ in 0..2 {
        imageproc::drawing::draw_text_mut(immagine, ARANCIONE_DATA_ORA, x, y, scala, &font, testo);
    }
    Ok(())
}

/// Looks for a usable system font for the text on the image.
fn carica_font() -> Result<ab_glyph::FontVec, String> {
    // Segoe UI is on every modern Windows; the others are fallbacks.
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
// Saving
// ---------------------------------------------------------------------------

/// Name of the JPG file for a test code and a serial number: `T-01-001.jpg`.
///
/// The serial number keeps the evidence of the same test distinct: every save
/// writes a new file instead of overwriting the previous one.
pub fn nome_file(codice: &str, seriale: u32) -> String {
    format!("{}-{seriale:03}.jpg", sanitizza(codice))
}

/// First free serial number of a code in the given folder: 1, 2, 3...
///
/// The numbering restarts for each test code, so `T-01` and `T-02` both begin
/// at 001. The first unused number is used (a gap left by a deleted file is
/// filled again).
pub fn seriale_libero(cartella: &Path, codice: &str) -> u32 {
    let mut seriale = 1;
    while cartella.join(nome_file(codice, seriale)).exists() {
        seriale += 1;
    }
    seriale
}

/// Path of the JPG file: same folder as the list file, name = code + serial
/// number of the first file still free.
pub fn percorso_salvataggio(cartella: &Path, codice: &str) -> PathBuf {
    cartella.join(nome_file(codice, seriale_libero(cartella, codice)))
}

/// Makes a code usable as a file name on Windows.
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
    // Windows does not allow trailing spaces and dots.
    nome = nome.trim().trim_end_matches(|c| c == '.' || c == ' ').to_string();
    if nome.is_empty() {
        return "test".to_string();
    }
    // Names reserved by the system (CON, PRN, ...): a suffix is enough.
    const RISERVATI: [&str; 22] = [
        "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7",
        "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
    ];
    if RISERVATI.contains(&nome.to_uppercase().as_str()) {
        nome.push('_');
    }
    nome
}

/// Writes the image as JPEG (quality 90) in the given path.
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

    /// A window that does not exist cannot be photographed.
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
        // Empty code or made only of forbidden characters: fallback name.
        assert_eq!(sanitizza(""), "test");
        assert_eq!(sanitizza("///"), "___");
        assert_eq!(sanitizza("   "), "test");
        // Names reserved by the system.
        assert_eq!(sanitizza("CON"), "CON_");
        assert_eq!(sanitizza("com1"), "com1_");
    }

    #[test]
    fn nome_del_file_con_seriale() {
        assert_eq!(nome_file("T-01", 1), "T-01-001.jpg");
        assert_eq!(nome_file("T-01", 12), "T-01-012.jpg");
        // Over 999 the numbering grows, it is not cut off.
        assert_eq!(nome_file("T-01", 1000), "T-01-1000.jpg");
        assert_eq!(nome_file("a/b", 3), "a_b-003.jpg");
        assert_eq!(nome_file("CON", 1), "CON_-001.jpg");
    }

    /// Unique temporary folder for the tests.
    fn cartella_di_prova(nome: &str) -> PathBuf {
        std::env::temp_dir().join(format!(
            "hypergrab_{}_{}_{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos(),
            nome
        ))
    }

    #[test]
    fn percorso_con_seriale_libero() {
        let cartella = cartella_di_prova("seriale");
        std::fs::create_dir_all(&cartella).unwrap();
        let immagine = RgbImage::from_pixel(8, 6, Rgb([10u8, 20u8, 30u8]));

        // Empty folder: the numbering starts at 001.
        let primo = percorso_salvataggio(&cartella, "T-01");
        assert_eq!(primo, cartella.join("T-01-001.jpg"));

        // Once the file has been written, the next one takes the
        // following number instead of overwriting it.
        salva_jpg(&immagine, &primo).unwrap();
        let secondo = percorso_salvataggio(&cartella, "T-01");
        assert_eq!(secondo, cartella.join("T-01-002.jpg"));
        salva_jpg(&immagine, &secondo).unwrap();
        assert_eq!(seriale_libero(&cartella, "T-01"), 3);

        // The numbering restarts for every code.
        assert_eq!(
            percorso_salvataggio(&cartella, "T-02"),
            cartella.join("T-02-001.jpg")
        );

        // A hole left by a deleted file is filled again.
        std::fs::remove_file(&secondo).unwrap();
        assert_eq!(
            percorso_salvataggio(&cartella, "T-01"),
            cartella.join("T-01-002.jpg")
        );

        let _ = std::fs::remove_dir_all(&cartella);
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
        // The file must be a real JPEG: the signature starts with FF D8.
        let bytes = std::fs::read(&path).unwrap();
        assert_eq!(&bytes[0..2], &[0xFF, 0xD8]);
        let _ = std::fs::remove_file(&path);
    }

    #[test]
    fn compone_data_e_ora_in_basso_a_destra() {
        // The test uses the system font: if it is missing (non-Windows
        // system), it is skipped.
        let Ok(_) = carica_font() else {
            eprintln!("carattere di sistema assente: verifica saltata");
            return;
        };

        let mut immagine = RgbImage::from_pixel(800, 600, Rgb([255u8, 255u8, 255u8]));
        let prima = immagine.clone();
        compone_data_ora(&mut immagine, "18/10/2026 14:03:22").unwrap();

        // At the top left nothing changes: the text is at the bottom
        // right.
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
                // Bright orange...
                if nuovo.0[0] > 200 && nuovo.0[1] > 80 && nuovo.0[1] < 230 && nuovo.0[2] < 100 {
                    arancioni += 1;
                }
                // ...with the dark outline that makes it readable.
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

        // Around the text nothing is drawn: neither a veil nor a
        // label. The pixels close to the rectangle of the text must
        // have remained intact.
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

    /// Verifies the copy chain from the screen (GetDC/BitBlt/GetDIBits)
    /// without depending on the foreground window: it copies a corner of the
    /// screen. In an environment with no screen the test is skipped.
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
