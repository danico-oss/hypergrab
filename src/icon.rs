//! Icona del programma: una cornice di cattura (quattro angoli da ritaglio)
//! con al centro il diaframma dell'obiettivo, come il mirino che inquadra
//! l'area da fotografare.
//!
//! Il disegno e' procedurale (nessuna libreria di immagini, nessun file
//! esterno): la stessa funzione produce i pixel della finestra e il file
//! `assets/hypergrab.ico` usato per l'icona del .exe.

use eframe::egui;

/// Blu degli angoli della cornice.
const CORNICE: [f32; 3] = [30.0, 92.0, 204.0];
/// Rosso caldo del diaframma centrale.
const DIAFRAMMA: [f32; 3] = [232.0, 84.0, 40.0];
/// Fattore di supercampionamento: i bordi restano nitidi senza antialiasing
/// esterno, poi l'immagine viene ridotta con filtro a scatola.
const SS: u32 = 4;

/// Icona da passare a `ViewportBuilder::with_icon`.
pub fn egui_icon(size: u32) -> egui::IconData {
    egui::IconData {
        rgba: render(size),
        width: size,
        height: size,
    }
}

/// Pixel RGBA (canale alfa non premoltiplicato) dell'icona a `size`x`size`.
pub fn render(size: u32) -> Vec<u8> {
    let size = size.max(8);
    let big = size * SS;
    let mut buf = vec![0u8; (big * big * 4) as usize];
    disegna(&mut buf, big);
    riduci(&buf, big, size)
}

/// Disegna la cornice e il diaframma su una tela `n`x`n` trasparente
/// (coordinate normalizzate 0..1).
fn disegna(buf: &mut [u8], n: u32) {
    // Spessore dei bracci, lunghezza dei bracci e margine dai bordi.
    let (t, a, m) = (0.085, 0.300, 0.115);
    let (x0, y0, x1, y1) = (m, m, 1.0 - m, 1.0 - m);

    // Angoli da ritaglio: due rettangoli per angolo (orizzontale + verticale).
    // In senso orario: alto sinistra, alto destra, basso destra, basso sinistra.
    rettangolo(buf, n, x0, y0, x0 + a, y0 + t, CORNICE); // alto sinistra orizzontale
    rettangolo(buf, n, x0, y0, x0 + t, y0 + a, CORNICE); // alto sinistra verticale
    rettangolo(buf, n, x1 - a, y0, x1, y0 + t, CORNICE); // alto destra orizzontale
    rettangolo(buf, n, x1 - t, y0, x1, y0 + a, CORNICE); // alto destra verticale
    rettangolo(buf, n, x1 - a, y1 - t, x1, y1, CORNICE); // basso destra orizzontale
    rettangolo(buf, n, x1 - t, y1 - a, x1, y1, CORNICE); // basso destra verticale
    rettangolo(buf, n, x0, y1 - t, x0 + a, y1, CORNICE); // basso sinistra orizzontale
    rettangolo(buf, n, x0, y1 - a, x0 + t, y1, CORNICE); // basso sinistra verticale

    // Diaframma: cerchio pieno al centro, con un disco piu' chiaro dentro.
    cerchio(buf, n, RAGGIO_DIAFRAMMA, DIAFRAMMA);
    let chiaro = [255.0, 214.0, 190.0];
    cerchio(buf, n, RAGGIO_DIAFRAMMA * 0.55, chiaro);
}

/// Raggio esterno del diaframma (in frazione del lato).
const RAGGIO_DIAFRAMMA: f32 = 0.150;

/// Colore di un punto, con l'alfa a 255: comodo per scrivere i pixel.
fn scrivi(buf: &mut [u8], n: u32, x: u32, y: u32, rgb: [f32; 3]) {
    let o = ((y * n + x) * 4) as usize;
    buf[o] = rgb[0].min(255.0) as u8;
    buf[o + 1] = rgb[1].min(255.0) as u8;
    buf[o + 2] = rgb[2].min(255.0) as u8;
    buf[o + 3] = 255;
}

/// Rettangolo pieno in coordinate normalizzate (`x0 < x1`, `y0 < y1`).
fn rettangolo(buf: &mut [u8], n: u32, x0: f32, y0: f32, x1: f32, y1: f32, rgb: [f32; 3]) {
    let f = n as f32;
    let (px0, px1) = ((x0 * f) as u32, (x1 * f) as u32);
    let (py0, py1) = ((y0 * f) as u32, (y1 * f) as u32);
    for y in py0..py1.min(n) {
        for x in px0..px1.min(n) {
            scrivi(buf, n, x, y, rgb);
        }
    }
}

/// Cerchio pieno di centro (0.5, 0.5) e raggio dato in frazione del lato.
fn cerchio(buf: &mut [u8], n: u32, raggio: f32, rgb: [f32; 3]) {
    let f = n as f32;
    let c = 0.5 * f;
    let r = raggio * f;
    let (a, b) = ((c - r).max(0.0) as u32, (c + r).min(f - 1.0) as u32);
    for y in a..=b {
        for x in a..=b {
            let (dx, dy) = (x as f32 + 0.5 - c, y as f32 + 0.5 - c);
            if dx * dx + dy * dy <= r * r {
                scrivi(buf, n, x, y, rgb);
            }
        }
    }
}

/// Riduzione con filtro a scatola: media pesata per l'alfa, cosi' i bordi non
/// diventano scuri quando l'immagine viene vista su sfondo chiaro.
fn riduci(buf: &[u8], big: u32, size: u32) -> Vec<u8> {
    let s = (big / size) as usize;
    let mut out = vec![0u8; (size * size * 4) as usize];
    for y in 0..size as usize {
        for x in 0..size as usize {
            let (mut ra, mut ga, mut ba, mut aa) = (0.0f32, 0.0f32, 0.0f32, 0.0f32);
            for j in 0..s {
                for i in 0..s {
                    let o = (((y * s + j) as u32 * big + (x * s + i) as u32) * 4) as usize;
                    let a = buf[o + 3] as f32;
                    ra += buf[o] as f32 * a;
                    ga += buf[o + 1] as f32 * a;
                    ba += buf[o + 2] as f32 * a;
                    aa += a;
                }
            }
            let k = (s * s) as f32;
            let o = (y * size as usize + x) * 4;
            if aa > 0.0 {
                out[o] = (ra / aa).round() as u8;
                out[o + 1] = (ga / aa).round() as u8;
                out[o + 2] = (ba / aa).round() as u8;
                out[o + 3] = (aa / k).round() as u8;
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::Path;

    /// Taglie inserite nel file `.ico` (come da pratica corrente su Windows).
    const TAGLIE: [u32; 6] = [16, 24, 32, 48, 64, 256];

    /// Scrive `assets/hypergrab.ico` (rigenerato ogni volta che girano i test):
    /// serve al `build.rs` per incorporare l'icona nel file .exe.
    #[test]
    fn scrivi_file_ico() {
        std::fs::create_dir_all("assets").expect("cartella assets");
        let bytes = costruisci_ico(&TAGLIE);
        assert_eq!(&bytes[0..4], &[0, 0, 1, 0], "intestazione ICO non valida");
        std::fs::write("assets/hypergrab.ico", &bytes).expect("scrittura assets/hypergrab.ico");
    }

    /// Controllo che le voci dell'ico siano coerenti (offset e dimensioni).
    #[test]
    fn ico_interna_coerente() {
        let bytes = costruisci_ico(&TAGLIE);
        let n = u16::from_le_bytes([bytes[4], bytes[5]]) as usize;
        assert_eq!(n, TAGLIE.len());
        for i in 0..n {
            let e = 6 + i * 16;
            let (dim, off) = (
                u32::from_le_bytes(bytes[e + 8..e + 12].try_into().unwrap()) as usize,
                u32::from_le_bytes(bytes[e + 12..e + 16].try_into().unwrap()) as usize,
            );
            assert!(off + dim <= bytes.len(), "voce {i} fuori dal file");
            let s = TAGLIE[i] as usize;
            assert_eq!(dim, 40 + s * s * 4 + ((s + 31) / 32 * 4) * s, "voce {i}");
        }
    }

    /// Genera un "foglio di contatto" PNG da guardare per verificare il
    /// disegno sia su sfondo chiaro sia su sfondo scuro, e salva anche
    /// `assets/hypergrab.png` a 256 px come riferimento.
    #[test]
    fn anteprime_png() {
        std::fs::create_dir_all("assets").expect("cartella assets");
        let (tela, w, h) = foglio(&[256, 64, 48, 32, 16]);
        let dir = std::env::var("TEMP").unwrap_or_else(|_| ".".to_string());
        scrivi_png(
            &std::path::Path::new(&dir).join("hypergrab_icona_contatto.png"),
            &tela,
            w,
            h,
        );
        let rgba = render(256);
        scrivi_png(Path::new("assets/hypergrab.png"), &rgba, 256, 256);
    }

    /// Campiona alcuni pixel per verificare che la geometria sia quella
    /// voluta: fuori dalla cornice e' trasparente, gli angoli sono blu, il
    /// centro e' il diaframma rosso e il vuoto al centro di ogni lato resta
    /// libero (la cornice non e' un rettangolo chiuso).
    #[test]
    fn geometria_cornice() {
        let s = 256u32;
        let rgba = render(s);
        let px = |nx: f32, ny: f32| -> [u8; 4] {
            let o = ((ny * s as f32) as u32 * s + (nx * s as f32) as u32) as usize * 4;
            [rgba[o], rgba[o + 1], rgba[o + 2], rgba[o + 3]]
        };

        // Fuori dagli angoli: trasparente.
        for (x, y) in [(0.02, 0.02), (0.98, 0.02), (0.02, 0.98), (0.98, 0.98)] {
            assert_eq!(px(x, y)[3], 0, "trasparente a ({x}, {y})");
        }
        // Centro: diaframma opaco e caldo (piu' rosso che blu).
        let centro = px(0.50, 0.50);
        assert_eq!(centro[3], 255, "diaframma opaco");
        assert!(centro[0] > centro[2], "diaframma caldo: {centro:?}");
        // Angolo alto sinistra: cornice opaca e fredda (piu' blu che rossa).
        let angolo = px(0.15, 0.15);
        assert_eq!(angolo[3], 255, "cornice opaca");
        assert!(angolo[2] > angolo[0], "cornice blu: {angolo:?}");
        // Metà di ogni lato: la cornice e' aperta, quindi trasparente.
        for (x, y) in [(0.50, 0.15), (0.50, 0.85), (0.15, 0.50), (0.85, 0.50)] {
            assert_eq!(px(x, y)[3], 0, "apertura della cornice a ({x}, {y})");
        }
    }

    // --- costruzione file -------------------------------------------------

    fn costruisci_ico(taglie: &[u32]) -> Vec<u8> {
        let voci: Vec<(u32, Vec<u8>)> = taglie
            .iter()
            .map(|&s| (s, dati_bmp_ico(&render(s), s)))
            .collect();
        let n = voci.len() as u16;
        let mut out: Vec<u8> = Vec::new();
        out.extend_from_slice(&0u16.to_le_bytes()); // riservato
        out.extend_from_slice(&1u16.to_le_bytes()); // tipo = icone
        out.extend_from_slice(&n.to_le_bytes());
        let mut offset = 6 + 16 * n as u32;
        for (s, data) in &voci {
            let b = if *s >= 256 { 0u8 } else { *s as u8 };
            out.push(b); // larghezza (0 = 256)
            out.push(b); // altezza
            out.push(0); // colori in palette
            out.push(0); // riservato
            out.extend_from_slice(&1u16.to_le_bytes()); // planes
            out.extend_from_slice(&32u16.to_le_bytes()); // bit/pixel
            out.extend_from_slice(&(data.len() as u32).to_le_bytes());
            out.extend_from_slice(&offset.to_le_bytes());
            offset += data.len() as u32;
        }
        for (_, data) in &voci {
            out.extend_from_slice(data);
        }
        out
    }

    /// Voce ICO in formato BMP 32 bit (XOR + maschera AND).
    fn dati_bmp_ico(rgba: &[u8], s: u32) -> Vec<u8> {
        let passo_maschera = s.div_ceil(32) * 4;
        let xor = s * s * 4;
        let totale = xor + passo_maschera * s;
        let mut d: Vec<u8> = Vec::with_capacity(40 + totale as usize);
        d.extend_from_slice(&40u32.to_le_bytes()); // dimensione BITMAPINFOHEADER
        d.extend_from_slice(&(s as i32).to_le_bytes());
        d.extend_from_slice(&((s * 2) as i32).to_le_bytes()); // altezza doppia
        d.extend_from_slice(&1u16.to_le_bytes()); // planes
        d.extend_from_slice(&32u16.to_le_bytes()); // bit/pixel
        d.extend_from_slice(&0u32.to_le_bytes()); // compressione BI_RGB
        d.extend_from_slice(&totale.to_le_bytes());
        d.extend_from_slice(&0i32.to_le_bytes()); // dpi x
        d.extend_from_slice(&0i32.to_le_bytes()); // dpi y
        d.extend_from_slice(&0u32.to_le_bytes()); // colori in palette
        d.extend_from_slice(&0u32.to_le_bytes()); // colori importanti
        for y in (0..s).rev() {
            for x in 0..s {
                let o = ((y * s + x) * 4) as usize;
                d.push(rgba[o + 2]); // B
                d.push(rgba[o + 1]); // G
                d.push(rgba[o]); // R
                d.push(rgba[o + 3]); // A (maschera AND: basta il canale alfa)
            }
        }
        d.extend(std::iter::repeat_n(0u8, totale as usize - xor as usize)); // AND a zero
        d
    }

    // --- anteprime --------------------------------------------------------

    /// Compone icone a grandezze diverse su due metà: bianco sopra, grigio
    /// scuro sotto. Restituisce la tela RGBA (larghezza, altezza).
    fn foglio(taglie: &[u32]) -> (Vec<u8>, u32, u32) {
        let w = 16 + taglie.iter().sum::<u32>() + 8 * (taglie.len() as u32 - 1);
        let riga = 16 + 256 + 8;
        let h = riga * 2;
        let mut tela = vec![0u8; (w * h * 4) as usize];
        let (chiaro, scuro) = ([255u8; 3], [44u8, 46, 50]);
        for y in 0..h {
            let sfondo = if y < riga { chiaro } else { scuro };
            for x in 0..w {
                let o = ((y * w + x) * 4) as usize;
                tela[o..o + 3].copy_from_slice(&sfondo);
                tela[o + 3] = 255;
            }
        }
        let mut x = 8;
        for &s in taglie {
            let rgba = render(s);
            for (meta, sfondo) in [chiaro, scuro].into_iter().enumerate() {
                let y0 = meta as u32 * riga + 8 + (256 - s) / 2;
                for y in 0..s {
                    for xx in 0..s {
                        let o = ((y * s + xx) * 4) as usize;
                        let a = rgba[o + 3] as u32;
                        let inv = 255 - a;
                        let p = (((y0 + y) * w + x + xx) * 4) as usize;
                        for k in 0..3 {
                            let v = (rgba[o + k] as u32 * a + sfondo[k] as u32 * inv) / 255;
                            tela[p + k] = v as u8;
                        }
                    }
                }
            }
            x += s + 8;
        }
        (tela, w, h)
    }

    // --- PNG (deflate con blocchi "stored": niente dipendenze) ------------

    fn scrivi_png(path: &Path, rgba: &[u8], w: u32, h: u32) {
        // Filtro "None" su ogni riga, poi zlib con blocchi non compressi.
        let mut grezzo = Vec::with_capacity((w * h * 4 + h) as usize);
        for y in 0..h {
            grezzo.push(0);
            grezzo.extend(&rgba[(y * w * 4) as usize..((y + 1) * w * 4) as usize]);
        }
        let mut idat = vec![0x78, 0x01];
        let mut i = 0;
        while i < grezzo.len() {
            let n = (grezzo.len() - i).min(0xFFFF);
            let ultimo = i + n == grezzo.len();
            idat.push(if ultimo { 1 } else { 0 }); // BFINAL + BTYPE=00
            idat.extend(&(n as u16).to_le_bytes());
            idat.extend(&(!(n as u16)).to_le_bytes());
            idat.extend(&grezzo[i..i + n]);
            i += n;
        }
        idat.extend(&adler32(&grezzo).to_be_bytes());

        let mut out: Vec<u8> = Vec::new();
        out.extend(&[0x89, b'P', b'N', b'G', 13, 10, 26, 10]);
        let mut ihdr = Vec::new();
        ihdr.extend(&w.to_be_bytes());
        ihdr.extend(&h.to_be_bytes());
        ihdr.extend(&[8, 6, 0, 0, 0]); // 8 bit, RGBA
        blocco_png(&mut out, b"IHDR", &ihdr);
        blocco_png(&mut out, b"IDAT", &idat);
        blocco_png(&mut out, b"IEND", &[]);
        std::fs::write(path, out).expect("scrittura anteprima");
    }

    fn blocco_png(out: &mut Vec<u8>, tipo: &[u8; 4], dati: &[u8]) {
        out.extend(&(dati.len() as u32).to_be_bytes());
        out.extend(tipo);
        out.extend(dati);
        let crc = crc32_tabellato();
        let mut c = 0xFFFF_FFFFu32;
        for b in tipo.iter().chain(dati.iter()) {
            c = crc[(*b ^ c as u8) as usize] ^ (c >> 8);
        }
        out.extend((c ^ 0xFFFF_FFFF).to_be_bytes());
    }

    fn crc32_tabellato() -> [u32; 256] {
        let mut t = [0u32; 256];
        for (i, v) in t.iter_mut().enumerate() {
            let mut c = i as u32;
            for _ in 0..8 {
                c = if c & 1 == 1 { 0xEDB8_8320 ^ (c >> 1) } else { c >> 1 };
            }
            *v = c;
        }
        t
    }

    fn adler32(dati: &[u8]) -> u32 {
        let (mut a, mut b) = (1u32, 0u32);
        for &x in dati {
            a = (a + x as u32) % 65521;
            b = (b + a) % 65521;
        }
        (b << 16) | a
    }
}
