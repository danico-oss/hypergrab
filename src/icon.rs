//! Program icon: a capture frame (four crop corners) with the lens
//! diaphragm at the center, like the viewfinder that frames the area to
//! photograph.
//!
//! The drawing is procedural (no image library, no external file): the
//! same function produces the window pixels and the
//! `assets/hypergrab.ico` file used for the .exe icon.

use eframe::egui;

/// Blue of the frame corners.
const CORNICE: [f32; 3] = [30.0, 92.0, 204.0];
/// Warm red of the central diaphragm.
const DIAFRAMMA: [f32; 3] = [232.0, 84.0, 40.0];
/// Supersampling factor: the edges stay sharp without external
/// antialiasing, then the image is downscaled with a box filter.
const SS: u32 = 4;

/// Icon to pass to `ViewportBuilder::with_icon`.
pub fn egui_icon(size: u32) -> egui::IconData {
    egui::IconData {
        rgba: render(size),
        width: size,
        height: size,
    }
}

/// RGBA pixels (non-premultiplied alpha channel) of the `size`x`size` icon.
pub fn render(size: u32) -> Vec<u8> {
    let size = size.max(8);
    let big = size * SS;
    let mut buf = vec![0u8; (big * big * 4) as usize];
    disegna(&mut buf, big);
    riduci(&buf, big, size)
}

/// Draws the frame and the diaphragm on a transparent `n`x`n` canvas
/// (normalized coordinates 0..1).
fn disegna(buf: &mut [u8], n: u32) {
    // Arm thickness, arm length and margin from the edges.
    let (t, a, m) = (0.085, 0.300, 0.115);
    let (x0, y0, x1, y1) = (m, m, 1.0 - m, 1.0 - m);

    // Crop corners: two rectangles per corner (horizontal + vertical).
    // Clockwise: top left, top right, bottom right, bottom left.
    rettangolo(buf, n, x0, y0, x0 + a, y0 + t, CORNICE); // top left horizontal
    rettangolo(buf, n, x0, y0, x0 + t, y0 + a, CORNICE); // top left vertical
    rettangolo(buf, n, x1 - a, y0, x1, y0 + t, CORNICE); // top right horizontal
    rettangolo(buf, n, x1 - t, y0, x1, y0 + a, CORNICE); // top right vertical
    rettangolo(buf, n, x1 - a, y1 - t, x1, y1, CORNICE); // bottom right horizontal
    rettangolo(buf, n, x1 - t, y1 - a, x1, y1, CORNICE); // bottom right vertical
    rettangolo(buf, n, x0, y1 - t, x0 + a, y1, CORNICE); // bottom left horizontal
    rettangolo(buf, n, x0, y1 - a, x0 + t, y1, CORNICE); // bottom left vertical

    // Diaphragm: a filled circle at the center, with a lighter disc inside.
    cerchio(buf, n, RAGGIO_DIAFRAMMA, DIAFRAMMA);
    let chiaro = [255.0, 214.0, 190.0];
    cerchio(buf, n, RAGGIO_DIAFRAMMA * 0.55, chiaro);
}

/// Outer radius of the diaphragm (as a fraction of the side).
const RAGGIO_DIAFRAMMA: f32 = 0.150;

/// Color of a pixel, with alpha at 255: convenient for writing pixels.
fn scrivi(buf: &mut [u8], n: u32, x: u32, y: u32, rgb: [f32; 3]) {
    let o = ((y * n + x) * 4) as usize;
    buf[o] = rgb[0].min(255.0) as u8;
    buf[o + 1] = rgb[1].min(255.0) as u8;
    buf[o + 2] = rgb[2].min(255.0) as u8;
    buf[o + 3] = 255;
}

/// Filled rectangle in normalized coordinates (`x0 < x1`, `y0 < y1`).
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

/// Filled circle centered at (0.5, 0.5) with radius as a fraction of the side.
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

/// Downscaling with a box filter: alpha-weighted average, so the edges do not
/// become dark when the image is viewed on a light background.
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

    /// Sizes inserted in the `.ico` file (as per common practice on Windows).
    const TAGLIE: [u32; 6] = [16, 24, 32, 48, 64, 256];

    /// Writes `assets/hypergrab.ico` (regenerated every time the tests run):
    /// used by `build.rs` to embed the icon into the .exe file.
    #[test]
    fn scrivi_file_ico() {
        std::fs::create_dir_all("assets").expect("cartella assets");
        let bytes = costruisci_ico(&TAGLIE);
        assert_eq!(&bytes[0..4], &[0, 0, 1, 0], "intestazione ICO non valida");
        std::fs::write("assets/hypergrab.ico", &bytes).expect("scrittura assets/hypergrab.ico");
    }

    /// Checks that the ico entries are consistent (offsets and sizes).
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

    /// Generates a PNG "contact sheet" to look at, to check the drawing on
    /// both a light and a dark background, and also saves
    /// `assets/hypergrab.png` at 256 px as a reference.
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

    /// Samples a few pixels to check that the geometry is as intended:
    /// outside the frame it is transparent, the corners are blue, the
    /// center is the red diaphragm and the gap at the middle of each side
    /// stays open (the frame is not a closed rectangle).
    #[test]
    fn geometria_cornice() {
        let s = 256u32;
        let rgba = render(s);
        let px = |nx: f32, ny: f32| -> [u8; 4] {
            let o = ((ny * s as f32) as u32 * s + (nx * s as f32) as u32) as usize * 4;
            [rgba[o], rgba[o + 1], rgba[o + 2], rgba[o + 3]]
        };

        // Outside the corners: transparent.
        for (x, y) in [(0.02, 0.02), (0.98, 0.02), (0.02, 0.98), (0.98, 0.98)] {
            assert_eq!(px(x, y)[3], 0, "trasparente a ({x}, {y})");
        }
        // Center: opaque and warm diaphragm (more red than blue).
        let centro = px(0.50, 0.50);
        assert_eq!(centro[3], 255, "diaframma opaco");
        assert!(centro[0] > centro[2], "diaframma caldo: {centro:?}");
        // Top left corner: opaque and cool frame (more blue than red).
        let angolo = px(0.15, 0.15);
        assert_eq!(angolo[3], 255, "cornice opaca");
        assert!(angolo[2] > angolo[0], "cornice blu: {angolo:?}");
        // Middle of each side: the frame is open, therefore transparent.
        for (x, y) in [(0.50, 0.15), (0.50, 0.85), (0.15, 0.50), (0.85, 0.50)] {
            assert_eq!(px(x, y)[3], 0, "apertura della cornice a ({x}, {y})");
        }
    }

    // --- file construction ------------------------------------------------

    fn costruisci_ico(taglie: &[u32]) -> Vec<u8> {
        let voci: Vec<(u32, Vec<u8>)> = taglie
            .iter()
            .map(|&s| (s, dati_bmp_ico(&render(s), s)))
            .collect();
        let n = voci.len() as u16;
        let mut out: Vec<u8> = Vec::new();
        out.extend_from_slice(&0u16.to_le_bytes()); // reserved
        out.extend_from_slice(&1u16.to_le_bytes()); // type = icons
        out.extend_from_slice(&n.to_le_bytes());
        let mut offset = 6 + 16 * n as u32;
        for (s, data) in &voci {
            let b = if *s >= 256 { 0u8 } else { *s as u8 };
            out.push(b); // width (0 = 256)
            out.push(b); // height
            out.push(0); // palette colors
            out.push(0); // reserved
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

    /// ICO entry in 32-bit BMP format (XOR + AND mask).
    fn dati_bmp_ico(rgba: &[u8], s: u32) -> Vec<u8> {
        let passo_maschera = s.div_ceil(32) * 4;
        let xor = s * s * 4;
        let totale = xor + passo_maschera * s;
        let mut d: Vec<u8> = Vec::with_capacity(40 + totale as usize);
        d.extend_from_slice(&40u32.to_le_bytes()); // size of BITMAPINFOHEADER
        d.extend_from_slice(&(s as i32).to_le_bytes());
        d.extend_from_slice(&((s * 2) as i32).to_le_bytes()); // double height
        d.extend_from_slice(&1u16.to_le_bytes()); // planes
        d.extend_from_slice(&32u16.to_le_bytes()); // bit/pixel
        d.extend_from_slice(&0u32.to_le_bytes()); // BI_RGB compression
        d.extend_from_slice(&totale.to_le_bytes());
        d.extend_from_slice(&0i32.to_le_bytes()); // dpi x
        d.extend_from_slice(&0i32.to_le_bytes()); // dpi y
        d.extend_from_slice(&0u32.to_le_bytes()); // palette colors
        d.extend_from_slice(&0u32.to_le_bytes()); // important colors
        for y in (0..s).rev() {
            for x in 0..s {
                let o = ((y * s + x) * 4) as usize;
                d.push(rgba[o + 2]); // B
                d.push(rgba[o + 1]); // G
                d.push(rgba[o]); // R
                d.push(rgba[o + 3]); // A (AND mask: the alpha channel is enough)
            }
        }
        d.extend(std::iter::repeat_n(0u8, totale as usize - xor as usize)); // AND at zero
        d
    }

    // --- previews ---------------------------------------------------------

    /// Composes icons of different sizes on two halves: white on top, dark
    /// gray below. Returns the RGBA canvas (width, height).
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

    // --- PNG (deflate with "stored" blocks: no dependencies) --------------

    fn scrivi_png(path: &Path, rgba: &[u8], w: u32, h: u32) {
        // "None" filter on every row, then zlib with uncompressed blocks.
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
