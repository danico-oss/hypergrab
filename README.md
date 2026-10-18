# Hypergrab

**Test screenshot capture**

A desktop application in Rust for collecting evidence while qualifying or
validating computerized systems. It imports a list of tests from a CSV or XLSX
file, captures the window of the application that stays in the foreground,
stamps the date and time onto the image and saves the JPG with the test code as
its name.

## Features

- **Importing the list**: CSV/TXT files (separator `;`, comma or tab character,
  recognized automatically) and Excel files (`.xlsx`, `.xlsm`, `.xls`, `.ods`,
  first sheet). Text files are read as UTF-8 and, when that is not valid, as
  Windows-1252: CSV files exported by Excel in Italian do not lose their
  accented characters.
- **Test list**: every row shows the code and the description; a click selects
  it (the selection is always a frame and never a fill). When the file is
  opened the first test is selected, so a capture can be taken right away.
- **Capture**: the "Run capture" button or the **F12** key hide Hypergrab for a
  moment and photograph the window that is in the foreground at that instant.
  If the window cannot be used (for example it is minimized) the whole virtual
  screen is captured, that is, all the monitors together.
- **Date and time**: drawn in the bottom right corner of the image in **bright
  orange, with no background**, with a solid dark outline that keeps them
  readable both on light and on dark backgrounds. The format follows the
  language of the interface: `18/10/2026 14:03:22` in Italian and Portuguese,
  `10/18/2026 14:03:22` in English. The font is taken from the Windows fonts
  (Segoe UI and fallbacks).
- **Preview and saving**: the preview shows exactly the image that will be
  written. "Save JPG" writes the file as JPEG (quality 90) in the **same folder
  as the list file**, with the **name equal to the test code followed by a
  progressive number** (`T-01-001.jpg`, `T-01-002.jpg`, …): the first serial
  number still free is used, so no evidence is ever overwritten.
- **Interface**: light (default) and dark theme, a compact style that also
  suits notebook screens, and the Italian / English / Brazilian Portuguese
  languages. The preferences (theme and language) are remembered in
  `hypergrab_settings.txt`, next to the executable.

## List file format

Two columns, in this order:

| Column | Content | Notes |
| --- | --- | --- |
| 1 | **Test code** | becomes the name of the JPG file (`T-01` → `T-01-001.jpg`) |
| 2 | **Test description** | optional |

The "The first row is a header" checkbox (on by default) discards the first
row: clear it when the file already contains data only. Rows without a code are
ignored.

Sample CSV:

```csv
Code;Description
T-01;Initial check of the environment
T-02;Check of the process parameters
T-03;Access attempt with an unauthorized user
```

A sample list is in `test-data/test_list.csv`.

In an Excel file the same layout in the first sheet is enough; a code stored as
a number (for example `1003`) is used as `1003`, not as `1003.0`.

## How to use it

1. "Open file…" and choose the list (`test-data/test_list.csv` for a try).
2. Bring the application you want to document to the foreground.
3. Select the test in the list and press **F12** (or use "Run capture").
   Hypergrab hides for about half a second and then comes back with the
   preview.
4. Check the preview and press "Save JPG".

The destination path is shown under the preview buttons: it is the first serial
number still free for the selected test, so you always know where the file will
end up.

## Technical notes

- The capture uses the Windows APIs: `GetForegroundWindow` for the window, a
  pixel copy from the screen (`BitBlt` onto a compatible bitmap) and, if
  something is not available, a fallback to the whole virtual screen. The alpha
  channel of the screen bitmaps is not meaningful and is ignored.
- The code sanitizes the file name (`\ / : * ? " < > |`, control characters,
  trailing dots and spaces) and appends a suffix to the names reserved by the
  system (`CON`, `PRN`, `COM1`…).
- The icon is drawn by hand in the code (`src/icon.rs`): the test
  `icon::tests::scrivi_file_ico` regenerates `assets/hypergrab.ico`, which
  `build.rs` embeds into the executable with `rc.exe` together with the version
  data. The `anteprime_png` test also produces a contact sheet with the various
  sizes on a light and a dark background.
- The release profile links against the static CRT (`.cargo/config.toml`), so
  `hypergrab.exe` does not require the Microsoft Visual C++ Redistributable on
  the target machine.

## Building

```powershell
cargo test          # also generates assets/hypergrab.ico
cargo run --release
```

A Rust toolchain for the `x86_64-pc-windows-msvc` target is required; to embed
the icon, `rc.exe` from the Windows SDK is needed (if it is missing, the build
goes on without an icon, with a warning).

## License

GNU General Public License v3.0 —
<https://github.com/danico-oss/hypergrab>

Author: Daniel Consonni. A special thank you to Geovanna Mendes de Souza,
whose observations were the starting point for the development of this
application.
