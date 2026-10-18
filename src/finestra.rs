//! Nascondere e rimostrare la finestra di Hypergrab con le API di Windows.
//!
//! Per fotografare l'applicazione che resta in primo piano la finestra di
//! Hypergrab deve sparire. Se la si nascondesse con un comando di egui, pero',
//! Windows smetterebbe di mandare WM_PAINT a una finestra invisibile: il ciclo
//! di disegno si fermerebbe e non sarebbe piu' in grado di rimostrarla (la
//! finestra resterebbe nascosta fino alla chiusura del programma). Percio'
//! nascondere e rimostrare avvengono direttamente, anche da un altro thread.

/// La finestra di questa applicazione.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Finestra {
    hwnd: isize,
}

#[cfg(windows)]
impl Finestra {
    /// La finestra di Hypergrab in primo piano, se c'e'.
    ///
    /// La cattura parte da un tasto premuto o da un clic sul pulsante, quindi
    /// la finestra in primo piano e' proprio questa; si controlla comunque che
    /// appartenga a questo processo, per non toccare finestre di altri
    /// programmi.
    pub fn attiva() -> Option<Self> {
        use windows_sys::Win32::System::Threading::GetCurrentProcessId;
        use windows_sys::Win32::UI::WindowsAndMessaging::{
            GetForegroundWindow, GetWindowThreadProcessId,
        };

        let hwnd = unsafe { GetForegroundWindow() };
        if hwnd.is_null() {
            return None;
        }
        let mut processo = 0u32;
        unsafe { GetWindowThreadProcessId(hwnd, &mut processo) };
        if processo != unsafe { GetCurrentProcessId() } {
            return None;
        }
        Some(Self {
            hwnd: hwnd as isize,
        })
    }

    /// Nasconde la finestra: l'applicazione sottostante torna in primo piano.
    pub fn nascondi(self) {
        use windows_sys::Win32::UI::WindowsAndMessaging::{ShowWindow, SW_HIDE};

        unsafe { ShowWindow(self.hwnd as _, SW_HIDE) };
    }

    /// Rimostra la finestra e la riporta davanti.
    pub fn mostra(self) {
        use windows_sys::Win32::UI::WindowsAndMessaging::{
            BringWindowToTop, IsIconic, SetForegroundWindow, ShowWindow, SW_RESTORE, SW_SHOW,
        };

        unsafe {
            let hwnd = self.hwnd as _;
            if IsIconic(hwnd) != 0 {
                ShowWindow(hwnd, SW_RESTORE);
            } else {
                ShowWindow(hwnd, SW_SHOW);
            }
            // Il primo piano spetta di norma all'applicazione attiva: il
            // tentativo puo' non riuscire, ma la finestra e' comunque visibile
            // e davanti alle altre.
            BringWindowToTop(hwnd);
            SetForegroundWindow(hwnd);
        }
    }
}

#[cfg(not(windows))]
impl Finestra {
    pub fn attiva() -> Option<Self> {
        None
    }

    pub fn nascondi(self) {}

    pub fn mostra(self) {}
}

#[cfg(all(test, windows))]
mod tests {
    use super::*;

    /// Un riferimento che non esiste non deve far saltare nulla.
    #[test]
    fn riferimento_inesistente_non_ha_effetto() {
        let finta = Finestra { hwnd: 0x1234 };
        finta.nascondi();
        finta.mostra();
    }
}
