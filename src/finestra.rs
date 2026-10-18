//! Hiding and re-showing the Hypergrab window with the Windows APIs.
//!
//! To photograph the application that stays in the foreground the Hypergrab
//! window has to disappear. If it were hidden with an egui command, though,
//! Windows would stop sending WM_PAINT to an invisible window: the drawing
//! loop would stop and would no longer be able to re-show it (the window
//! would stay hidden until the program is closed). Therefore hiding and
//! re-showing happen directly, even from another thread.

/// The window of this application.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Finestra {
    hwnd: isize,
}

#[cfg(windows)]
impl Finestra {
    /// The Hypergrab window in the foreground, if there is one.
    ///
    /// The capture starts from a key press or from a click on the button, so
    /// the foreground window is exactly this one; it still checks that it
    /// belongs to this process, in order not to touch windows of other
    /// programs.
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

    /// Hides the window: the underlying application returns to the foreground.
    pub fn nascondi(self) {
        use windows_sys::Win32::UI::WindowsAndMessaging::{ShowWindow, SW_HIDE};

        unsafe { ShowWindow(self.hwnd as _, SW_HIDE) };
    }

    /// Re-shows the window and brings it back to the front.
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
            // The foreground normally belongs to the active application: the
            // attempt may fail, but the window is anyway visible
            // and in front of the others.
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

    /// A reference that does not exist must not make anything panic.
    #[test]
    fn riferimento_inesistente_non_ha_effetto() {
        let finta = Finestra { hwnd: 0x1234 };
        finta.nascondi();
        finta.mostra();
    }
}
