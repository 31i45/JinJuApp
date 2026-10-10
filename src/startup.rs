//! Opt-in autostart via `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`.
//!
//! Off by default. The user enables it from the tray menu; nothing is
//! written to the registry without that explicit choice.

use crate::win32::*;

const RUN_KEY: &str = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
const VALUE_NAME: &str = "JinJuApp";
const ERROR_FILE_NOT_FOUND: i32 = 2;

/// The exact command line stored in the Run key (pure, unit-tested).
pub fn command_line(exe_path: &str) -> String {
    format!("\"{exe_path}\" --quiet")
}

fn to_wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain([0]).collect()
}

/// The command line currently stored in the Run key, if any (trimmed).
/// Distinguishing "no entry" from "a different entry" lets the tests
/// avoid touching a developer's real autostart toggle.
fn stored_command() -> Option<String> {
    unsafe {
        let sub = to_wide(RUN_KEY);
        let name = to_wide(VALUE_NAME);
        let mut hkey: HKEY = core::ptr::null_mut();
        if RegOpenKeyExW(HKEY_CURRENT_USER, sub.as_ptr(), 0, KEY_READ, &mut hkey) != 0 {
            return None;
        }
        let mut data: [u8; 1024] = [0; 1024];
        let mut size: u32 = 1024;
        let mut ty: u32 = 0;
        let ok = RegQueryValueExW(
            hkey,
            name.as_ptr(),
            core::ptr::null_mut(),
            &mut ty,
            data.as_mut_ptr(),
            &mut size,
        );
        RegCloseKey(hkey);
        if ok != 0 || ty != REG_SZ {
            return None;
        }
        let wide: Vec<u16> = data[..(size as usize).min(data.len())]
            .as_chunks::<2>()
            .0
            .iter()
            .map(|c| u16::from_le_bytes(*c))
            .take_while(|&c| c != 0)
            .collect();
        Some(String::from_utf16_lossy(&wide).trim().to_string())
    }
}

/// Is autostart currently enabled for this executable?
pub fn is_enabled(exe_path: &str) -> bool {
    stored_command().is_some_and(|c| c == command_line(exe_path))
}

/// Enable autostart for `exe_path`.
pub fn enable(exe_path: &str) -> Result<(), String> {
    let cmd = command_line(exe_path);
    let bytes: Vec<u8> = to_wide(&cmd).iter().flat_map(|c| c.to_le_bytes()).collect();
    unsafe {
        let sub = to_wide(RUN_KEY);
        let name = to_wide(VALUE_NAME);
        let mut hkey: HKEY = core::ptr::null_mut();
        let rc = RegCreateKeyExW(
            HKEY_CURRENT_USER,
            sub.as_ptr(),
            0,
            core::ptr::null(),
            0,
            KEY_SET_VALUE,
            core::ptr::null_mut(),
            &mut hkey,
            core::ptr::null_mut(),
        );
        if rc != 0 {
            return Err(format!("无法打开注册表 Run 键 (错误 {rc})"));
        }
        let rc = RegSetValueExW(
            hkey,
            name.as_ptr(),
            0,
            REG_SZ,
            bytes.as_ptr(),
            bytes.len() as u32,
        );
        RegCloseKey(hkey);
        if rc != 0 {
            return Err(format!("无法写入自启动项 (错误 {rc})"));
        }
        Ok(())
    }
}

/// Disable autostart. Succeeds (no-op) if there is nothing to remove.
pub fn disable() -> Result<(), String> {
    unsafe {
        let sub = to_wide(RUN_KEY);
        let name = to_wide(VALUE_NAME);
        let mut hkey: HKEY = core::ptr::null_mut();
        let rc = RegOpenKeyExW(HKEY_CURRENT_USER, sub.as_ptr(), 0, KEY_SET_VALUE, &mut hkey);
        if rc == ERROR_FILE_NOT_FOUND {
            return Ok(()); // Run key itself absent: nothing to remove
        }
        if rc != 0 {
            return Err(format!("无法打开注册表 Run 键 (错误 {rc})"));
        }
        let rc = RegDeleteValueW(hkey, name.as_ptr());
        RegCloseKey(hkey);
        if rc != 0 && rc != ERROR_FILE_NOT_FOUND {
            return Err(format!("无法删除自启动项 (错误 {rc})"));
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn command_line_quotes_path_and_adds_quiet() {
        assert_eq!(
            command_line(r"C:\Apps\JinJuApp.exe"),
            r#""C:\Apps\JinJuApp.exe" --quiet"#
        );
    }

    #[test]
    fn command_line_handles_spaces() {
        assert_eq!(
            command_line(r"C:\Program Files\JinJu App\JinJuApp.exe"),
            r#""C:\Program Files\JinJu App\JinJuApp.exe" --quiet"#
        );
    }

    /// Real-registry integration test: exercises the exact FFI path
    /// (including the access-mask constants) that the tray toggle uses.
    /// This is the test that would have caught the `KEY_SET_VALUE`
    /// typo that broke the toggle with ERROR_ACCESS_DENIED.
    ///
    /// Skipped on machines where a foreign autostart entry exists, so a
    /// developer's real toggle is never touched; the drop guard removes
    /// whatever we wrote even when an assertion fails.
    #[test]
    fn enable_disable_round_trip() {
        struct Cleanup;
        impl Drop for Cleanup {
            fn drop(&mut self) {
                let _ = disable();
            }
        }

        let test_path = r"C:\__jinju_test__\JinJuApp.exe";
        if stored_command().is_some_and(|c| c != command_line(test_path)) {
            return; // a real (non-test) entry is present — leave it alone
        }

        let _guard = Cleanup;
        assert!(!is_enabled(test_path), "precondition: no test entry");
        enable(test_path).expect("enable must succeed");
        assert!(is_enabled(test_path), "entry must be readable back");
        disable().expect("disable must succeed");
        assert!(!is_enabled(test_path), "entry must be gone");
    }
}
