//! Opt-in autostart via `HKCU\Software\Microsoft\Windows\CurrentVersion\Run`.
//!
//! Off by default. The user enables it from the tray menu; nothing is
//! written to the registry without that explicit choice.

use crate::win32::*;

const RUN_KEY: &str = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
const VALUE_NAME: &str = "JinJuApp";

/// The exact command line stored in the Run key (pure, unit-tested).
pub fn command_line(exe_path: &str) -> String {
    format!("\"{exe_path}\" --quiet")
}

fn to_wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain([0]).collect()
}

/// Is autostart currently enabled for this executable?
pub fn is_enabled(exe_path: &str) -> bool {
    unsafe {
        let sub = to_wide(RUN_KEY);
        let name = to_wide(VALUE_NAME);
        let mut hkey: HKEY = core::ptr::null_mut();
        if RegOpenKeyExW(HKEY_CURRENT_USER, sub.as_ptr(), 0, KEY_READ, &mut hkey) != 0 {
            return false;
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
            return false;
        }
        let stored: Vec<u16> = data[..(size as usize).min(data.len())]
            .as_chunks::<2>()
            .0
            .iter()
            .map(|c| u16::from_le_bytes(*c))
            .take_while(|&c| c != 0)
            .collect();
        let stored = String::from_utf16_lossy(&stored);
        stored.trim() == command_line(exe_path).trim()
    }
}

/// Enable autostart for `exe_path`.
pub fn enable(exe_path: &str) -> Result<(), String> {
    let cmd = command_line(exe_path);
    let mut wide = to_wide(&cmd);
    let bytes: Vec<u8> = wide.iter().flat_map(|c| c.to_le_bytes()).collect();
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
        wide.clear();
        Ok(())
    }
}

/// Disable autostart. Succeeds (no-op) if it was not enabled.
pub fn disable() -> Result<(), String> {
    unsafe {
        let sub = to_wide(RUN_KEY);
        let name = to_wide(VALUE_NAME);
        let mut hkey: HKEY = core::ptr::null_mut();
        if RegOpenKeyExW(HKEY_CURRENT_USER, sub.as_ptr(), 0, KEY_SET_VALUE, &mut hkey) != 0 {
            return Ok(()); // nothing to remove
        }
        let rc = RegDeleteValueW(hkey, name.as_ptr());
        RegCloseKey(hkey);
        if rc != 0 && rc != 2 {
            // 2 = ERROR_FILE_NOT_FOUND: already absent
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
}
