//! Notification-area icon and its context menu — the app's only handle for
//! user interaction (open the quotes file, toggle autostart, exit).

use crate::win32::*;

/// Callback message the shell sends for tray events.
pub const WM_TRAY: u32 = WM_APP + 1;

/// Menu command ids (also used as `WM_COMMAND` wParam low word).
pub const CMD_OPEN_QUOTES: usize = 100;
pub const CMD_TOGGLE_STARTUP: usize = 101;
pub const CMD_EXIT: usize = 102;

fn widez(s: &str) -> Vec<u16> {
    s.encode_utf16().chain([0]).collect()
}

/// Add the tray icon for `hwnd`.
pub fn add(hwnd: HWND) -> bool {
    unsafe {
        let mut ni = NOTIFYICONDATAW {
            cbSize: core::mem::size_of::<NOTIFYICONDATAW>() as u32,
            hWnd: hwnd,
            uID: 1,
            uFlags: NIF_MESSAGE | NIF_ICON | NIF_TIP,
            uCallbackMessage: WM_TRAY,
            hIcon: LoadIconW(GetModuleHandleW(core::ptr::null()), IMAGE_ICON_APP),
            szTip: [0; 128],
            dwState: 0,
            dwStateMask: 0,
            szInfo: [0; 256],
            uTimeout: 0,
            szInfoTitle: [0; 64],
            dwInfoFlags: 0,
            guidItem: GUID {
                Data1: 0,
                Data2: 0,
                Data3: 0,
                Data4: [0; 8],
            },
            hBalloonIcon: core::ptr::null_mut(),
        };
        let tip: Vec<u16> = "JinJuApp — 桌面金句".encode_utf16().take(127).collect();
        ni.szTip[..tip.len()].copy_from_slice(&tip);
        Shell_NotifyIconW(NIM_ADD, &mut ni) != 0
    }
}

/// Remove the tray icon (called on teardown and on explorer restart).
pub fn remove(hwnd: HWND) {
    unsafe {
        let mut ni = NOTIFYICONDATAW {
            cbSize: core::mem::size_of::<NOTIFYICONDATAW>() as u32,
            hWnd: hwnd,
            uID: 1,
            uFlags: 0,
            uCallbackMessage: 0,
            hIcon: core::ptr::null_mut(),
            szTip: [0; 128],
            dwState: 0,
            dwStateMask: 0,
            szInfo: [0; 256],
            uTimeout: 0,
            szInfoTitle: [0; 64],
            dwInfoFlags: 0,
            guidItem: GUID {
                Data1: 0,
                Data2: 0,
                Data3: 0,
                Data4: [0; 8],
            },
            hBalloonIcon: core::ptr::null_mut(),
        };
        Shell_NotifyIconW(NIM_DELETE, &mut ni);
    }
}

/// Show the context menu at the cursor; returns the chosen command id
/// (`0` if dismissed). `startup_on` sets the check state of the toggle.
pub fn show_menu(hwnd: HWND, startup_on: bool) -> usize {
    unsafe {
        let menu = CreatePopupMenu();
        if menu.is_null() {
            return 0;
        }
        AppendMenuW(
            menu,
            MF_STRING,
            CMD_OPEN_QUOTES,
            widez("编辑金句(&E)...").as_ptr(),
        );
        AppendMenuW(
            menu,
            MF_STRING,
            CMD_TOGGLE_STARTUP,
            widez("开机自启动(&S)").as_ptr(),
        );
        AppendMenuW(menu, MF_SEPARATOR, 0, core::ptr::null());
        AppendMenuW(menu, MF_STRING, CMD_EXIT, widez("退出(&X)").as_ptr());
        CheckMenuItem(
            menu,
            CMD_TOGGLE_STARTUP as u32,
            if startup_on { MF_CHECKED } else { 0 },
        );

        let mut pt = POINT { x: 0, y: 0 };
        GetCursorPos(&mut pt);
        // Classic recipe: the menu owner must be foreground, and a
        // WM_NULL must follow, or the menu may refuse to dismiss.
        SetForegroundWindow(hwnd);
        let cmd = TrackPopupMenu(
            menu,
            TPM_RIGHTBUTTON | TPM_BOTTOMALIGN | TPM_RETURNCMD,
            pt.x,
            pt.y,
            0,
            hwnd,
            core::ptr::null_mut(),
        );
        PostMessageW(hwnd, WM_NULL, 0, 0);
        DestroyMenu(menu);
        cmd as usize
    }
}

/// Open the quotes file with its associated editor.
pub fn open_quotes_file(path: &str) {
    unsafe {
        ShellExecuteW(
            core::ptr::null_mut(),
            core::ptr::null(), // default verb: open with the associated app
            widez(path).as_ptr(),
            core::ptr::null(),
            core::ptr::null(),
            SW_SHOWNOACTIVATE,
        );
    }
}
