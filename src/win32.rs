//! Hand-written Win32 FFI declarations.
//!
//! Zero third-party bindings. Every system call this program makes is
//! declared here, in one place, visible at a glance — the complete API
//! surface, nothing hidden behind a binding layer.

#![allow(
    non_snake_case,
    non_camel_case_types,
    non_upper_case_globals,
    clippy::upper_case_acronyms
)]

use core::ffi::c_void;

// ---------------------------------------------------------------------------
// Base types
// ---------------------------------------------------------------------------

pub type HANDLE = *mut c_void;
pub type HWND = HANDLE;
pub type HDC = HANDLE;
pub type HICON = HANDLE;
pub type HCURSOR = HICON;
pub type HFONT = HANDLE;
pub type HBITMAP = HANDLE;
pub type HBRUSH = HANDLE;
pub type HMENU = HANDLE;
pub type HINSTANCE = HANDLE;
pub type HGDIOBJ = HANDLE;
pub type HKEY = *mut c_void;
pub type LPARAM = isize;
pub type WPARAM = usize;
pub type LRESULT = isize;
pub type ATOM = u16;
pub type BOOL = i32;
pub type DWORD = u32;
pub type UINT = u32;
pub type LONG = i32;
pub type WORD = u16;
pub type COLORREF = u32;

// ---------------------------------------------------------------------------
// Structures (field order and layout follow the Win32 ABI, repr(C))
// ---------------------------------------------------------------------------

#[repr(C)]
#[derive(Clone, Copy)]
pub struct RECT {
    pub left: LONG,
    pub top: LONG,
    pub right: LONG,
    pub bottom: LONG,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct POINT {
    pub x: LONG,
    pub y: LONG,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct MSG {
    pub hwnd: HWND,
    pub message: UINT,
    pub wParam: WPARAM,
    pub lParam: LPARAM,
    pub time: DWORD,
    pub pt: POINT,
    pub lPrivate: DWORD,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct WNDCLASSEXW {
    pub cbSize: UINT,
    pub style: UINT,
    pub lpfnWndProc: Option<unsafe extern "system" fn(HWND, UINT, WPARAM, LPARAM) -> LRESULT>,
    pub cbClsExtra: i32,
    pub cbWndExtra: i32,
    pub hInstance: HINSTANCE,
    pub hIcon: HICON,
    pub hCursor: HCURSOR,
    pub hbrBackground: HBRUSH,
    pub lpszMenuName: *const u16,
    pub lpszClassName: *const u16,
    pub hIconSm: HICON,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct LOGFONTW {
    pub lfHeight: LONG,
    pub lfWidth: LONG,
    pub lfEscapement: LONG,
    pub lfOrientation: LONG,
    pub lfWeight: LONG,
    pub lfItalic: u8,
    pub lfUnderline: u8,
    pub lfStrikeOut: u8,
    pub lfCharSet: u8,
    pub lfOutPrecision: u8,
    pub lfClipPrecision: u8,
    pub lfQuality: u8,
    pub lfPitchAndFamily: u8,
    pub lfFaceName: [u16; 32],
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct NONCLIENTMETRICSW {
    pub cbSize: UINT,
    pub iBorderWidth: i32,
    pub iScrollWidth: i32,
    pub iScrollHeight: i32,
    pub iCaptionWidth: i32,
    pub iCaptionHeight: i32,
    pub lfCaptionFont: LOGFONTW,
    pub iSmCaptionWidth: i32,
    pub iSmCaptionHeight: i32,
    pub lfSmCaptionFont: LOGFONTW,
    pub iMenuWidth: i32,
    pub iMenuHeight: i32,
    pub lfMenuFont: LOGFONTW,
    pub iStatusWidth: i32,
    pub iStatusHeight: i32,
    pub lfStatusFont: LOGFONTW,
    pub iMessageWidth: i32,
    pub lfMessageFont: LOGFONTW,
    pub iPaddedBorderWidth: i32,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct MONITORINFO {
    pub cbSize: DWORD,
    pub rcMonitor: RECT,
    pub rcWork: RECT,
    pub dwFlags: DWORD,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct APPBARDATA {
    pub cbSize: DWORD,
    pub hWnd: HWND,
    pub uCallbackMessage: UINT,
    pub uEdge: UINT,
    pub rc: RECT,
    pub lParam: LPARAM,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct GUID {
    pub Data1: u32,
    pub Data2: u16,
    pub Data3: u16,
    pub Data4: [u8; 8],
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct NOTIFYICONDATAW {
    pub cbSize: DWORD,
    pub hWnd: HWND,
    pub uID: UINT,
    pub uFlags: UINT,
    pub uCallbackMessage: UINT,
    pub hIcon: HICON,
    pub szTip: [u16; 128],
    pub dwState: DWORD,
    pub dwStateMask: DWORD,
    pub szInfo: [u16; 256],
    pub uTimeout: UINT,
    pub szInfoTitle: [u16; 64],
    pub dwInfoFlags: DWORD,
    pub guidItem: GUID,
    pub hBalloonIcon: HICON,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct BITMAPINFOHEADER {
    pub biSize: DWORD,
    pub biWidth: LONG,
    pub biHeight: LONG,
    pub biPlanes: WORD,
    pub biBitCount: WORD,
    pub biCompression: DWORD,
    pub biSizeImage: DWORD,
    pub biXPelsPerMeter: LONG,
    pub biYPelsPerMeter: LONG,
    pub biClrUsed: DWORD,
    pub biClrImportant: DWORD,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct BLENDFUNCTION {
    pub BlendOp: u8,
    pub BlendFlags: u8,
    pub SourceConstantAlpha: u8,
    pub AlphaFormat: u8,
}

#[repr(C)]
pub struct CREATESTRUCTW {
    pub lpCreateParams: *mut c_void,
    pub hInstance: HINSTANCE,
    pub hMenu: HMENU,
    pub hWndParent: HWND,
    pub cy: i32,
    pub cx: i32,
    pub y: i32,
    pub x: i32,
    pub style: i32,
    pub lpszName: *const u16,
    pub lpszClass: *const u16,
    pub dwExStyle: u32,
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

// window styles
pub const WS_POPUP: u32 = 0x8000_0000;
pub const WS_EX_TOPMOST: u32 = 0x0000_0008;
pub const WS_EX_TRANSPARENT: u32 = 0x0000_0020;
pub const WS_EX_TOOLWINDOW: u32 = 0x0000_0080;
pub const WS_EX_LAYERED: u32 = 0x0008_0000;
pub const WS_EX_NOACTIVATE: u32 = 0x0800_0000;

// messages
pub const WM_NULL: u32 = 0x0000;
pub const WM_DESTROY: u32 = 0x0002;
pub const WM_SETTINGCHANGE: u32 = 0x001A;
pub const WM_DISPLAYCHANGE: u32 = 0x007E;
pub const WM_NCCREATE: u32 = 0x0081;
pub const WM_TIMER: u32 = 0x0113;
pub const WM_APP: u32 = 0x8000;
pub const WM_RBUTTONUP: u32 = 0x0205;
pub const WM_LBUTTONDBLCLK: u32 = 0x0203;
pub const WM_DPICHANGED: u32 = 0x02E0;

// show / positioning
pub const SW_SHOWNOACTIVATE: i32 = 4;
pub const HWND_TOPMOST: HWND = core::ptr::without_provenance_mut(usize::MAX);
pub const SWP_NOSIZE: u32 = 0x0001;
pub const SWP_NOMOVE: u32 = 0x0002;
pub const CW_USEDEFAULT: i32 = 0x8000_0000u32 as i32;

// layered window
pub const ULW_ALPHA: u32 = 0x0000_0002;
pub const AC_SRC_ALPHA: u8 = 0x01;
pub const DIB_RGB_COLORS: u32 = 0;
pub const BI_RGB: u32 = 0;

// GDI
pub const FW_LIGHT: i32 = 300;
pub const ANTIALIASED_QUALITY: u8 = 4;
pub const DT_SINGLELINE: u32 = 0x20;
pub const DT_VCENTER: u32 = 0x04;
pub const DT_NOPREFIX: u32 = 0x0800;
pub const DT_CALCRECT: u32 = 0x0400;
pub const BLACK_BRUSH: i32 = 4;

// notifications area
pub const NIM_ADD: u32 = 0;
pub const NIM_DELETE: u32 = 2;
pub const NIF_MESSAGE: u32 = 0x01;
pub const NIF_ICON: u32 = 0x02;
pub const NIF_TIP: u32 = 0x04;

// menu
pub const MF_STRING: u32 = 0x0000_0000;
pub const MF_SEPARATOR: u32 = 0x0000_0800;
pub const MF_CHECKED: u32 = 0x0000_0008;
pub const TPM_RIGHTBUTTON: u32 = 0x0000_0002;
pub const TPM_BOTTOMALIGN: u32 = 0x0000_0020;
pub const TPM_RETURNCMD: u32 = 0x0000_0100;

// appbar
pub const ABM_GETTASKBARPOS: u32 = 5;

// misc
pub const SPI_GETNONCLIENTMETRICS: u32 = 0x0029;
pub const LOGPIXELSX: i32 = 88;
pub const ERROR_ALREADY_EXISTS: u32 = 183;
pub const MONITOR_DEFAULTTOPRIMARY: u32 = 1;
pub const MB_OK: u32 = 0x0000_0000;
pub const MB_ICONWARNING: u32 = 0x0000_0030;
pub const MB_SETFOREGROUND: u32 = 0x0001_0000;
pub const IDC_ARROW: *const u16 = core::ptr::without_provenance(32512); // MAKEINTRESOURCE
pub const IMAGE_ICON_APP: *const u16 = core::ptr::without_provenance(1); // winres icon resource ID 1

// registry
pub const HKEY_CURRENT_USER: HKEY = 0x8000_0001usize as HKEY;
pub const KEY_READ: u32 = 0x0002_0019;
pub const KEY_SET_VALUE: u32 = 0x2002;
pub const REG_SZ: u32 = 1;

// ---------------------------------------------------------------------------
// kernel32
// ---------------------------------------------------------------------------

#[link(name = "kernel32")]
unsafe extern "system" {
    pub fn GetDC(hWnd: HWND) -> HDC;
    pub fn ReleaseDC(hWnd: HWND, hdc: HDC) -> i32;
    pub fn CreateMutexW(
        lpMutexAttributes: *mut c_void,
        bInitialOwner: BOOL,
        lpName: *const u16,
    ) -> HANDLE;
    pub fn GetLastError() -> u32;
    pub fn GetModuleFileNameW(hModule: HINSTANCE, lpFilename: *mut u16, nSize: u32) -> u32;
    pub fn GetModuleHandleW(lpModuleName: *const u16) -> HINSTANCE;
    pub fn GetTickCount64() -> u64;
    pub fn MulDiv(nNumber: i32, nNumerator: i32, nDenominator: i32) -> i32;
}

// ---------------------------------------------------------------------------
// user32
// ---------------------------------------------------------------------------

#[link(name = "user32")]
unsafe extern "system" {
    pub fn RegisterClassExW(lpWndClass: *const WNDCLASSEXW) -> ATOM;
    pub fn CreateWindowExW(
        dwExStyle: u32,
        lpClassName: *const u16,
        lpWindowName: *const u16,
        dwStyle: u32,
        x: i32,
        y: i32,
        nWidth: i32,
        nHeight: i32,
        hWndParent: HWND,
        hMenu: HMENU,
        hInstance: HINSTANCE,
        lpParam: *mut c_void,
    ) -> HWND;
    pub fn DefWindowProcW(hWnd: HWND, Msg: UINT, wParam: WPARAM, lParam: LPARAM) -> LRESULT;
    pub fn DestroyWindow(hWnd: HWND) -> BOOL;
    pub fn ShowWindow(hWnd: HWND, nCmdShow: i32) -> BOOL;
    pub fn MoveWindow(
        hWnd: HWND,
        x: i32,
        y: i32,
        nWidth: i32,
        nHeight: i32,
        bRepaint: BOOL,
    ) -> BOOL;
    pub fn UpdateLayeredWindow(
        hWnd: HWND,
        hdcDst: HDC,
        pptDst: *const POINT,
        psize: *const POINT,
        hdcSrc: HDC,
        pptSrc: *const POINT,
        crKey: COLORREF,
        pblend: *const BLENDFUNCTION,
        dwFlags: DWORD,
    ) -> BOOL;
    pub fn SetWindowPos(
        hWnd: HWND,
        hWndInsertAfter: HWND,
        x: i32,
        y: i32,
        cx: i32,
        cy: i32,
        uFlags: u32,
    ) -> BOOL;
    pub fn SetTimer(hWnd: HWND, nIDEvent: usize, uElapse: u32, lpTimerFunc: *mut c_void) -> usize;
    pub fn KillTimer(hWnd: HWND, uIDEvent: usize) -> BOOL;
    pub fn GetMessageW(lpMsg: *mut MSG, hWnd: HWND, wMsgFilterMin: u32, wMsgFilterMax: u32) -> i32;
    pub fn PostQuitMessage(nExitCode: i32);
    pub fn TranslateMessage(lpMsg: *const MSG) -> BOOL;
    pub fn DispatchMessageW(lpMsg: *const MSG) -> LRESULT;
    pub fn PostMessageW(hWnd: HWND, Msg: UINT, wParam: WPARAM, lParam: LPARAM) -> BOOL;
    pub fn SetForegroundWindow(hWnd: HWND) -> BOOL;
    pub fn SystemParametersInfoW(
        uiAction: u32,
        uiParam: u32,
        pvParam: *mut c_void,
        fWinIni: u32,
    ) -> BOOL;
    pub fn SetProcessDpiAwarenessContext(value: *mut c_void) -> BOOL;
    pub fn LoadCursorW(hInstance: HINSTANCE, lpCursorName: *const u16) -> HCURSOR;
    pub fn LoadIconW(hInstance: HINSTANCE, lpIconName: *const u16) -> HICON;
    pub fn RegisterWindowMessageW(lpString: *const u16) -> u32;
    pub fn CreatePopupMenu() -> HMENU;
    pub fn AppendMenuW(hMenu: HMENU, uFlags: u32, uIDNewItem: usize, lpNewItem: *const u16)
    -> BOOL;
    pub fn TrackPopupMenu(
        hMenu: HMENU,
        uFlags: u32,
        x: i32,
        y: i32,
        nReserved: i32,
        hWnd: HWND,
        prcRect: *mut RECT,
    ) -> BOOL;
    pub fn DestroyMenu(hMenu: HMENU) -> BOOL;
    pub fn CheckMenuItem(hMenu: HMENU, uIDCheckItem: u32, uCheck: u32) -> u32;
    pub fn GetCursorPos(lpPoint: *mut POINT) -> BOOL;
    pub fn SetWindowLongPtrW(hWnd: HWND, nIndex: i32, dwNewLong: isize) -> isize;
    pub fn GetWindowLongPtrW(hWnd: HWND, nIndex: i32) -> isize;
    pub fn MonitorFromPoint(pt: POINT, dwFlags: u32) -> HANDLE;
    pub fn GetMonitorInfoW(hMonitor: HANDLE, lpmi: *mut MONITORINFO) -> BOOL;
    pub fn MessageBoxW(hWnd: HWND, lpText: *const u16, lpCaption: *const u16, uType: u32) -> i32;
    pub fn IsWindowVisible(hWnd: HWND) -> BOOL;
    pub fn FillRect(hdc: HDC, lprc: *mut RECT, hbr: HBRUSH) -> i32;
}

pub const GWLP_USERDATA: i32 = -21;

/// Per-Monitor v2 DPI awareness context handle (Windows 10 1703+).
pub fn dpi_awareness_per_monitor_v2() -> *mut c_void {
    (-4isize) as *mut c_void
}

// ---------------------------------------------------------------------------
// gdi32
// ---------------------------------------------------------------------------

#[link(name = "gdi32")]
unsafe extern "system" {
    pub fn CreateCompatibleDC(hdc: HDC) -> HDC;
    pub fn DeleteDC(hdc: HDC) -> BOOL;
    pub fn GetDeviceCaps(hdc: HDC, index: i32) -> i32;
    pub fn SelectObject(hdc: HDC, h: HGDIOBJ) -> HGDIOBJ;
    pub fn DeleteObject(ho: HGDIOBJ) -> BOOL;
    pub fn GetStockObject(i: i32) -> HGDIOBJ;
    pub fn CreateDIBSection(
        hdc: HDC,
        pbmi: *const BITMAPINFOHEADER,
        usage: u32,
        ppvBits: *mut *mut c_void,
        hSection: HANDLE,
        dwOffset: u32,
    ) -> HBITMAP;
    pub fn CreateFontIndirectW(lplf: *const LOGFONTW) -> HFONT;
    pub fn SetBkColor(hdc: HDC, color: COLORREF) -> COLORREF;
    pub fn SetTextColor(hdc: HDC, color: COLORREF) -> COLORREF;
    pub fn DrawTextW(
        hdc: HDC,
        lpchText: *const u16,
        cchText: i32,
        lprc: *mut RECT,
        format: u32,
    ) -> i32;
}

/// Win32 `RGB` macro.
pub fn rgb(r: u8, g: u8, b: u8) -> COLORREF {
    u32::from(r) | (u32::from(g) << 8) | (u32::from(b) << 16)
}

// ---------------------------------------------------------------------------
// shell32
// ---------------------------------------------------------------------------

#[link(name = "shell32")]
unsafe extern "system" {
    pub fn Shell_NotifyIconW(dwMessage: u32, lpData: *mut NOTIFYICONDATAW) -> BOOL;
    pub fn ShellExecuteW(
        hwnd: HWND,
        lpVerb: *const u16,
        lpFile: *const u16,
        lpParameters: *const u16,
        lpDirectory: *const u16,
        nShow: i32,
    ) -> HANDLE;
    pub fn SHAppBarMessage(dwMessage: u32, pData: *mut APPBARDATA) -> usize;
}

// ---------------------------------------------------------------------------
// advapi32
// ---------------------------------------------------------------------------

#[link(name = "advapi32")]
unsafe extern "system" {
    pub fn RegOpenKeyExW(
        hKey: HKEY,
        lpSubKey: *const u16,
        ulOptions: u32,
        samDesired: u32,
        phkResult: *mut HKEY,
    ) -> i32;
    pub fn RegCreateKeyExW(
        hKey: HKEY,
        lpSubKey: *const u16,
        reserved: u32,
        lpClass: *const u16,
        dwOptions: u32,
        samDesired: u32,
        lpSecurityAttributes: *mut c_void,
        phkResult: *mut HKEY,
        lpdwDisposition: *mut u32,
    ) -> i32;
    pub fn RegQueryValueExW(
        hKey: HKEY,
        lpValueName: *const u16,
        lpReserved: *mut u32,
        lpType: *mut u32,
        lpData: *mut u8,
        lpcbData: *mut u32,
    ) -> i32;
    pub fn RegSetValueExW(
        hKey: HKEY,
        lpValueName: *const u16,
        reserved: u32,
        dwType: u32,
        lpData: *const u8,
        cbData: u32,
    ) -> i32;
    pub fn RegDeleteValueW(hKey: HKEY, lpValueName: *const u16) -> i32;
    pub fn RegCloseKey(hKey: HKEY) -> i32;
}
