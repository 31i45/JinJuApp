//! JinJuApp — a taskbar-side quote ticker for Windows.
//!
//! Pure-Win32 rewrite of the original PyQt application: no bootloader, no
//! packer, no framework — a native binary that renders one line of text
//! through a per-pixel-alpha layered window. All decision logic lives in
//! pure, unit-tested modules (`quotes`, `geometry`, `layout`, `render`);
//! this file is the thin host that wires them to the message loop.

#![windows_subsystem = "windows"]

mod geometry;
mod layout;
mod quotes;
mod render;
mod startup;
mod tray;
mod win32;

use geometry::{Edge, Rect};
use layout::Scroll;
use quotes::{Bag, Rng, Source};
use render::Surface;
use win32::*;

// ---------------------------------------------------------------------------
// Tunables (logical px @ 96 DPI unless noted)
// ---------------------------------------------------------------------------

const WINDOW_CLASS: &str = "JinJuApp";
const MUTEX_NAME: &str = "JinJuApp-SingleInstance";
const WINDOW_W: i32 = 500;
const WINDOW_H: i32 = 45;
const PLACE_MARGIN: i32 = 10;
const QUOTE_INTERVAL_MS: u32 = 3 * 60 * 1000; // every 3 minutes, as before
const QUOTE_TIMER: usize = 1;
const SCROLL_TIMER: usize = 2;
const SCROLL_TICK_MS: u32 = 15;
/// Scroll speed in logical px per second @ 96 DPI (scaled to the DPI at
/// which the surface renders).
const SCROLL_SPEED_LOGICAL: f32 = 45.0;
const Z_ORDER_TIMER: usize = 3;
/// The strip shares the taskbar band with explorer, which re-asserts its
/// own topmost status whenever the taskbar is activated (icon clicks,
/// new buttons, flyouts). Diagnostics showed a demoted strip stayed
/// hidden until the next quote change (minutes). Re-assert at this
/// cadence so any cover is bounded to one interval.
const Z_ORDER_INTERVAL_MS: u32 = 500;

/// Message id the shell broadcasts when the taskbar is (re)created.
static TASKBAR_CREATED: std::sync::OnceLock<u32> = std::sync::OnceLock::new();

// ---------------------------------------------------------------------------
// Command line (pure, unit-tested)
// ---------------------------------------------------------------------------

/// `--quiet`: started by autostart — degrade silently instead of popping
/// a message box in the user's face at logon.
#[derive(Debug, Default, PartialEq, Eq)]
struct Cli {
    quiet: bool,
}

fn parse_cli(args: &[String]) -> Cli {
    Cli {
        quiet: args.iter().any(|a| a == "--quiet" || a == "-q"),
    }
}

// ---------------------------------------------------------------------------
// Application state (owned by the window via GWLP_USERDATA)
// ---------------------------------------------------------------------------

struct App {
    hwnd: HWND,
    dpi: u32,
    win_w: i32,
    win_h: i32,
    surface: Surface,
    scroll: Scroll,
    text: String,
    quotes: Source,
    bag: Bag,
    rng: Rng,
    last_tick_ms: u64,
    last_presented_x: i32,
    exe_path: String,
}

impl App {
    fn new(hwnd: HWND, dpi: u32, quotes: Source, exe_path: String) -> Option<App> {
        let win_w = scale_pt(WINDOW_W, dpi);
        let win_h = scale_pt(WINDOW_H, dpi);
        let surface = Surface::new(win_w, win_h, dpi)?;
        let bag = Bag::new(quotes.lines().len());
        let speed_px_per_s = SCROLL_SPEED_LOGICAL * dpi as f32 / 96.0;
        Some(App {
            hwnd,
            dpi,
            win_w,
            win_h,
            surface,
            scroll: Scroll::new(win_w, 0, speed_px_per_s),
            text: String::new(),
            quotes,
            bag,
            rng: Rng::from_entropy(),
            last_tick_ms: 0,
            last_presented_x: i32::MIN,
            exe_path,
        })
    }

    /// Pick the next quote, adapt the scroller, present.
    fn next_quote(&mut self) {
        let i = self.bag.next(&mut self.rng);
        self.text = self.quotes.lines()[i].clone();
        self.refresh();
    }

    /// Re-measure the current text and (re)start or stop the scroll timer.
    fn refresh(&mut self) {
        let text_w = self.surface.measure(&self.text);
        self.scroll.set_text(text_w);
        unsafe {
            if self.scroll.scrolling() {
                SetTimer(
                    self.hwnd,
                    SCROLL_TIMER,
                    SCROLL_TICK_MS,
                    core::ptr::null_mut(),
                );
            } else {
                KillTimer(self.hwnd, SCROLL_TIMER);
            }
            self.last_tick_ms = GetTickCount64();
        }
        self.present();
    }

    /// Composite the text at the current scroll offset onto the screen.
    fn present(&mut self) {
        let x = self.scroll.x();
        self.last_presented_x = x;
        self.surface.draw(&self.text, x);
        unsafe {
            // Frame pacing: submit on the compositor's boundary instead of
            // mid-scan, so marquee frames don't land at arbitrary offsets
            // within the DWM refresh. Best-effort — on failure we simply
            // keep the previous pacing.
            let _ = DwmFlush();
            let size = POINT {
                x: self.win_w,
                y: self.win_h,
            };
            let src = POINT { x: 0, y: 0 };
            let blend = BLENDFUNCTION {
                BlendOp: 0, // AC_SRC_OVER
                BlendFlags: 0,
                SourceConstantAlpha: 255,
                AlphaFormat: AC_SRC_ALPHA,
            };
            UpdateLayeredWindow(
                self.hwnd,
                core::ptr::null_mut(),
                core::ptr::null_mut(),
                &size,
                self.surface.hdc(),
                &src,
                0,
                &blend,
                ULW_ALPHA,
            );
        }
    }

    /// Advance the marquee by the time since the last tick.
    ///
    /// The 15 ms cadence (~64 Hz, near display refresh) keeps the per-frame
    /// step under 1 px, so rounding produces a steady 0/1 px pattern instead
    /// of the 2/3 px alternation that reads as stutter at 20 Hz.
    fn tick_scroll(&mut self) {
        let now = unsafe { GetTickCount64() };
        // Clamp: after a sleep/suspend the delta is meaningless.
        let dt = (now - self.last_tick_ms).min(250) as f32;
        self.last_tick_ms = now;
        self.scroll.tick(dt);
        if self.scroll.x() != self.last_presented_x {
            self.present();
        }
    }

    /// Move to the strip adjacent to the taskbar (any docked edge).
    fn reposition(&mut self) {
        let Some((taskbar, edge, monitor)) = desktop_geometry() else {
            return;
        };
        let r = geometry::place(
            &taskbar,
            &monitor,
            edge,
            self.win_w,
            self.win_h,
            PLACE_MARGIN,
        );
        unsafe {
            MoveWindow(self.hwnd, r.left, r.top, self.win_w, self.win_h, 0);
        }
    }

    /// DPI changed (window moved between monitors): rebuild at new size.
    fn on_dpi_changed(&mut self, dpi: u32, suggested: &RECT) {
        let Some(surface) = Surface::new(scale_pt(WINDOW_W, dpi), scale_pt(WINDOW_H, dpi), dpi)
        else {
            return;
        };
        self.surface = surface;
        self.dpi = dpi;
        self.win_w = scale_pt(WINDOW_W, dpi);
        self.win_h = scale_pt(WINDOW_H, dpi);
        // The scroller caches view width and speed in px — both are
        // DPI-dependent, so rebuild it (refresh() re-measures the text).
        self.scroll = Scroll::new(self.win_w, 0, SCROLL_SPEED_LOGICAL * dpi as f32 / 96.0);
        unsafe {
            SetWindowPos(
                self.hwnd,
                core::ptr::null_mut(),
                suggested.left,
                suggested.top,
                suggested.right - suggested.left,
                suggested.bottom - suggested.top,
                SWP_NOACTIVATE, // size already correct; follow the suggestion
            );
        }
        self.refresh();
    }

    fn quotes_file_path(&self) -> String {
        let dir = std::path::Path::new(&self.exe_path)
            .parent()
            .unwrap_or(std::path::Path::new("."));
        dir.join(quotes::FILE_NAME).to_string_lossy().into_owned()
    }
}

/// Logical px @ 96 DPI → physical px at `dpi`.
fn scale_pt(v: i32, dpi: u32) -> i32 {
    ((i64::from(v) * i64::from(dpi) + 95) / 96) as i32
}

/// Taskbar rect + docked edge + primary monitor rect.
fn desktop_geometry() -> Option<(Rect, Edge, Rect)> {
    unsafe {
        let monitor = MonitorFromPoint(POINT { x: 0, y: 0 }, MONITOR_DEFAULTTOPRIMARY);
        let mut mi = MONITORINFO {
            cbSize: core::mem::size_of::<MONITORINFO>() as u32,
            rcMonitor: RECT {
                left: 0,
                top: 0,
                right: 0,
                bottom: 0,
            },
            rcWork: RECT {
                left: 0,
                top: 0,
                right: 0,
                bottom: 0,
            },
            dwFlags: 0,
        };
        if GetMonitorInfoW(monitor, &mut mi) == 0 {
            return None;
        }
        let mon = Rect::from(mi.rcMonitor);
        let mut abd = APPBARDATA {
            cbSize: core::mem::size_of::<APPBARDATA>() as u32,
            hWnd: core::ptr::null_mut(),
            uCallbackMessage: 0,
            uEdge: 0,
            rc: RECT {
                left: 0,
                top: 0,
                right: 0,
                bottom: 0,
            },
            lParam: 0,
        };
        if SHAppBarMessage(ABM_GETTASKBARPOS, &mut abd) == 0 {
            // No appbar answer: assume a bottom taskbar of typical height.
            let tb = Rect::new(
                mon.left,
                mon.bottom - scale_pt(48, 96),
                mon.width(),
                scale_pt(48, 96),
            );
            return Some((tb, Edge::Bottom, mon));
        }
        Some((Rect::from(abd.rc), Edge::from_abe(abd.uEdge), mon))
    }
}

impl From<RECT> for Rect {
    fn from(r: RECT) -> Rect {
        Rect {
            left: r.left,
            top: r.top,
            right: r.right,
            bottom: r.bottom,
        }
    }
}

// ---------------------------------------------------------------------------
// Window procedure
// ---------------------------------------------------------------------------

unsafe fn app_of(hwnd: HWND) -> Option<&'static mut App> {
    unsafe {
        let p = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut App;
        if p.is_null() { None } else { Some(&mut *p) }
    }
}

unsafe extern "system" fn wnd_proc(hwnd: HWND, msg: UINT, wp: WPARAM, lp: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_NCCREATE => {
                // Take ownership of the App box passed through CreateWindowExW.
                let cs = lp as *const CREATESTRUCTW;
                let app = (*cs).lpCreateParams as *mut App;
                if !app.is_null() {
                    (*app).hwnd = hwnd;
                    SetWindowLongPtrW(hwnd, GWLP_USERDATA, app as isize);
                }
                DefWindowProcW(hwnd, msg, wp, lp)
            }
            WM_TIMER => {
                let Some(app) = app_of(hwnd) else { return 0 };
                match wp {
                    QUOTE_TIMER => {
                        // Gentle recovery: if hidden (e.g. Win+D), re-show at
                        // the natural quote-change moment — never fight the
                        // user at 1 Hz like the original watchdog did.
                        if IsWindowVisible(hwnd) == 0 {
                            ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                        }
                        app.next_quote();
                    }
                    SCROLL_TIMER => app.tick_scroll(),
                    Z_ORDER_TIMER => {
                        // NOACTIVATE: never steal focus. No SHOWWINDOW: never
                        // fight Win+D (visibility is owned by the quote tick).
                        SetWindowPos(
                            hwnd,
                            HWND_TOPMOST,
                            0,
                            0,
                            0,
                            0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                        );
                    }
                    _ => {}
                }
                0
            }
            WM_DPICHANGED => {
                let Some(app) = app_of(hwnd) else { return 0 };
                let dpi = ((wp >> 16) & 0xFFFF) as u32;
                let suggested = lp as *const RECT;
                if !suggested.is_null() {
                    app.on_dpi_changed(dpi, &*suggested);
                }
                0
            }
            WM_DISPLAYCHANGE | WM_SETTINGCHANGE => {
                let Some(app) = app_of(hwnd) else { return 0 };
                app.reposition();
                0
            }
            WM_DESTROY => {
                tray::remove(hwnd);
                // Drop the App box now, before the window handle dies.
                let p = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut App;
                if !p.is_null() {
                    drop(Box::from_raw(p));
                    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
                }
                PostQuitMessage(0);
                0
            }
            _ => {
                if TASKBAR_CREATED.get() == Some(&msg) {
                    // Explorer restarted: the tray icon is gone, re-add it.
                    tray::add(hwnd);
                    return 0;
                }
                if msg == tray::WM_TRAY {
                    let Some(app) = app_of(hwnd) else { return 0 };
                    let event = lp as u32;
                    if event == WM_RBUTTONUP {
                        let on = startup::is_enabled(&app.exe_path);
                        match tray::show_menu(hwnd, on) {
                            tray::CMD_OPEN_QUOTES => {
                                tray::open_quotes_file(&app.quotes_file_path());
                            }
                            tray::CMD_TOGGLE_STARTUP => {
                                let result = if on {
                                    startup::disable()
                                } else {
                                    startup::enable(&app.exe_path)
                                };
                                if let Err(e) = result {
                                    alert(hwnd, &format!("切换开机自启动失败：{e}"));
                                }
                            }
                            tray::CMD_EXIT => {
                                DestroyWindow(hwnd);
                            }
                            _ => {}
                        }
                    } else if event == WM_LBUTTONDBLCLK {
                        tray::open_quotes_file(&app.quotes_file_path());
                    }
                    return 0;
                }
                DefWindowProcW(hwnd, msg, wp, lp)
            }
        }
    }
}

fn alert(hwnd: HWND, text: &str) {
    unsafe {
        let msg: Vec<u16> = text.encode_utf16().chain([0]).collect();
        let caption: Vec<u16> = "JinJuApp".encode_utf16().chain([0]).collect();
        MessageBoxW(
            hwnd,
            msg.as_ptr(),
            caption.as_ptr(),
            MB_OK | MB_ICONWARNING | MB_SETFOREGROUND,
        );
    }
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

fn widez(s: &str) -> Vec<u16> {
    s.encode_utf16().chain([0]).collect()
}

fn exe_path() -> String {
    unsafe {
        let mut buf = [0u16; 1024];
        let len = GetModuleFileNameW(core::ptr::null_mut(), buf.as_mut_ptr(), 1024) as usize;
        String::from_utf16_lossy(&buf[..len])
    }
}

fn primary_dpi() -> u32 {
    unsafe {
        let hdc = GetDC(core::ptr::null_mut());
        let dpi = if hdc.is_null() {
            96
        } else {
            GetDeviceCaps(hdc, LOGPIXELSX).max(96) as u32
        };
        ReleaseDC(core::ptr::null_mut(), hdc);
        dpi
    }
}

fn main() {
    let cli = parse_cli(&std::env::args().skip(1).collect::<Vec<_>>());

    // Per-monitor v2: crisp text at any display scale (best-effort on
    // older systems; the call simply fails and we stay system-default).
    unsafe {
        SetProcessDpiAwarenessContext(dpi_awareness_per_monitor_v2());
    }

    // Single instance.
    unsafe {
        let name = widez(MUTEX_NAME);
        let mutex = CreateMutexW(core::ptr::null_mut(), 0, name.as_ptr());
        if !mutex.is_null() && GetLastError() == ERROR_ALREADY_EXISTS {
            return; // already running — silently defer to the first instance
        }
        let _ = mutex; // held until process exit (raw pointer, no drop)
    }

    let exe = exe_path();
    let exe_dir = std::path::Path::new(&exe)
        .parent()
        .map(std::path::Path::to_path_buf)
        .unwrap_or_else(|| std::path::PathBuf::from("."));
    let quotes = quotes::load(&exe_dir);
    if let Some(reason) = quotes.note().filter(|_| !cli.quiet) {
        alert(
            core::ptr::null_mut(),
            &format!(
                "{reason}\n\n本次运行使用内置金句。修正该文件后重启应用即可。\n文件格式：[{{\"quote\": \"…\"}}, …]"
            ),
        );
    }

    unsafe {
        let hinst = GetModuleHandleW(core::ptr::null_mut());
        let class_name = widez(WINDOW_CLASS);

        let wc = WNDCLASSEXW {
            cbSize: core::mem::size_of::<WNDCLASSEXW>() as u32,
            style: 0,
            lpfnWndProc: Some(wnd_proc),
            cbClsExtra: 0,
            cbWndExtra: 0,
            hInstance: hinst,
            hIcon: LoadIconW(hinst, IMAGE_ICON_APP),
            hCursor: LoadCursorW(core::ptr::null_mut(), IDC_ARROW),
            hbrBackground: core::ptr::null_mut(),
            lpszMenuName: core::ptr::null(),
            lpszClassName: class_name.as_ptr(),
            hIconSm: LoadIconW(hinst, IMAGE_ICON_APP),
        };
        if RegisterClassExW(&wc) == 0 {
            alert(core::ptr::null_mut(), "窗口类注册失败");
            return;
        }

        TASKBAR_CREATED.get_or_init(|| RegisterWindowMessageW(widez("TaskbarCreated").as_ptr()));

        let dpi = primary_dpi();
        let app = match App::new(core::ptr::null_mut(), dpi, quotes, exe.clone()) {
            Some(a) => a,
            None => {
                alert(core::ptr::null_mut(), "初始化渲染表面失败");
                return;
            }
        };
        let app = Box::into_raw(Box::new(app));

        let ex_style =
            WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE;
        let hwnd = CreateWindowExW(
            ex_style,
            class_name.as_ptr(),
            widez("JinJuApp").as_ptr(),
            WS_POPUP,
            // Position is finalized by reposition(); pass CW_USEDEFAULT.
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            scale_pt(WINDOW_W, dpi),
            scale_pt(WINDOW_H, dpi),
            core::ptr::null_mut(),
            core::ptr::null_mut(),
            hinst,
            app as *mut core::ffi::c_void,
        );
        if hwnd.is_null() {
            drop(Box::from_raw(app));
            alert(core::ptr::null_mut(), "窗口创建失败");
            return;
        }

        {
            let app = &mut *app;
            app.reposition();
            app.next_quote();
            SetTimer(hwnd, QUOTE_TIMER, QUOTE_INTERVAL_MS, core::ptr::null_mut());
            SetTimer(
                hwnd,
                Z_ORDER_TIMER,
                Z_ORDER_INTERVAL_MS,
                core::ptr::null_mut(),
            );
        }

        if !tray::add(hwnd) {
            alert(
                hwnd,
                "无法创建托盘图标：请通过任务管理器结束 JinJuApp.exe 退出。",
            );
        }
        ShowWindow(hwnd, SW_SHOWNOACTIVATE);

        let mut msg: MSG = core::mem::zeroed();
        loop {
            let r = GetMessageW(&mut msg, core::ptr::null_mut(), 0, 0);
            if r <= 0 {
                // 0 = WM_QUIT, -1 = error: either way, stop.
                break;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
}

// ---------------------------------------------------------------------------
// Tests (pure halves)
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cli_defaults_to_interactive() {
        assert_eq!(parse_cli(&[]), Cli { quiet: false });
    }

    #[test]
    fn cli_accepts_quiet_and_short_form() {
        assert_eq!(parse_cli(&["--quiet".to_string()]), Cli { quiet: true });
        assert_eq!(parse_cli(&["-q".to_string()]), Cli { quiet: true });
    }

    #[test]
    fn cli_ignores_other_arguments() {
        assert_eq!(
            parse_cli(&["--foo".to_string(), "bar".to_string()]),
            Cli { quiet: false }
        );
    }

    #[test]
    fn scale_pt_scales_by_dpi() {
        assert_eq!(scale_pt(500, 96), 500);
        assert_eq!(scale_pt(500, 192), 1000);
        assert_eq!(scale_pt(500, 144), 750);
        assert_eq!(scale_pt(45, 96), 45);
    }

    #[test]
    fn scale_pt_rounds_up_at_half() {
        // 45 * 1.25 = 56.25 → 57 (96→120 dpi)
        assert_eq!(scale_pt(45, 120), 57);
    }
}
