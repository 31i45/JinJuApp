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

use geometry::{Edge, Rect, TaskbarAnchors};
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
/// Quote-change crossfade duration: the outgoing line fades out where it
/// stands while the incoming line fades in at rest.
const TRANSITION_MS: u64 = 400;
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

/// State of an in-flight quote crossfade.
struct Transition {
    /// The outgoing quote, fading out where it stands.
    prev_text: String,
    /// Its position and opacity at the quote change, captured before the
    /// scroller resets — the fade starts from exactly what was on screen,
    /// so a change landing mid-marquee-fade can't pop the old line back
    /// to full opacity.
    prev_x: i32,
    prev_alpha: f32,
    /// GetTickCount64 at the quote change.
    start: u64,
}

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
    last_presented_alpha: i32,
    /// In-flight quote crossfade; `None` = no transition.
    transition: Option<Transition>,
    exe_path: String,
    /// Current theme color (R, G, B) — re-queried from the system on
    /// every WM_SETTINGCHANGE so the strip stays visible under both
    /// light and dark Windows themes.
    current_color: (u8, u8, u8),
    /// True when a fullscreen application (game, video, browser) is
    /// active and the strip must stay hidden so it doesn't block the view.
    is_fullscreen: bool,
}

impl App {
    fn new(hwnd: HWND, dpi: u32, quotes: Source, exe_path: String) -> Option<App> {
        let win_w = scale_pt(WINDOW_W, dpi);
        let win_h = scale_pt(WINDOW_H, dpi);
        let current_color = render::current_theme_color();
        let surface = Surface::with_color(win_w, win_h, dpi, current_color)?;
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
            last_presented_alpha: -1,
            transition: None,
            exe_path,
            current_color,
            is_fullscreen: false,
        })
    }

    /// Pick the next quote, adapt the scroller, present.
    fn next_quote(&mut self) {
        let i = self.bag.next(&mut self.rng);
        let new_text = self.quotes.lines()[i].clone();
        // Crossfade: keep the outgoing line at its current spot for one
        // TRANSITION_MS envelope while the new line fades in at rest.
        // Position AND opacity must be captured before refresh() resets
        // the scroller — otherwise a change landing mid-marquee-fade
        // would pop the old line back to full opacity.
        if !self.text.is_empty() && self.text != new_text {
            self.transition = Some(Transition {
                prev_text: std::mem::replace(&mut self.text, new_text),
                prev_x: self.scroll.x(),
                prev_alpha: self.scroll.alpha(),
                start: unsafe { GetTickCount64() },
            });
        } else {
            self.text = new_text;
        }
        self.refresh();
    }

    /// Re-measure the current text and (re)start or stop the scroll timer.
    fn refresh(&mut self) {
        let text_w = self.surface.measure(&self.text);
        self.scroll.set_text(text_w);
        unsafe {
            // The scroll timer also drives an in-flight quote crossfade.
            if self.scroll.scrolling() || self.transition.is_some() {
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
        let alpha = self.scroll.alpha();
        // Crossfade envelope: the outgoing layer runs prev_alpha → 0, the
        // incoming layer 0 → alpha, over TRANSITION_MS.
        let mut t = 1.0_f32;
        if let Some(tr) = &self.transition {
            let now = unsafe { GetTickCount64() };
            t = crossfade_alpha(now - tr.start);
        }
        if t >= 1.0 {
            self.transition = None;
        }
        self.last_presented_x = x;
        self.last_presented_alpha = (alpha * t * 255.0).round() as i32;
        match &self.transition {
            Some(tr) => self.surface.draw(&[
                (tr.prev_text.as_str(), tr.prev_x, tr.prev_alpha * (1.0 - t)),
                (self.text.as_str(), x, alpha * t),
            ]),
            None => self.surface.draw(&[(self.text.as_str(), x, alpha * t)]),
        }
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
        let alpha_q = (self.scroll.alpha() * 255.0).round() as i32;
        if self.transition.is_some()
            || self.scroll.x() != self.last_presented_x
            || alpha_q != self.last_presented_alpha
        {
            self.present();
        }
        // Nothing left to animate: stop the 64 Hz cadence.
        if !self.scroll.scrolling() && self.transition.is_none() {
            unsafe { KillTimer(self.hwnd, SCROLL_TIMER) };
        }
    }

    /// Move to the strip adjacent to the taskbar, dynamically avoiding
    /// icon clusters by enumerating `Shell_TrayWnd` children. When a free
    /// band is found that's smaller than the default strip width, we shrink
    /// the window (Surface + Scroller) and request a crossfade reset so the
    /// text re-measures against the new width.
    fn reposition(&mut self) {
        let Some((taskbar, edge, monitor)) = desktop_geometry() else {
            return;
        };
        let anchors = taskbar_anchors(&taskbar, edge);
        let desired_w = scale_pt(WINDOW_W, self.dpi);
        let desired_h = scale_pt(WINDOW_H, self.dpi);
        let r = geometry::place_dynamic(
            &taskbar,
            &monitor,
            edge,
            desired_w,
            desired_h,
            PLACE_MARGIN,
            &anchors,
        );

        // Dynamic width — rebuild Surface + Scroller only if it actually changed.
        if r.width() != self.win_w {
            let new_dpi_sized_w = r.width();
            let new_h = r.height();
            let Some(surface) =
                Surface::with_color(new_dpi_sized_w, new_h, self.dpi, self.current_color)
            else {
                return;
            };
            self.surface = surface;
            self.win_w = new_dpi_sized_w;
            self.win_h = new_h;
            let speed = SCROLL_SPEED_LOGICAL * self.dpi as f32 / 96.0;
            self.scroll = Scroll::new(self.win_w, 0, speed);
            self.transition = None;
            self.refresh(); // re-measure text at the new width and redraw
        }

        unsafe {
            MoveWindow(self.hwnd, r.left, r.top, self.win_w, self.win_h, 0);
        }
    }

    /// System theme (dark/light) may have flipped — re-query the button
    /// text color, re-tint the surface, and redraw the current frame.
    fn on_theme_change(&mut self) {
        let new_color = render::current_theme_color();
        if new_color != self.current_color {
            self.current_color = new_color;
            self.surface.set_color(new_color);
            self.present();
        }
    }

    /// DPI changed (window moved between monitors): rebuild at new size.
    fn on_dpi_changed(&mut self, dpi: u32, suggested: &RECT) {
        let Some(surface) = Surface::with_color(
            scale_pt(WINDOW_W, dpi),
            scale_pt(WINDOW_H, dpi),
            dpi,
            self.current_color,
        ) else {
            return;
        };
        self.surface = surface;
        self.dpi = dpi;
        self.win_w = scale_pt(WINDOW_W, dpi);
        self.win_h = scale_pt(WINDOW_H, dpi);
        // The scroller caches view width and speed in px — both are
        // DPI-dependent, so rebuild it (refresh() re-measures the text).
        self.scroll = Scroll::new(self.win_w, 0, SCROLL_SPEED_LOGICAL * dpi as f32 / 96.0);
        // An in-flight crossfade's cached geometry belongs to the old DPI.
        self.transition = None;
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

/// Crossfade envelope: 0 → 1 linearly over TRANSITION_MS, clamped.
fn crossfade_alpha(elapsed_ms: u64) -> f32 {
    (elapsed_ms as f32 / TRANSITION_MS as f32).min(1.0)
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

/// Discover `MSTaskListWClass` (icon cluster) and `TrayNotifyWnd` (system
/// tray) inside `Shell_TrayWnd`. Returns `TaskbarAnchors`; missing fields
/// become `None` (caller falls back to centered placement).
fn taskbar_anchors(taskbar: &Rect, edge: Edge) -> TaskbarAnchors {
    if edge != Edge::Bottom {
        return TaskbarAnchors::default();
    }

    unsafe {
        let tray_class = "Shell_TrayWnd";
        let wide: Vec<u16> = tray_class.encode_utf16().chain([0]).collect();
        let h_tray = FindWindowW(wide.as_ptr(), core::ptr::null());
        if h_tray.is_null() {
            return TaskbarAnchors::default();
        }

        let mut collector = AnchorCollector::default();
        let ptr = &mut collector as *mut AnchorCollector as LPARAM;
        EnumChildWindows(h_tray, Some(anchor_cb), ptr);
        collector.filter_to_taskbar(taskbar);
        collector.into_anchors()
    }
}

/// Collects icon cluster + tray rects during `EnumChildWindows`.
#[derive(Default)]
struct AnchorCollector {
    icons: Option<Rect>,
    tray: Option<Rect>,
}

impl AnchorCollector {
    fn filter_to_taskbar(&mut self, tb: &Rect) {
        self.icons = self
            .icons
            .filter(|r| r.right > tb.left && r.left < tb.right && r.width() > 10 && r.height() > 0);
        self.tray = self
            .tray
            .filter(|r| r.right > tb.left && r.left < tb.right && r.width() > 10 && r.height() > 0);
    }

    fn into_anchors(self) -> TaskbarAnchors {
        TaskbarAnchors {
            icons: self.icons,
            tray: self.tray,
        }
    }
}

unsafe extern "system" fn anchor_cb(hwnd: HWND, lparam: LPARAM) -> BOOL {
    unsafe {
        let mut class_buf = [0u16; 128];
        let class_len = GetClassNameW(hwnd, class_buf.as_mut_ptr(), class_buf.len() as i32);
        if class_len <= 0 {
            return 1;
        }
        let class = String::from_utf16_lossy(&class_buf[..class_len as usize]);

        // Icon container class names — multiple exist across Windows versions
        // and Win11 sometimes marks MSTaskListWClass as invisible while the
        // actual visible content lives in ReBarWindow32 / MSTaskSwWClass.
        let is_icons = matches!(
            class.as_str(),
            "ReBarWindow32" | "MSTaskSwWClass" | "MSTaskListWClass"
        );
        let is_tray = class == "TrayNotifyWnd";
        if !is_icons && !is_tray {
            return 1;
        }

        let mut wr = RECT {
            left: 0,
            top: 0,
            right: 0,
            bottom: 0,
        };
        if GetWindowRect(hwnd, &mut wr) == 0 {
            return 1;
        }
        if wr.right <= wr.left || wr.bottom <= wr.top {
            return 1;
        }
        let p = lparam as *mut AnchorCollector;
        if is_icons {
            let new = Rect::from(wr);
            match (*p).icons {
                Some(existing) if new.width() > existing.width() => (*p).icons = Some(new),
                None => (*p).icons = Some(new),
                _ => {}
            }
        } else {
            (*p).tray = Some(Rect::from(wr));
        }
        1
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
// Fullscreen detection
// ---------------------------------------------------------------------------

/// Tolerance (in pixels) for treating two rectangles as "the same". Some
/// games report a window rect that is off-by-one from the monitor rect
/// (DPI scaling quirks, non-client borders), so a small slack avoids
/// false negatives.
const RECT_TOLERANCE_PX: i32 = 2;

/// Returns `true` when the foreground window covers the entire monitor that
/// the JinJu strip currently lives on — i.e. the user is watching a movie,
/// playing a game, or has a browser in true fullscreen.
///
/// Uses two independent signals so we catch both classic Win32 fullscreen
/// (game, media player) and modern fullscreen (Chrome F11, video overlay):
///
///   1. `ShQueryUserNotificationState` — shell-level notification of
///      `QUNS_FULL_SCREEN` or `QUNS_BUSY` (presentation, gaming, focus
///      assist). This covers both exclusive fullscreen and borderless
///      fullscreen apps that have told the shell they're in immersive mode.
///
///   2. Foreground-window-vs-monitor rect comparison — falls back for apps
///      that don't advertise to the shell.
///
/// `strip_hwnd` is the JinJu window; we ask for the monitor it's on so
/// multi-monitor setups only hide the strip that actually lives on the
/// monitor being covered.
unsafe fn is_fullscreen_active(strip_hwnd: HWND) -> bool {
    unsafe {
    // --- Signal 1: shell notification state ---
    // Dynamically loaded because older Windows SDKs don't ship the
    // static import for this function. Falls back to Signal 2 when
    // the function isn't available.
    //
    // IMPORTANT: Signal 1 tells us "some app is fullscreening *any*
    // display in the system". We still need Signal 2 to confirm that
    // it's the *same* display the JinJu strip lives on — otherwise a
    // fullscreen movie on monitor 1 would hide the strip on monitor 2
    // too, which would defeat AC-5 (multi-monitor isolation).
    let shell_says_fullscreen = if let Some(func) = sh_query_notification_state() {
        let state = func();
        state == QUNS_FULL_SCREEN || state == QUNS_BUSY
    } else {
        false
    };

    // --- Signal 2: foreground-window rect vs. monitor rect ---
    // Get the monitor the strip currently occupies.
    let monitor = MonitorFromWindow(strip_hwnd, MONITOR_DEFAULTTONEAREA);
    if monitor.is_null() {
        return false;
    }
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
        return false;
    }

    // Foreground window (walk up to root in case of a child or tool
    // window claiming foreground).
    let fg = GetForegroundWindow();
    if fg.is_null() || fg == strip_hwnd {
        return false;
    }
    let root = GetAncestor(fg, GA_ROOT);
    if root.is_null() || root == strip_hwnd {
        return false;
    }
    let mut fg_rect = RECT {
        left: 0,
        top: 0,
        right: 0,
        bottom: 0,
    };
    if GetWindowRect(root, &mut fg_rect) == 0 {
        return false;
    }

    // Partition: is the foreground window even *on the same monitor*
    // as the strip? If it lives on a different display then neither
    // signal should cause the strip to hide here — that display's
    // immersive mode doesn't concern us (AC-5 multi-monitor isolation).
    let fg_monitor = MonitorFromWindow(root, MONITOR_DEFAULTTONEAREA);
    let same_monitor = !fg_monitor.is_null() && fg_monitor == monitor;

    let rect_fits_fullscreen =
        rect_approximately_equal(&fg_rect, &mi.rcMonitor, RECT_TOLERANCE_PX);

    // Decision table — both signals must agree that this monitor is
    // the one being covered:
    //
    //   Signal 2 (rect fits)  Signal 1 (shell says)  same_monitor  Result
    //   --------------------  --------------------  ------------  ------
    //   true                  *                    true          hide ✅
    //   false                 true                 true          hide ✅
    //                                 (fallback: borderless
    //                                  fullscreen missed by
    //                                  rect comp but shell saw)
    //   *                     *                    false         show ✅
    //                                 (other display's
    //                                  fullscreen — not ours)
    //   false                 false                true          show ✅
    //                                 (normal window on our
    //                                  display — strip stays)
    same_monitor && (rect_fits_fullscreen || shell_says_fullscreen)
    }
}

/// `true` when `a` and `b` differ by at most `tolerance` pixels on each side.
/// Exposed for unit tests (the pure half of the function).
fn rect_approximately_equal(a: &RECT, b: &RECT, tolerance: i32) -> bool {
    let d_left = (a.left - b.left).abs();
    let d_top = (a.top - b.top).abs();
    let d_right = (a.right - b.right).abs();
    let d_bottom = (a.bottom - b.bottom).abs();
    d_left <= tolerance && d_top <= tolerance && d_right <= tolerance && d_bottom <= tolerance
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
                        // BUT: never resurrect the strip while a fullscreen
                        // app is active — that's the caller's job (see
                        // Z_ORDER_TIMER), not the quote tick's.
                        if !app.is_fullscreen && IsWindowVisible(hwnd) == 0 {
                            ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                        }
                        app.next_quote();
                    }
                    SCROLL_TIMER => app.tick_scroll(),
                    Z_ORDER_TIMER => {
                        // Fullscreen auto-hide check — runs every 500 ms so
                        // the strip disappears within half a second of any
                        // fullscreen app (Chrome F11, game, video player,
                        // PowerPoint slide show).
                        let fullscreen_now = is_fullscreen_active(hwnd);
                        if fullscreen_now != app.is_fullscreen {
                            if fullscreen_now {
                                ShowWindow(hwnd, SW_HIDE);
                            } else {
                                ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                                app.reposition();
                            }
                            app.is_fullscreen = fullscreen_now;
                        }

                        if app.is_fullscreen {
                            // Strip is hidden — don't bother re-asserting
                            // Z-order or repositioning. Skip both so we do
                            // one lightweight call instead of three.
                        } else {
                            // NOACTIVATE: never steal focus. No SHOWWINDOW:
                            // never fight Win+D (visibility is owned by
                            // the quote tick).
                            SetWindowPos(
                                hwnd,
                                HWND_TOPMOST,
                                0,
                                0,
                                0,
                                0,
                                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE,
                            );
                            // Cheap re-anchor every 500ms — catches taskbar
                            // alignment changes (centered ↔ left) and dock-
                            // edge moves that WM_SETTINGCHANGE misses. No
                            // render cost: reposition only rebuilds Surface
                            // if width changed, and EnumChildWindows is pure
                            // Win32.
                            app.reposition();
                        }
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
                app.on_theme_change();
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
                "{reason}\n\n本次运行使用内置金句。修正该文件后重启应用即可。\n文件格式：每行一句"
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
    fn crossfade_alpha_ramps_linearly_and_clamps() {
        assert_eq!(crossfade_alpha(0), 0.0);
        assert_eq!(crossfade_alpha(TRANSITION_MS / 2), 0.5);
        assert_eq!(crossfade_alpha(TRANSITION_MS), 1.0);
        assert_eq!(crossfade_alpha(u64::MAX), 1.0);
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

    #[test]
    fn rect_equal_when_identical() {
        let a = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        let b = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        assert!(rect_approximately_equal(&a, &b, RECT_TOLERANCE_PX));
    }

    #[test]
    fn rect_equal_within_tolerance() {
        let a = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        // 1 px off on every side — still within tolerance
        let b = RECT {
            left: -1,
            top: 1,
            right: 1921,
            bottom: 1079,
        };
        assert!(rect_approximately_equal(&a, &b, RECT_TOLERANCE_PX));
    }

    #[test]
    fn rect_not_equal_outside_tolerance() {
        let a = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        // 50 px off — clearly not fullscreen
        let b = RECT {
            left: 100,
            top: 50,
            right: 1600,
            bottom: 900,
        };
        assert!(!rect_approximately_equal(&a, &b, RECT_TOLERANCE_PX));
    }

    #[test]
    fn rect_not_equal_single_side_far_outside_tolerance() {
        let a = RECT {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        // Only 1 side is outside tolerance but by a lot — should still fail
        let b = RECT {
            left: 3,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        assert!(!rect_approximately_equal(&a, &b, RECT_TOLERANCE_PX));
    }
}
