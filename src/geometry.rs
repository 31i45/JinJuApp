//! Taskbar geometry — anchor-driven placement.
//!
//! Bottom taskbar placement has exactly two modes, distinguished by
//! whether the icon cluster (`MSTaskListWClass`) sits flush to the left
//! edge or has a gap between it and the edge:
//!
//! **居中模式** (`icons.left > taskbar.left`):
//!   Strip is at the **very left of the taskbar**, its right edge lands
//!   just before `icons.left`. Width = `icons.left - taskbar.left - 2*margin`.
//!   As centered icons grow wider, the strip's right edge shrinks leftward.
//!
//! **靠左模式** (`icons.left ≈ taskbar.left`):
//!   Strip sits **between the icon cluster and the tray**, anchored at
//!   `icons.right` and ending at `tray.left`. Width = `tray.left - icons.right - 2*margin`.
//!   The strip shrinks when either the icon cluster grows rightward or
//!   the tray grows leftward — both sides constrain it.
//!
//! Enumeration provides two rectangles; if either is missing we fall
//! back to centered-inside-the-taskbar. Top/side-docked taskbars are
//! simple: strip lives on the desktop side of the bar.

/// Which screen edge the taskbar is docked to.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Edge {
    Top,
    Right,
    Bottom,
    Left,
}

impl Edge {
    /// From a Win32 `ABE_*` constant.
    pub fn from_abe(abe: u32) -> Edge {
        match abe {
            1 => Edge::Top,
            2 => Edge::Right,
            3 => Edge::Bottom,
            _ => Edge::Left,
        }
    }
}

/// Rectangle in physical screen pixels.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Rect {
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
}

impl Rect {
    pub const fn new(left: i32, top: i32, width: i32, height: i32) -> Rect {
        Rect {
            left,
            top,
            right: left + width,
            bottom: top + height,
        }
    }

    pub const fn width(&self) -> i32 {
        self.right - self.left
    }

    pub const fn height(&self) -> i32 {
        self.bottom - self.top
    }
}

/// Two rectangles discovered in `Shell_TrayWnd`.
#[derive(Debug, Clone, Default)]
pub struct TaskbarAnchors {
    /// Icon cluster — class `"MSTaskListWClass"`.
    pub icons: Option<Rect>,
    /// System tray + clock — class `"TrayNotifyWnd"`.
    pub tray: Option<Rect>,
}

/// Max gap between taskbar.left and icons.left that we still treat as
/// "flush left". The Start button sits at the very left (~48-56 px wide
/// depending on scaling), so a gap of ≤ ~80 px means there's no real
/// idle space — it's just the Start button. Any gap larger than that
/// is Win11's centered-icon pattern with genuine idle space on the left.
const FLUSH_LEFT_MAX_GAP: i32 = 80;

/// Absolute minimum strip width — below this not worth showing.
const MIN_STRIP_W: i32 = 120;

/// Gap between strip bottom and screen bottom.
const BOTTOM_INSET: i32 = 2;

/// Place the strip. See module-level doc for the two-mode rule.
pub fn place_dynamic(
    taskbar: &Rect,
    monitor: &Rect,
    edge: Edge,
    desired_w: i32,
    win_h: i32,
    margin: i32,
    anchors: &TaskbarAnchors,
) -> Rect {
    match edge {
        Edge::Bottom => {
            let (x, w) = place_bottom(taskbar, monitor, desired_w, margin, anchors);
            Rect::new(x, monitor.bottom - win_h - BOTTOM_INSET, w, win_h)
        }
        Edge::Top => Rect::new(
            monitor.left + margin,
            taskbar.bottom + margin,
            desired_w,
            win_h,
        ),
        Edge::Left => Rect::new(
            taskbar.right + margin,
            monitor.top + margin,
            desired_w,
            win_h,
        ),
        Edge::Right => Rect::new(
            taskbar.left - desired_w - margin,
            monitor.top + margin,
            desired_w,
            win_h,
        ),
    }
}

fn place_bottom(
    taskbar: &Rect,
    monitor: &Rect,
    desired_w: i32,
    margin: i32,
    a: &TaskbarAnchors,
) -> (i32, i32) {
    if let Some(icons) = a.icons {
        let flush = icons.left - taskbar.left <= FLUSH_LEFT_MAX_GAP;

        if flush {
            // --- 靠左模式: 条 = icons.right ~ tray.left 之间 ---
            if let Some(tray) = a.tray {
                let strip_w = tray.left - icons.right - margin * 2;
                if strip_w >= MIN_STRIP_W {
                    let w = strip_w.min(desired_w);
                    let x = icons.right + margin;
                    let x = x.clamp(monitor.left + margin, monitor.right - margin - w);
                    return (x, w);
                }
            }
        } else {
            // --- 居中模式: 条 = taskbar.left ~ icons.left 之间 ---
            let strip_w = icons.left - taskbar.left - margin * 2;
            if strip_w >= MIN_STRIP_W {
                let w = strip_w.min(desired_w);
                let x = taskbar.left + margin;
                let x = x.clamp(monitor.left + margin, monitor.right - margin - w);
                return (x, w);
            }
        }
    }

    // --- 兜底: 居中 + clamp ---
    let center = taskbar.left + taskbar.width() / 2;
    let w = desired_w.clamp(MIN_STRIP_W, taskbar.width() - margin * 2);
    let x = (center - w / 2).clamp(monitor.left + margin, monitor.right - margin - w);
    (x, w)
}

// ---------------------------------------------------------------------------
// Tests — pure geometry, no Win32 involved
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    const M: Rect = Rect::new(0, 0, 1920, 1080);
    const TB: Rect = Rect::new(0, 1032, 1920, 48);

    fn a(icons: Option<Rect>, tray: Option<Rect>) -> TaskbarAnchors {
        TaskbarAnchors { icons, tray }
    }

    // -- 居中模式: icons.left > 任务栏左边缘 --------------------

    #[test]
    fn centered_strip_at_taskbar_left_edge() {
        let icons = Rect::new(560, 1032, 800, 48);
        let r = place_dynamic(&TB, &M, Edge::Bottom, 500, 45, 10, &a(Some(icons), None));
        // strip_w = 560 - 0 - 20 = 540. min(540, 500) = 500.
        assert_eq!(r.width(), 500);
        assert_eq!(r.left, 10); // taskbar.left + margin
        assert_eq!(r.right, 510);
        assert!(r.right <= 560 - 10, "右端不能超出 icons.left");
    }

    #[test]
    fn centered_icons_grow_strip_shrinks_right_edge() {
        let icons = Rect::new(250, 1032, 1200, 48); // icons.left 从 560 → 250
        let r = place_dynamic(&TB, &M, Edge::Bottom, 500, 45, 10, &a(Some(icons), None));
        // strip_w = 250 - 20 = 230. min(230, 500) = 230.
        assert_eq!(r.width(), 230);
        assert_eq!(r.left, 10);
        assert_eq!(r.right, 240); // 紧贴图标左侧
    }

    #[test]
    fn centered_too_narrow_falls_back() {
        let icons = Rect::new(110, 1032, 1600, 48);
        let r = place_dynamic(&TB, &M, Edge::Bottom, 500, 45, 10, &a(Some(icons), None));
        // strip_w = 110 - 20 = 90 < MIN_STRIP_W 120 → 退到居中
        assert_eq!(r.width(), 500);
        assert_eq!(r.left, 710);
    }

    // -- 靠左模式: icons.left ≈ taskbar.left --------------------

    #[test]
    fn left_aligned_strip_between_icons_and_tray() {
        // Win10: icons start at ~50px (right after Start button), gap = 50 ≤ 80 → 靠左模式.
        let icons = Rect::new(50, 1032, 900, 48);
        let tray = Rect::new(1800, 1032, 120, 48);
        let r = place_dynamic(
            &TB,
            &M,
            Edge::Bottom,
            500,
            45,
            10,
            &a(Some(icons), Some(tray)),
        );
        // strip_w = 1800 - 950 - 20 = 830. min(830, 500) = 500.
        assert_eq!(r.width(), 500);
        assert_eq!(r.left, 960);
        assert_eq!(r.right, 1460);
    }

    #[test]
    fn left_aligned_shrink_when_tray_grows_left() {
        let icons = Rect::new(50, 1032, 900, 48);
        let tray = Rect::new(1400, 1032, 120, 48);
        let r = place_dynamic(
            &TB,
            &M,
            Edge::Bottom,
            500,
            45,
            10,
            &a(Some(icons), Some(tray)),
        );
        assert_eq!(r.width(), 430);
        assert_eq!(r.left, 960);
    }

    #[test]
    fn left_aligned_shrink_when_icons_grow_right() {
        let icons = Rect::new(50, 1032, 1200, 48);
        let tray = Rect::new(1800, 1032, 120, 48);
        let r = place_dynamic(
            &TB,
            &M,
            Edge::Bottom,
            500,
            45,
            10,
            &a(Some(icons), Some(tray)),
        );
        assert_eq!(r.width(), 500);
        assert_eq!(r.left, 1260); // icons.right = 1250 + margin 10
    }

    #[test]
    fn left_aligned_no_tray_falls_back() {
        let icons = Rect::new(50, 1032, 900, 48);
        let r = place_dynamic(&TB, &M, Edge::Bottom, 500, 45, 10, &a(Some(icons), None));
        // 靠左模式需要 tray；缺失 → 退到居中
        assert_eq!(r.left, 710);
    }

    // -- Fallback -------------------------------

    #[test]
    fn no_icons_falls_back_to_center() {
        let r = place_dynamic(
            &TB,
            &M,
            Edge::Bottom,
            500,
            45,
            10,
            &TaskbarAnchors::default(),
        );
        assert_eq!(r.width(), 500);
        assert_eq!(r.left, 710);
    }

    // -- 其他 docked 边 ------------------------

    #[test]
    fn top_taskbar_strip_below_bar() {
        let top_tb = Rect::new(0, 0, 1920, 48);
        let r = place_dynamic(
            &top_tb,
            &M,
            Edge::Top,
            500,
            45,
            10,
            &TaskbarAnchors::default(),
        );
        assert_eq!(r.top, 58);
        assert_eq!(r.left, 10);
    }

    #[test]
    fn left_taskbar_strip_right_of_bar() {
        let tb = Rect::new(0, 0, 48, 1080);
        let r = place_dynamic(&tb, &M, Edge::Left, 500, 45, 10, &TaskbarAnchors::default());
        assert_eq!(r.left, 58);
    }

    #[test]
    fn right_taskbar_strip_left_of_bar() {
        let tb = Rect::new(1872, 0, 48, 1080);
        let r = place_dynamic(
            &tb,
            &M,
            Edge::Right,
            500,
            45,
            10,
            &TaskbarAnchors::default(),
        );
        assert_eq!(r.right, 1862);
    }

    #[test]
    fn edge_from_abe_maps_all_four() {
        assert_eq!(Edge::from_abe(0), Edge::Left);
        assert_eq!(Edge::from_abe(1), Edge::Top);
        assert_eq!(Edge::from_abe(2), Edge::Right);
        assert_eq!(Edge::from_abe(3), Edge::Bottom);
    }

    #[test]
    fn rect_helpers() {
        let r = Rect::new(5, 6, 7, 8);
        assert_eq!(r.width(), 7);
        assert_eq!(r.height(), 8);
        assert_eq!(r.right, 12);
        assert_eq!(r.bottom, 14);
    }

    #[test]
    fn autohide_taskbar_strip_stays_visible() {
        let tb = Rect::new(0, 1072, 1920, 8);
        let r = place_dynamic(
            &tb,
            &M,
            Edge::Bottom,
            500,
            45,
            10,
            &TaskbarAnchors::default(),
        );
        assert_eq!(r.bottom, M.bottom - BOTTOM_INSET);
        assert!(r.top < tb.top);
    }
}
