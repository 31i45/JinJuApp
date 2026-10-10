//! Pure taskbar geometry: docked-edge classification, screen-space
//! rectangles, and the window placement derived from them. No Win32
//! types — fully unit-tested.

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
}

/// Gap between the strip and the screen's bottom edge (keeps glyph
/// descenders off the very last pixel row).
const BOTTOM_INSET: i32 = 2;

/// Place the window relative to the taskbar.
///
/// Bottom-docked taskbar (the default): the strip anchors to the
/// screen's bottom-left corner — inside the taskbar band, exactly where
/// v1 drew it. It reads as taskbar text, and it stays put when the bar
/// auto-hides. Top/side-docked taskbars: the strip goes on the desktop
/// side of the bar (horizontal text cannot live inside a vertical one),
/// never overlapping it.
pub fn place(
    taskbar: &Rect,
    monitor: &Rect,
    edge: Edge,
    win_w: i32,
    win_h: i32,
    margin: i32,
) -> Rect {
    match edge {
        Edge::Bottom => Rect::new(
            monitor.left + margin,
            monitor.bottom - win_h - BOTTOM_INSET,
            win_w,
            win_h,
        ),
        Edge::Top => Rect::new(monitor.left + margin, taskbar.bottom + margin, win_w, win_h),
        Edge::Left => Rect::new(taskbar.right + margin, monitor.top + margin, win_w, win_h),
        Edge::Right => Rect::new(
            taskbar.left - win_w - margin,
            monitor.top + margin,
            win_w,
            win_h,
        ),
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    const M: Rect = Rect {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };

    #[test]
    fn bottom_taskbar_strip_sits_in_screen_corner_inside_the_bar() {
        // v1 parity: for a bottom-docked taskbar the strip anchors to the
        // screen corner (inside the taskbar band), not to the bar's
        // desktop side.
        let tb = Rect {
            left: 0,
            top: 1032,
            right: 1920,
            bottom: 1080,
        };
        let r = place(&tb, &M, Edge::Bottom, 500, 45, 10);
        assert_eq!(r, Rect::new(10, 1080 - 45 - BOTTOM_INSET, 500, 45));
        assert!(
            r.bottom > tb.top && r.bottom <= M.bottom,
            "strip must be inside the taskbar band"
        );
    }

    #[test]
    fn bottom_taskbar_strip_survives_autohide() {
        // With an auto-hidden taskbar only a sliver is on screen; the
        // strip, anchored to the screen corner, remains fully visible.
        let tb = Rect {
            left: 0,
            top: 1072,
            right: 1920,
            bottom: 1080,
        };
        let r = place(&tb, &M, Edge::Bottom, 500, 45, 10);
        assert_eq!(r.bottom, M.bottom - BOTTOM_INSET);
        assert!(r.top < tb.top, "strip rides above the hidden sliver");
    }

    #[test]
    fn places_below_top_taskbar() {
        let tb = Rect {
            left: 0,
            top: 0,
            right: 1920,
            bottom: 48,
        };
        let r = place(&tb, &M, Edge::Top, 500, 45, 10);
        assert_eq!(r, Rect::new(10, 58, 500, 45));
        assert!(r.top >= tb.bottom);
    }

    #[test]
    fn places_right_of_left_taskbar() {
        let tb = Rect {
            left: 0,
            top: 0,
            right: 48,
            bottom: 1080,
        };
        let r = place(&tb, &M, Edge::Left, 500, 45, 10);
        assert_eq!(r, Rect::new(58, 10, 500, 45));
        assert!(r.left >= tb.right);
    }

    #[test]
    fn places_left_of_right_taskbar() {
        let tb = Rect {
            left: 1872,
            top: 0,
            right: 1920,
            bottom: 1080,
        };
        let r = place(&tb, &M, Edge::Right, 500, 45, 10);
        assert_eq!(r, Rect::new(1872 - 500 - 10, 10, 500, 45));
        assert!(r.right <= tb.left);
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
        assert_eq!(r.right, 12);
        assert_eq!(r.bottom, 14);
    }
}
