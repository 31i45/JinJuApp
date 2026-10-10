//! Pure layout logic: marquee scroll state machine and taskbar-anchored
//! window placement. No Win32 types — fully unit-tested.

/// Horizontal padding (px) between window edge and resting text.
pub const PAD: i32 = 15;

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

/// Marquee scroller.
///
/// At rest the text sits left-aligned with `PAD` margin. When wider than
/// the view it cycles: enter from the right edge, travel left, exit fully,
/// re-enter — the period is exactly `text_w + view_w`.
pub struct Scroll {
    text_w: i32,
    view_w: i32,
    speed: f32,  // px per ms
    offset: f32, // px travelled in the current cycle, in [0, period)
}

impl Scroll {
    pub fn new(view_w: i32, text_w: i32, speed_px_per_s: f32) -> Scroll {
        Scroll {
            text_w,
            view_w,
            speed: speed_px_per_s / 1000.0,
            offset: 0.0,
        }
    }

    pub fn scrolling(&self) -> bool {
        self.text_w > self.view_w - 2 * PAD
    }

    /// New text resets the cycle.
    pub fn set_text(&mut self, text_w: i32) {
        self.text_w = text_w;
        self.offset = 0.0;
    }

    fn period(&self) -> f32 {
        ((self.text_w + self.view_w).max(1)) as f32
    }

    pub fn tick(&mut self, dt_ms: f32) {
        if self.scrolling() {
            self.offset = (self.offset + self.speed * dt_ms) % self.period();
        }
    }

    /// Text origin x for the current state.
    pub fn x(&self) -> i32 {
        if !self.scrolling() {
            PAD
        } else {
            self.view_w - self.offset.round() as i32
        }
    }
}

/// Place the window on the desktop strip adjacent to the taskbar.
///
/// `monitor` is the full monitor rect (not the work area) so the math is
/// independent of where the taskbar is docked; the window always lands on
/// the desktop side of the taskbar, never overlapping it.
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
            taskbar.top - win_h - margin,
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

    // -- Scroll ------------------------------------------------------------

    #[test]
    fn short_text_rests_left_aligned() {
        let s = Scroll::new(500, 200, 30.0);
        assert!(!s.scrolling());
        assert_eq!(s.x(), PAD);
    }

    #[test]
    fn exactly_fitting_text_rests() {
        // text_w == view_w - 2*PAD fits without scrolling
        let s = Scroll::new(500, 470, 30.0);
        assert!(!s.scrolling());
        assert_eq!(s.x(), PAD);
    }

    #[test]
    fn long_text_starts_off_right_edge() {
        let s = Scroll::new(500, 900, 30.0);
        assert!(s.scrolling());
        assert_eq!(s.x(), 500); // fully off-screen to the right
    }

    #[test]
    fn tick_advances_monotonically_until_wrap() {
        let mut s = Scroll::new(500, 900, 30.0); // speed = 0.03 px/ms
        let mut last = s.x();
        // 10 ms steps, well before the 1400 px period
        for _ in 0..200 {
            s.tick(10.0);
            let x = s.x();
            assert!(x <= last, "x must never increase (got {x} after {last})");
            last = x;
        }
    }

    #[test]
    fn tick_wraps_exactly_at_period() {
        let mut s = Scroll::new(500, 900, 1000.0); // 1 px per ms
        s.tick(1400.0 - 0.5); // period = 500 + 900 = 1400
        let before_wrap = s.x();
        s.tick(0.5);
        assert!(
            before_wrap < -850,
            "just before wrap text is exiting left (x={before_wrap})"
        );
        assert_eq!(s.x(), 500, "after wrap text re-enters from the right edge");
    }

    #[test]
    fn tick_is_noop_for_resting_text() {
        let mut s = Scroll::new(500, 200, 30.0);
        s.tick(10_000.0);
        assert_eq!(s.x(), PAD);
    }

    #[test]
    fn set_text_resets_cycle() {
        let mut s = Scroll::new(500, 900, 1000.0);
        s.tick(700.0);
        assert_ne!(s.x(), 500);
        s.set_text(400);
        assert!(!s.scrolling());
        assert_eq!(s.x(), PAD);
    }

    // -- place -------------------------------------------------------------

    const M: Rect = Rect {
        left: 0,
        top: 0,
        right: 1920,
        bottom: 1080,
    };

    #[test]
    fn places_above_bottom_taskbar() {
        let tb = Rect {
            left: 0,
            top: 1032,
            right: 1920,
            bottom: 1080,
        };
        let r = place(&tb, &M, Edge::Bottom, 500, 45, 10);
        assert_eq!(r, Rect::new(10, 1032 - 45 - 10, 500, 45));
        assert!(r.bottom <= tb.top, "must not overlap the taskbar");
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
