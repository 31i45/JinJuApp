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
/// Fitting text rests left-aligned at `PAD`. Wider text cycles, matching
/// the original PyQt app (the window never moves — only the text does):
/// rest at PAD for [`HOLD_MS`] (the readable head of the quote), slide
/// left until the tail reaches the right edge — the strip stays filled
/// the whole time — then wipe rightward back into the rest position and
/// repeat. Every quote change restarts from the rest position so the
/// user immediately sees the first screenful.
pub struct Scroll {
    text_w: i32,
    view_w: i32,
    speed: f32,   // px per ms
    hold_px: f32, // hold phase expressed as scroll distance
    offset: f32,  // distance travelled in the current cycle, in [0, cycle)
}

/// Pause at the rest position before each scroll cycle (ms).
const HOLD_MS: f32 = 2000.0;

impl Scroll {
    pub fn new(view_w: i32, text_w: i32, speed_px_per_s: f32) -> Scroll {
        let speed = speed_px_per_s / 1000.0;
        Scroll {
            text_w,
            view_w,
            speed,
            hold_px: speed * HOLD_MS,
            offset: 0.0,
        }
    }

    pub fn scrolling(&self) -> bool {
        self.text_w > self.view_w - 2 * PAD
    }

    /// New text resets the cycle to the rest position.
    pub fn set_text(&mut self, text_w: i32) {
        self.text_w = text_w;
        self.offset = 0.0;
    }

    /// Leftward travel: from the rest position until the text tail
    /// reaches the right edge. Positive whenever `scrolling()`.
    fn max_scroll(&self) -> f32 {
        (self.text_w - (self.view_w - 2 * PAD)).max(1) as f32
    }

    fn cycle(&self) -> f32 {
        self.hold_px + self.max_scroll() + self.text_w.max(1) as f32
    }

    pub fn tick(&mut self, dt_ms: f32) {
        if self.scrolling() {
            self.offset = (self.offset + self.speed * dt_ms) % self.cycle();
        }
    }

    /// Text origin x for the current state.
    ///
    /// Phases by `a = offset - hold_px`:
    /// - `a < 0`: hold, x = PAD
    /// - `0 <= a <= max_scroll`: slide left, x = PAD - a (ends with the
    ///   tail at the right edge — the strip is never left empty)
    /// - otherwise: wipe back right, x = PAD - text_w + (a - max_scroll),
    ///   ending exactly at PAD so the wrap is continuous
    pub fn x(&self) -> i32 {
        if !self.scrolling() {
            return PAD;
        }
        let pad = PAD as f32;
        let tw = self.text_w as f32;
        let a = self.offset - self.hold_px;
        let ms = self.max_scroll();
        let x = if a <= 0.0 {
            pad
        } else if a <= ms {
            pad - a
        } else {
            pad - tw + (a - ms)
        };
        x.round() as i32
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
    fn long_text_rests_first_then_slides_left() {
        // A new long quote is immediately readable at the rest position
        // (hold phase) instead of entering from off-screen.
        let s = Scroll::new(500, 900, 30.0);
        assert!(s.scrolling());
        assert_eq!(s.x(), PAD);
    }

    #[test]
    fn hold_lasts_exactly_hold_ms() {
        // 1 px/ms: the hold covers the first 2000 px of travel.
        let mut s = Scroll::new(500, 900, 1000.0);
        s.tick(1999.0);
        assert_eq!(s.x(), PAD, "still holding at 1999 ms");
        s.tick(2.0);
        assert_eq!(s.x(), PAD - 1, "slide begins right after the hold");
    }

    #[test]
    fn slide_left_ends_with_tail_at_right_edge() {
        // avail = 500 - 2*PAD = 470, max_scroll = 900 - 470 = 430
        let mut s = Scroll::new(500, 900, 1000.0);
        s.tick(2000.0 + 430.0);
        // The text's last pixel sits at the right padding boundary: the
        // strip is still filled, the quote's ending fully readable.
        assert_eq!(s.x(), 500 - PAD - 900);
    }

    #[test]
    fn wipe_phase_reenters_from_left_moving_right() {
        let mut s = Scroll::new(500, 900, 1000.0);
        s.tick(2000.0 + 430.0 + 1.0); // one px into the wipe
        // The text block starts just off the left edge and slides right.
        assert_eq!(s.x(), PAD - 900 + 1);
    }

    #[test]
    fn wrap_is_continuous_back_to_rest() {
        // cycle = hold 2000 + slide 430 + wipe 900 = 3330
        let mut s = Scroll::new(500, 900, 1000.0);
        s.tick(3329.0);
        let before_wrap = s.x();
        s.tick(1.0);
        // The underlying function is continuous; integer rounding may
        // differ by 1 px across the wrap boundary.
        assert!(
            (before_wrap - PAD).abs() <= 1,
            "approaching wrap returns to rest (got {before_wrap})"
        );
        assert_eq!(s.x(), PAD, "cycle wraps exactly to the rest position");
    }

    #[test]
    fn slide_is_monotonic_within_each_phase() {
        let mut s = Scroll::new(500, 900, 1000.0);
        s.tick(2000.0); // consume the hold
        // slide-left phase: 430 px, x from PAD down to 500-PAD-900 = -415
        let mut last = s.x();
        for _ in 0..430 {
            s.tick(1.0);
            let x = s.x();
            assert!(x <= last, "slide-left: x must not increase ({x} > {last})");
            last = x;
        }
        assert_eq!(last, 500 - PAD - 900, "slide ends with tail at right edge");
        // The slide→wipe transition is the one intended jump (rewind).
        s.tick(1.0);
        let mut last = s.x();
        assert_eq!(last, PAD - 900 + 1, "wipe starts just off the left edge");
        // wipe phase: 900 px rightward back to rest
        for _ in 0..899 {
            s.tick(1.0);
            let x = s.x();
            assert!(x >= last, "wipe: x must not decrease ({x} < {last})");
            last = x;
        }
        assert_eq!(last, PAD, "wipe ends exactly at the rest position");
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
        s.tick(2600.0); // mid-slide
        assert_ne!(s.x(), PAD);
        s.set_text(400);
        assert!(!s.scrolling());
        assert_eq!(s.x(), PAD);
    }

    #[test]
    fn set_text_restarts_at_rest_for_long_text() {
        let mut s = Scroll::new(500, 900, 1000.0);
        s.tick(2600.0); // mid-slide
        s.set_text(1200); // still long
        assert!(s.scrolling());
        assert_eq!(s.x(), PAD, "quote change returns to the rest position");
    }

    // -- place -------------------------------------------------------------

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
