//! Pure marquee scroll state machine. No Win32 types — fully unit-tested.

/// Horizontal padding (px) between window edge and resting text.
pub const PAD: i32 = 15;

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
}
