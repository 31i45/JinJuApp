//! Pure marquee scroll state machine. No Win32 types — fully unit-tested.

/// Horizontal padding (px) between window edge and resting text.
pub const PAD: i32 = 15;

/// Marquee scroller.
///
/// Fitting text rests left-aligned at `PAD`. Wider text cycles, matching
/// the original PyQt app (the window never moves — only the text does):
/// rest at PAD for [`HOLD_MS`] (the readable head of the quote), slide
/// left until the tail reaches the right edge, hold there briefly, then
/// fade out, snap back invisibly, and fade in at the rest position —
/// the reset jump can never be seen. Every quote change restarts from
/// the rest position so the user immediately sees the first screenful.
pub struct Scroll {
    text_w: i32,
    view_w: i32,
    speed: f32,   // px per ms
    hold_px: f32, // hold phase expressed as scroll distance
    tail_hold_px: f32,
    fade_px: f32, // one fade (out or in) expressed as scroll distance
    offset: f32,  // distance travelled in the current cycle, in [0, cycle)
}

/// Pause at the rest position before each scroll cycle (ms).
const HOLD_MS: f32 = 2000.0;
/// Pause at the tail before the fade-out (ms).
const TAIL_HOLD_MS: f32 = 600.0;
/// Duration of each fade — out at the tail, in at the rest position (ms).
const FADE_MS: f32 = 250.0;

impl Scroll {
    pub fn new(view_w: i32, text_w: i32, speed_px_per_s: f32) -> Scroll {
        let speed = speed_px_per_s / 1000.0;
        Scroll {
            text_w,
            view_w,
            speed,
            hold_px: speed * HOLD_MS,
            tail_hold_px: speed * TAIL_HOLD_MS,
            fade_px: speed * FADE_MS,
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
        self.hold_px + self.max_scroll() + self.tail_hold_px + 2.0 * self.fade_px
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
    ///   tail at the right edge)
    /// - tail hold + fade-out: x frozen at the tail
    /// - fade-in: x = PAD — the snap back happens at alpha 0, so the
    ///   reset is never visible; the wrap is continuous
    pub fn x(&self) -> i32 {
        if !self.scrolling() {
            return PAD;
        }
        let pad = PAD as f32;
        let a = self.offset - self.hold_px;
        let ms = self.max_scroll();
        let x = if a <= 0.0 {
            pad
        } else if a <= ms {
            pad - a
        } else if a <= ms + self.tail_hold_px + self.fade_px {
            pad - ms // tail hold and fade-out: frozen at the tail
        } else {
            pad // fade-in happens back at the rest position
        };
        x.round() as i32
    }

    /// Text opacity for the current state, in [0, 1].
    ///
    /// Full through the hold and slide, linearly down to 0 during the
    /// fade-out at the tail, then linearly back up to 1 during the
    /// fade-in at the rest position — the position reset only ever
    /// happens while the text is invisible.
    pub fn alpha(&self) -> f32 {
        if !self.scrolling() {
            return 1.0;
        }
        let a = self.offset - self.hold_px;
        let ms = self.max_scroll();
        let fade_out_start = ms + self.tail_hold_px;
        let fade_in_start = fade_out_start + self.fade_px;
        if a <= fade_out_start {
            1.0
        } else if a <= fade_in_start {
            1.0 - (a - fade_out_start) / self.fade_px
        } else {
            ((a - fade_in_start) / self.fade_px).min(1.0)
        }
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
    fn tail_holds_before_reset() {
        // cycle phases at 1 px/ms: hold 2000, slide 430, tail hold 600,
        // fade-out 250, fade-in 250 — cycle = 3530.
        let mut s = Scroll::new(500, 900, 1000.0);
        s.tick(2000.0 + 430.0 + 300.0); // mid tail-hold
        assert_eq!(s.x(), 500 - PAD - 900, "position frozen at the tail");
        assert_eq!(s.alpha(), 1.0, "tail hold is fully opaque");
    }

    #[test]
    fn fade_out_is_linear_and_position_frozen() {
        let mut s = Scroll::new(500, 900, 1000.0);
        s.tick(2000.0 + 430.0 + 600.0 + 125.0); // halfway through fade-out
        let a = s.alpha();
        assert!((a - 0.5).abs() < 0.01, "mid-fade alpha ≈ 0.5 (got {a})");
        assert_eq!(s.x(), 500 - PAD - 900, "position frozen during fade");
    }

    #[test]
    fn reset_jump_is_invisible() {
        let mut s = Scroll::new(500, 900, 1000.0);
        s.tick(2000.0 + 430.0 + 600.0 + 250.0); // fade-out complete
        assert_eq!(s.x(), 500 - PAD - 900, "still at the tail");
        assert_eq!(s.alpha(), 0.0, "fully faded out");
        s.tick(1.0); // snapped back to rest
        assert_eq!(s.x(), PAD);
        assert!(
            s.alpha() < 0.01,
            "the position jump happens while invisible (alpha = {})",
            s.alpha()
        );
    }

    #[test]
    fn wrap_is_continuous_back_to_rest() {
        let mut s = Scroll::new(500, 900, 1000.0);
        s.tick(3529.0); // one ms before cycle end
        assert_eq!(s.x(), PAD, "fade-in happens at the rest position");
        assert!(s.alpha() > 0.99, "faded back in before wrap");
        s.tick(1.0);
        assert_eq!(s.x(), PAD, "cycle wraps exactly to the rest position");
        assert_eq!(s.alpha(), 1.0);
    }

    #[test]
    fn slide_is_monotonic_and_reset_jumps_once() {
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
        // tail hold + fade-out: 850 px of travel, position frozen
        for _ in 0..850 {
            s.tick(1.0);
            assert_eq!(s.x(), 500 - PAD - 900, "frozen through hold and fade");
        }
        // The single intended jump, taken while alpha is 0.
        s.tick(1.0);
        assert_eq!(s.x(), PAD);
        // fade-in: stays at rest until the wrap
        for _ in 0..249 {
            s.tick(1.0);
            assert_eq!(s.x(), PAD, "fade-in keeps the rest position");
        }
        assert_eq!(s.alpha(), 1.0, "cycle ends fully opaque");
    }

    #[test]
    fn alpha_is_full_through_hold_and_slide() {
        let mut s = Scroll::new(500, 900, 1000.0);
        for _ in 0..2430 {
            // hold 2000 + slide 430
            assert_eq!(s.alpha(), 1.0, "no fading while reading");
            s.tick(1.0);
        }
        assert_eq!(s.x(), 500 - PAD - 900);
    }

    #[test]
    fn resting_text_is_fully_opaque() {
        let mut s = Scroll::new(500, 200, 30.0);
        s.tick(10_000.0);
        assert_eq!(s.alpha(), 1.0);
    }

    #[test]
    fn set_text_restores_full_opacity() {
        let mut s = Scroll::new(500, 900, 1000.0);
        s.tick(2000.0 + 430.0 + 600.0 + 125.0); // mid fade-out
        assert!(s.alpha() < 1.0);
        s.set_text(1200); // still long
        assert!(s.scrolling());
        assert_eq!(s.alpha(), 1.0, "quote change is fully opaque at rest");
        assert_eq!(s.x(), PAD);
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
