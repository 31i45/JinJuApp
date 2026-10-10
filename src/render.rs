//! Text rendering into a per-pixel-alpha ARGB surface.
//!
//! Pipeline: GDI draws the text **white on black** into a 32-bpp top-down
//! DIB (the red channel then carries per-pixel coverage — grayscale
//! antialiasing, correct on alpha surfaces where ClearType is not), and
//! [`premultiply`] converts that mask into the final premultiplied ARGB
//! pixels that `UpdateLayeredWindow` expects. The conversion is a pure
//! function; the GDI half is a thin resource-owning host.

use crate::win32::*;

/// Text color and maximum opacity, matching the original app's
/// `rgba(30, 30, 30, 0.85)`.
pub const COLOR: (u8, u8, u8) = (30, 30, 30);
pub const MAX_ALPHA: u8 = 217; // 0.85 * 255

/// Font size in points (the original app used 13 pt).
const FONT_PT: i32 = 13;

fn scale(c: u8, a: u8) -> u8 {
    ((u32::from(c) * u32::from(a) + 127) / 255) as u8
}

/// In-place: grayscale mask (white text on black) → premultiplied ARGB.
///
/// `buf` is BGRA, row-major, top-down; `chunks_exact` skips any trailing
/// bytes. Pure: no GDI, no globals.
pub fn premultiply(buf: &mut [u8], color: (u8, u8, u8), max_alpha: u8) {
    let (pixels, _) = buf.as_chunks_mut::<4>();
    for px in pixels {
        let coverage = px[2]; // red channel carries coverage
        let a = scale(max_alpha, coverage);
        px[0] = scale(color.2, a);
        px[1] = scale(color.1, a);
        px[2] = scale(color.0, a);
        px[3] = a;
    }
}

/// System UI font (Microsoft YaHei UI on zh-CN, Segoe UI elsewhere) at
/// `FONT_PT` points for the given DPI — no hard-coded typeface name.
unsafe fn create_font(dpi: u32) -> HFONT {
    unsafe {
        let mut ncm: NONCLIENTMETRICSW = core::mem::zeroed();
        ncm.cbSize = core::mem::size_of::<NONCLIENTMETRICSW>() as u32;
        let ok = SystemParametersInfoW(
            SPI_GETNONCLIENTMETRICS,
            core::mem::size_of::<NONCLIENTMETRICSW>() as u32,
            &mut ncm as *mut _ as *mut core::ffi::c_void,
            0,
        );
        let mut lf = if ok != 0 {
            ncm.lfMessageFont
        } else {
            core::mem::zeroed()
        };
        // Negative height = character height; points → pixels at this DPI.
        lf.lfHeight = -MulDiv(FONT_PT, dpi as i32, 72);
        lf.lfWeight = FW_LIGHT;
        // Grayscale AA: ClearType assumes an opaque destination and would
        // produce colored fringes on a per-pixel-alpha surface.
        lf.lfQuality = ANTIALIASED_QUALITY;
        CreateFontIndirectW(&lf)
    }
}

/// An offscreen DIB + memory DC + font, owned and cleaned up on drop.
pub struct Surface {
    pub width: i32,
    pub height: i32,
    hdc: HDC,
    hbmp: HBITMAP,
    bits: *mut u8,
    font: HFONT,
    old_bmp: HGDIOBJ,
    old_font: HGDIOBJ,
}

impl Surface {
    /// Create a surface of `width` × `height` px at `dpi`.
    pub fn new(width: i32, height: i32, dpi: u32) -> Option<Surface> {
        if width <= 0 || height <= 0 {
            return None;
        }
        unsafe {
            let hdc = CreateCompatibleDC(core::ptr::null_mut());
            if hdc.is_null() {
                return None;
            }
            let header = BITMAPINFOHEADER {
                biSize: core::mem::size_of::<BITMAPINFOHEADER>() as u32,
                biWidth: width,
                biHeight: -height, // top-down
                biPlanes: 1,
                biBitCount: 32,
                biCompression: BI_RGB,
                biSizeImage: 0,
                biXPelsPerMeter: 0,
                biYPelsPerMeter: 0,
                biClrUsed: 0,
                biClrImportant: 0,
            };
            let mut bits: *mut core::ffi::c_void = core::ptr::null_mut();
            let hbmp = CreateDIBSection(
                hdc,
                &header,
                DIB_RGB_COLORS,
                &mut bits,
                core::ptr::null_mut(),
                0,
            );
            if hbmp.is_null() || bits.is_null() {
                DeleteDC(hdc);
                return None;
            }
            let old_bmp = SelectObject(hdc, hbmp);
            let font = create_font(dpi);
            let old_font = SelectObject(hdc, font);
            Some(Surface {
                width,
                height,
                hdc,
                hbmp,
                bits: bits as *mut u8,
                font,
                old_bmp,
                old_font,
            })
        }
    }

    /// Measure the rendered width of `text` (px).
    pub fn measure(&self, text: &str) -> i32 {
        let mut wide: Vec<u16> = text.encode_utf16().collect();
        let mut rc = RECT {
            left: 0,
            top: 0,
            right: 0,
            bottom: self.height,
        };
        unsafe {
            DrawTextW(
                self.hdc,
                wide.as_mut_ptr(),
                wide.len() as i32,
                &mut rc,
                DT_SINGLELINE | DT_CALCRECT | DT_NOPREFIX,
            );
        }
        rc.right.max(0)
    }

    /// Draw `text` at x-offset `x`, producing the final ARGB pixels in
    /// [`bits`]. White-on-black into the DIB, then premultiply.
    pub fn draw(&mut self, text: &str, x: i32) {
        unsafe {
            let mut rc = RECT {
                left: 0,
                top: 0,
                right: self.width,
                bottom: self.height,
            };
            FillRect(self.hdc, &mut rc, GetStockObject(BLACK_BRUSH));
            SetBkColor(self.hdc, rgb(0, 0, 0));
            SetTextColor(self.hdc, rgb(255, 255, 255));
            let mut wide: Vec<u16> = text.encode_utf16().collect();
            if !wide.is_empty() {
                let mut target = RECT {
                    left: x,
                    top: 0,
                    right: x + self.width,
                    bottom: self.height,
                };
                // DT_NOCLIP is essential: DrawTextW by default clips to the
                // layout rect, and this rect slides with the text origin —
                // the tail would never render past (x + width). With
                // DT_NOCLIP, GDI still clips at the DIB surface, which is
                // exactly the marquee's window edge.
                DrawTextW(
                    self.hdc,
                    wide.as_mut_ptr(),
                    wide.len() as i32,
                    &mut target,
                    DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_NOCLIP,
                );
            }
        }
        premultiply(self.bits_mut(), COLOR, MAX_ALPHA);
    }

    /// The DIB pixel buffer, `w*h*4` bytes, BGRA top-down.
    pub fn bits_mut(&mut self) -> &mut [u8] {
        let len = (self.width as usize) * (self.height as usize) * 4;
        unsafe { core::slice::from_raw_parts_mut(self.bits, len) }
    }

    pub fn hdc(&self) -> HDC {
        self.hdc
    }
}

impl Drop for Surface {
    fn drop(&mut self) {
        unsafe {
            SelectObject(self.hdc, self.old_font);
            SelectObject(self.hdc, self.old_bmp);
            DeleteObject(self.font);
            DeleteObject(self.hbmp);
            DeleteDC(self.hdc);
        }
    }
}

// ---------------------------------------------------------------------------
// Tests (pure half)
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn px(b: &[u8]) -> [u8; 4] {
        [b[0], b[1], b[2], b[3]]
    }

    #[test]
    fn background_is_fully_transparent() {
        let mut buf = [0u8, 0, 0, 0xFF, 0, 0, 0, 0xFF];
        premultiply(&mut buf, COLOR, MAX_ALPHA);
        assert_eq!(px(&buf[..4]), [0, 0, 0, 0]);
        assert_eq!(px(&buf[4..]), [0, 0, 0, 0]);
    }

    #[test]
    fn solid_pixel_gets_full_color_and_capped_alpha() {
        // White solid: coverage 255. Premultiplied: rgb = color * alpha,
        // so 30 * 217/255 = 26 — the *displayed* color stays 30/255 at
        // 85% opacity, which is exactly what premultiplication means.
        let mut buf = [0u8, 0, 255, 0];
        premultiply(&mut buf, COLOR, MAX_ALPHA);
        let expected = [
            scale(COLOR.0, MAX_ALPHA),
            scale(COLOR.1, MAX_ALPHA),
            scale(COLOR.2, MAX_ALPHA),
            MAX_ALPHA,
        ];
        assert_eq!(px(&buf), expected);
        assert_eq!(expected[0], 26); // 30 * 217 / 255, rounded
    }

    #[test]
    fn half_coverage_blends_by_coverage() {
        let mut buf = [0u8, 0, 128, 0];
        premultiply(&mut buf, COLOR, MAX_ALPHA);
        let a = scale(217, 128);
        assert_eq!(px(&buf), [scale(30, a), scale(30, a), scale(30, a), a]);
    }

    #[test]
    fn full_opacity_max_alpha_is_255() {
        let mut buf = [0u8, 0, 255, 0];
        premultiply(&mut buf, (10, 20, 30), 255);
        assert_eq!(px(&buf), [30, 20, 10, 255]);
    }

    #[test]
    fn trailing_partial_chunk_is_untouched() {
        let mut buf = [0u8, 0, 255, 0, 9, 9];
        premultiply(&mut buf, COLOR, MAX_ALPHA);
        assert_eq!(&buf[4..], &[9, 9]);
    }

    #[test]
    fn scale_rounds_correctly() {
        assert_eq!(scale(255, 255), 255);
        assert_eq!(scale(0, 255), 0);
        assert_eq!(scale(255, 0), 0);
        // 30*217/255 = 25.53 → 26
        assert_eq!(scale(30, 217), 26);
    }

    /// Regression: DrawTextW without DT_NOCLIP clips to the layout rect,
    /// which slides with the text origin during the marquee — the tail of
    /// a long quote never rendered (user-visible as ~29 CJK chars max).
    /// A long text drawn at a negative origin must still put ink at the
    /// strip's right edge, where the tail arrives during the slide.
    #[test]
    fn draw_renders_tail_when_origin_is_negative() {
        let mut s = Surface::new(500, 36, 96).expect("GDI surface");
        // ~120 ASCII chars ≈ 800 px wide at 13 pt — far wider than 500.
        let text = "The quick brown fox jumps over the lazy dog and keeps \
                    running far beyond the edge of the strip ABCDEFGHIJKL \
                    0123456789 abcdefghijklmnopqrstuvwxyz";
        assert!(s.measure(text) > 700, "test text must overflow the strip");
        s.draw(text, -300);
        // Any ink (alpha > 0) in the rightmost 100 px column band.
        let (w, h) = (s.width as usize, s.height as usize);
        let buf = s.bits_mut();
        let mut ink = 0;
        for y in 0..h {
            for xx in (w - 100)..w {
                if buf[(y * w + xx) * 4 + 3] > 0 {
                    ink += 1;
                }
            }
        }
        assert!(ink > 0, "tail must render at the right edge (DT_NOCLIP)");
    }

    /// The user-reported quote must overflow the 470 px resting width so
    /// the marquee engages at all (guards CJK width measurement).
    #[test]
    fn measure_cjk_long_quote_overflows_strip() {
        let s = Surface::new(500, 36, 96).expect("GDI surface");
        let q = "翻山越岭只为一场遇见，全力以赴只为一次改变。答案就在眼泪里，\
                 也凝结在汗水中、感官和心灵。";
        let w = s.measure(q);
        assert!(w > 470, "44-char CJK quote must overflow 470 px (got {w})");
    }

    /// The symmetric half: with the origin at rest the head is visible.
    #[test]
    fn draw_renders_head_at_rest() {
        let mut s = Surface::new(500, 36, 96).expect("GDI surface");
        s.draw("金句", 15);
        let (w, h) = (s.width as usize, s.height as usize);
        let buf = s.bits_mut();
        let mut ink = 0;
        for y in 0..h {
            for xx in 0..100 {
                if buf[(y * w + xx) * 4 + 3] > 0 {
                    ink += 1;
                }
            }
        }
        assert!(ink > 0, "head must render at the left edge");
    }
}
