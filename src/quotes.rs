//! Quote corpus: parsing, on-disk loading, first-run export, fallback.
//!
//! The corpus lives in an editable `quotes.json` next to the executable —
//! the single channel for content updates, so no re-release is ever needed
//! to change what the app displays. The original 451-quote corpus is
//! embedded at compile time as the default and is exported on first run.
//! A broken user file degrades to the embedded corpus with a
//! human-readable reason instead of failing to start.

use serde::Deserialize;
use std::fs;
use std::path::Path;
use std::time::{SystemTime, UNIX_EPOCH};

/// Default corpus, embedded at compile time.
pub const EMBEDDED: &str = include_str!("../quotes.json");

/// Where the corpus lives relative to the executable.
pub const FILE_NAME: &str = "quotes.json";

/// How the corpus for this session came to be.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Source {
    /// Loaded from the user's editable file.
    User(Vec<String>),
    /// No file existed; the embedded corpus was exported and is in use.
    Exported(Vec<String>),
    /// The user file is unusable; running on the embedded corpus.
    Fallback { lines: Vec<String>, reason: String },
}

impl Source {
    pub fn lines(&self) -> &[String] {
        match self {
            Source::User(v) | Source::Exported(v) => v,
            Source::Fallback { lines, .. } => lines,
        }
    }

    /// Human-readable problem report, if any, shown once at startup.
    pub fn note(&self) -> Option<&str> {
        match self {
            Source::Fallback { reason, .. } => Some(reason),
            _ => None,
        }
    }
}

#[derive(Deserialize)]
struct Item {
    quote: String,
}

/// Parse the corpus format: a JSON array of `{"quote": "..."}` objects,
/// UTF-8 with an optional BOM. Blank entries are dropped.
pub fn parse(raw: &str) -> Result<Vec<String>, String> {
    let s = raw.strip_prefix('\u{feff}').unwrap_or(raw);
    if s.trim().is_empty() {
        return Err("文件为空".to_string());
    }
    let items: Vec<Item> = serde_json::from_str(s).map_err(|e| format!("JSON 解析失败: {e}"))?;
    let lines: Vec<String> = items
        .into_iter()
        .map(|i| i.quote)
        .filter(|q| !q.trim().is_empty())
        .collect();
    if lines.is_empty() {
        Err("没有非空的 \"quote\" 条目".to_string())
    } else {
        Ok(lines)
    }
}

fn embedded() -> Vec<String> {
    parse(EMBEDDED).expect("embedded corpus is valid (guarded by unit test)")
}

/// Load the corpus from `dir` (the executable's directory).
///
/// Missing file → export the embedded corpus, then use it.
/// Present but unreadable/unparseable → embedded corpus + reason.
pub fn load(dir: &Path) -> Source {
    let path = dir.join(FILE_NAME);
    match fs::read_to_string(&path) {
        Ok(raw) => match parse(&raw) {
            Ok(lines) => Source::User(lines),
            Err(reason) => Source::Fallback {
                lines: embedded(),
                reason: format!("{}: {reason}", path.display()),
            },
        },
        Err(e) if e.kind() == std::io::ErrorKind::NotFound => match fs::write(&path, EMBEDDED) {
            Ok(()) => Source::Exported(embedded()),
            Err(e) => Source::Fallback {
                lines: embedded(),
                reason: format!("无法创建 {}: {e}", path.display()),
            },
        },
        Err(e) => Source::Fallback {
            lines: embedded(),
            reason: format!("无法读取 {}: {e}", path.display()),
        },
    }
}

// ---------------------------------------------------------------------------
// Randomness: xorshift64* with Lemire range reduction (no modulo bias).
// ---------------------------------------------------------------------------

/// Small deterministic PRNG — plenty for picking quotes.
pub struct Rng(u64);

impl Rng {
    pub fn new(seed: u64) -> Self {
        Rng(if seed == 0 {
            0x9E37_79B9_7F4A_7C15
        } else {
            seed
        })
    }

    /// Seed from wall clock and process id; unique per process run.
    pub fn from_entropy() -> Self {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|d| d.as_nanos() as u64)
            .unwrap_or(0);
        Rng::new(nanos ^ (u64::from(std::process::id()) << 32))
    }

    pub fn next_u64(&mut self) -> u64 {
        let mut x = self.0;
        x ^= x >> 12;
        x ^= x << 25;
        x ^= x >> 27;
        self.0 = x;
        x.wrapping_mul(0x2545_F491_4F6C_DD1D)
    }

    /// Uniform in `[0, n)`; returns 0 for `n == 0`.
    pub fn below(&mut self, n: u64) -> u64 {
        if n == 0 {
            return 0;
        }
        ((self.next_u64() as u128 * n as u128) >> 64) as u64
    }
}

// ---------------------------------------------------------------------------
// Shuffle bag: every quote is shown exactly once per round (Fisher–Yates),
// then the order is reshuffled. No immediate repeats within a round.
// ---------------------------------------------------------------------------

pub struct Bag {
    order: Vec<u32>,
    pos: usize,
}

impl Bag {
    /// Bag over `len` items. Requires `len >= 1` (guaranteed by `parse`).
    pub fn new(len: usize) -> Self {
        debug_assert!(len >= 1, "corpus is never empty (guarded by parse)");
        Bag {
            order: (0..len as u32).collect(),
            pos: len, // force a shuffle on first draw
        }
    }

    /// Draw the next index.
    pub fn next(&mut self, rng: &mut Rng) -> usize {
        debug_assert!(!self.order.is_empty());
        if self.pos >= self.order.len() {
            shuffle(&mut self.order, rng);
            self.pos = 0;
        }
        let i = self.order[self.pos] as usize;
        self.pos += 1;
        i
    }
}

/// Fisher–Yates, in-place.
fn shuffle(v: &mut [u32], rng: &mut Rng) {
    for i in (1..v.len()).rev() {
        let j = rng.below((i + 1) as u64) as usize;
        v.swap(i, j);
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn temp_dir(tag: &str) -> std::path::PathBuf {
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let d = std::env::temp_dir().join(format!("jinju-{tag}-{}-{nanos}", std::process::id()));
        fs::create_dir_all(&d).unwrap();
        d
    }

    // -- parse -------------------------------------------------------------

    #[test]
    fn parse_accepts_valid_array() {
        let v = parse(r#"[{"quote":"梦在前方，路在脚下。"},{"quote":"每一天都是新的开始。"}]"#);
        assert_eq!(
            v.unwrap(),
            vec!["梦在前方，路在脚下。", "每一天都是新的开始。"]
        );
    }

    #[test]
    fn parse_strips_bom() {
        let raw = "\u{feff}[{\"quote\":\"测试\"}]";
        assert_eq!(parse(raw).unwrap(), vec!["测试"]);
    }

    #[test]
    fn parse_rejects_empty_file() {
        assert!(parse("   \n\t ").unwrap_err().contains("文件为空"));
    }

    #[test]
    fn parse_rejects_invalid_json() {
        assert!(parse("not json at all").unwrap_err().contains("JSON"));
    }

    #[test]
    fn parse_rejects_wrong_top_level_type() {
        assert!(
            parse("{\"quote\":\"对象而非数组\"}")
                .unwrap_err()
                .contains("JSON")
        );
    }

    #[test]
    fn parse_rejects_missing_quote_field() {
        assert!(parse(r#"[{"text":"nope"}]"#).unwrap_err().contains("JSON"));
    }

    #[test]
    fn parse_filters_blank_entries() {
        let v = parse(r#"[{"quote":"保留"},{"quote":"  "},{"quote":""}]"#);
        assert_eq!(v.unwrap(), vec!["保留"]);
    }

    #[test]
    fn parse_rejects_all_blank() {
        assert!(
            parse(r#"[{"quote":""},{"quote":" "}]"#)
                .unwrap_err()
                .contains("quote")
        );
    }

    #[test]
    fn embedded_corpus_is_valid_and_large() {
        let lines = parse(EMBEDDED).expect("embedded corpus must parse");
        assert!(
            lines.len() >= 450,
            "embedded corpus shrank: {}",
            lines.len()
        );
        assert_eq!(lines[0], "梦在前方，路在脚下。");
    }

    // -- load --------------------------------------------------------------

    #[test]
    fn load_exports_embedded_when_missing() {
        let dir = temp_dir("export");
        assert!(!dir.join(FILE_NAME).exists());
        let src = load(&dir);
        assert!(matches!(src, Source::Exported(_)));
        // The file now exists and round-trips to the same corpus.
        let on_disk = fs::read_to_string(dir.join(FILE_NAME)).unwrap();
        assert_eq!(parse(&on_disk).unwrap(), src.lines().to_vec());
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn load_reads_user_file() {
        let dir = temp_dir("user");
        fs::write(dir.join(FILE_NAME), r#"[{"quote":"自定义句子"}]"#).unwrap();
        let src = load(&dir);
        assert_eq!(src.lines(), &["自定义句子".to_string()]);
        assert!(matches!(src, Source::User(_)));
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn load_falls_back_on_broken_file() {
        let dir = temp_dir("broken");
        fs::write(dir.join(FILE_NAME), "{{{ not json").unwrap();
        let src = load(&dir);
        let Source::Fallback { lines, reason } = &src else {
            panic!("expected fallback, got {src:?}");
        };
        assert!(reason.contains("JSON"), "reason: {reason}");
        assert_eq!(lines.len(), parse(EMBEDDED).unwrap().len());
        fs::remove_dir_all(&dir).ok();
    }

    #[test]
    fn load_falls_back_on_non_utf8() {
        let dir = temp_dir("gbk");
        fs::write(dir.join(FILE_NAME), [0xD6, 0xD0, 0xCE, 0xC4]).unwrap(); // GBK bytes
        let src = load(&dir);
        assert!(matches!(src, Source::Fallback { .. }));
        fs::remove_dir_all(&dir).ok();
    }

    // -- Rng ---------------------------------------------------------------

    #[test]
    fn rng_is_deterministic_per_seed() {
        let mut a = Rng::new(42);
        let mut b = Rng::new(42);
        for _ in 0..16 {
            assert_eq!(a.next_u64(), b.next_u64());
        }
        let mut c = Rng::new(43);
        assert_ne!(a.next_u64(), c.next_u64());
    }

    #[test]
    fn rng_zero_seed_is_not_degenerate() {
        let mut r = Rng::new(0);
        assert_ne!(r.next_u64(), 0);
    }

    #[test]
    fn rng_below_stays_in_range() {
        let mut r = Rng::new(7);
        for _ in 0..10_000 {
            assert!(r.below(451) < 451);
        }
        assert_eq!(r.below(0), 0);
        assert_eq!(r.below(1), 0);
    }

    #[test]
    fn rng_below_covers_range() {
        let mut r = Rng::new(0xC0FFEE);
        let mut seen = vec![false; 451];
        for _ in 0..200_000 {
            seen[r.below(451) as usize] = true;
        }
        assert!(seen.iter().all(|&s| s), "some values were never drawn");
    }

    // -- Bag ---------------------------------------------------------------

    #[test]
    fn bag_returns_a_permutation_each_round() {
        let mut rng = Rng::new(1);
        let mut bag = Bag::new(8);
        let first: Vec<usize> = (0..8).map(|_| bag.next(&mut rng)).collect();
        let mut sorted = first.clone();
        sorted.sort();
        assert_eq!(sorted, (0..8).collect::<Vec<_>>());
    }

    #[test]
    fn bag_reshuffles_between_rounds() {
        let mut rng = Rng::new(2);
        let mut bag = Bag::new(50);
        let first: Vec<usize> = (0..50).map(|_| bag.next(&mut rng)).collect();
        let second: Vec<usize> = (0..50).map(|_| bag.next(&mut rng)).collect();
        assert_ne!(first, second, "two consecutive rounds should differ");
    }

    #[test]
    fn bag_single_item_always_zero() {
        let mut rng = Rng::new(3);
        let mut bag = Bag::new(1);
        for _ in 0..10 {
            assert_eq!(bag.next(&mut rng), 0);
        }
    }

    #[test]
    fn bag_no_repeat_within_round() {
        // Within one round each index appears exactly once, so consecutive
        // draws can never repeat (rounds have length >= 2 here).
        let mut rng = Rng::new(4);
        let mut bag = Bag::new(20);
        let mut prev = bag.next(&mut rng);
        for _ in 1..20 {
            let cur = bag.next(&mut rng);
            assert_ne!(prev, cur);
            prev = cur;
        }
    }
}
