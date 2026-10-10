// Embed the application icon into the PE resource section at build time
// (the same channel the shell/taskbar reads; no runtime extraction).
fn main() {
    if std::env::var("CARGO_CFG_TARGET_OS").unwrap_or_default() == "windows" {
        winres::WindowsResource::new()
            .set_icon("logo.ico")
            .compile()
            .expect("winres: failed to embed icon");
    }
}
