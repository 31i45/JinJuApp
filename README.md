# JinJuApp

在任务栏旁常驻显示一行金句的 Windows 桌面小工具：每 3 分钟换一句，长句自动横向滚动，不打扰、不抢焦点、不写注册表（除非你主动开启自启动）。

v2.0.0 起为纯 Rust + Win32 原生实现（原为 Python/PyQt，见文末「为什么重写」）。

## 使用

1. 从 [Releases](https://github.com/31i45/JinJuApp/releases) 下载 `JinJuApp.exe`，放入任意可写目录；
2. 双击运行：金句条出现在任务栏旁，托盘出现图标；
3. **换内容**：托盘右键 →「编辑金句」，或直接编辑 exe 旁的 `quotes.json`；
4. **开机自启动**：托盘右键 →「开机自启动」（默认关闭，由你决定）；
5. **退出**：托盘右键 →「退出」。

无需安装，无运行时依赖，约 250 KB。

## 金句数据

`quotes.json` 位于 exe 同目录，首次运行时自动导出内置的 451 条人民日报金句：

```json
[
  { "quote": "梦在前方，路在脚下。" },
  { "quote": "每一天都是新的开始。" }
]
```

- 改文件即改内容，**不需要重新下载程序**——英语单词、待办、座右铭，放什么都行；
- 文件损坏或语法错误时，弹窗给出具体行号与原因，本次回退内置金句运行，不会静默失败；
- 显示顺序为洗牌袋随机：每一轮 451 句全部出现且不重复，轮空后重新洗牌。

## 构建

```bash
cargo build --release        # 产物 target/release/JinJuApp.exe，约 250 KB
cargo test                   # 47 个单元测试
cargo clippy --all-targets -- -D warnings   # 零警告
powershell -File e2e.ps1     # 11 项端到端冒烟检查
```

要求：Rust 1.85+（edition 2024）与 Windows SDK（资源编译器，用于嵌入图标）。二进制只由 GitHub Actions 构建发布（tag 触发，附 SHA-256 校验和），仓库不跟踪任何可执行文件。

## 工程说明

- **模块**：`quotes`（数据加载/校验/兜底 + 洗牌袋 + xorshift）、`layout`（跑马灯状态机 + 任务栏四边几何）、`render`（灰度掩码 → 预乘 ARGB 纯函数 + GDI 宿主）、`startup`（opt-in 注册表自启动）、`tray`（托盘与菜单）、`win32`（手写 FFI——本程序对系统的全部调用集中在一个文件）；
- **渲染**：per-pixel-alpha 分层窗口，GDI 白字黑底获得逐像素覆盖度，再转为预乘 ARGB 上屏；灰度抗锯齿（ClearType 在透明表面上会产生彩边）；
- **适配**：任务栏停靠任意边、DPI 变化、explorer 重启（托盘自动恢复）、单实例互斥、`Win+D` 后在下一次换句时温和恢复显示（不做每秒对抗式自愈）；
- **测试边界**：全部决策逻辑为纯函数并以单元测试覆盖；窗口/GDI/注册表等系统交互由 `e2e.ps1` 做程序化端到端验证，不为 FFI 边界做 mock。

## 为什么从 Python 重写

原 Python/PyInstaller 版本（v1.0，已从主干移除，历史保留在 git）存在结构性问题：

| 问题 | 原版 | v2.0 |
|---|---|---|
| 体积 | 35.2 MB | 248.5 KB |
| 杀软误报 | PyInstaller bootloader + UPX 壳命中 Defender 机器学习「dropper」画像（[#1](https://github.com/31i45/JinJuApp/issues/1) 报 `Wacatac.C!ml`） | 纯静态原生代码，无 bootloader、无壳、无打包器 |
| 内容更新 | 数据焊死在 exe 内，改一个字要重新发版 | 外置 `quotes.json`，用户即改即生效 |
| 自启动 | 首次运行静默写注册表，无退出途径 | 默认关闭，托盘一键开关，可退出 |
| 错误处理 | 数据错误静默吞掉，只显示兜底句 | 错误弹窗直达用户，含原因与修复路径 |
| 可测性 | 纯逻辑焊死在 QWidget 中，0 测试 | 47 单元测试 + 11 项 E2E，全绿 |

## 许可

GPL-3.0（见 [LICENSE](LICENSE)）。金句数据精选自人民日报。

> AI生成
