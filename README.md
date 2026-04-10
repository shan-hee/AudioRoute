# AudioRoute

轻量级 Windows 音频路由工具，基于 WinUI 3 构建。以系统托盘面板的形式运行，提供按应用的音频设备路由和音量控制功能。

## 功能特性

- **按应用音频路由** — 将任意应用的音频输出/输入重定向到指定的播放或录音设备
- **独立音量控制** — 为每个应用单独调节音量，类似 Windows 原生音量合成器
- **输出/输入双模式** — 支持播放设备（输出）和录音设备（输入）的管理
- **系统托盘常驻** — 作为托盘程序后台运行，左键点击托盘图标唤出面板，右键打开菜单
- **自动会话发现** — 实时检测并显示所有活跃的音频会话
- **轻量低开销** — 无后台服务，面板隐藏时几乎零资源占用

## 与 Windows 原生音量合成器的区别

| 功能 | Windows 原生 | AudioRoute |
|------|:---:|:---:|
| 按应用音量调节 | ✅ | ✅ |
| 按应用设备路由 | ❌ | ✅ |
| 输出/输入切换 | ❌ | ✅ |
| 当前/目标设备显示 | ❌ | ✅ |
| 托盘面板模式 | ❌ | ✅ |

## 使用方式

```powershell
AudioRoute.exe
```

- 启动后自动弹出面板，位于屏幕右下角
- 按 `Esc` 或点击面板外部区域隐藏到托盘
- 左键点击托盘图标重新打开面板
- 右键点击托盘图标打开菜单（主页 / 退出）

## 系统要求

- Windows 10 1809 (Build 17763) 或更高版本
- .NET 8.0
- Windows App SDK 1.8+

## 构建

```powershell
dotnet build -c Release
```

- 构建版本号统一读取仓库根目录的 `VERSION`
- `VERSION` 必须使用 `major.minor.patch` 格式，例如 `0.0.2`

## 自动发版

- 推送到 `main` 时，如果提交中包含 `VERSION` 变更，会自动触发 GitHub Actions 构建并发布 Release
- Release 标签格式为 `v<version>`，例如 `v0.0.2`
- Release 名称格式为 `AudioRoute v<version>`
- 发布产物为 `win-x64` ZIP 包，并附带 `.sha256` 校验文件
- 如需补发同一提交，可手动触发 `Release` 工作流

## 技术架构

```
AudioRoute/
├── App.xaml(.cs)                # 应用入口
├── MainWindow.xaml(.cs)         # 主面板窗口、托盘图标、动画逻辑
├── SessionCardControl.xaml(.cs) # 单个音频会话卡片控件
├── AudioPolicy.cs               # 核心路由引擎 — 调用 Windows 内部 AudioPolicyConfig COM 接口
├── AudioSessionService.cs       # 音频会话枚举与聚合（基于 NAudio）
├── DeviceHelper.cs              # 音频设备枚举（IMMDeviceEnumerator COM）
└── MixerModels.cs               # 数据模型与事件定义
```


## 依赖

- [Microsoft.WindowsAppSDK](https://www.nuget.org/packages/Microsoft.WindowsAppSDK) — WinUI 3 框架
- [NAudio](https://www.nuget.org/packages/NAudio) — 音频会话枚举

## 许可证

MIT
