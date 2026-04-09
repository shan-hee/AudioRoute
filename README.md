# AudioRoute

AudioRoute is now a compact WinUI 3 audio routing panel with tray mode.

## Launch

```powershell
AudioRoute.exe
```

- Launching the app opens a fixed lower-right panel.
- The app creates a tray icon and keeps running in the background.
- The panel animates in on show and animates out when hidden to tray.
- Press `Esc` or click outside the panel to hide it to tray.
- Left-click the tray icon to open the panel.
- Right-click the tray icon to open a menu with `主页` and `退出`.

## Current Scope

- Show active output and input sessions.
- Rebind an app session to a specific playback or recording device.
- Adjust per-session volume from the panel.
- Run as a lightweight tray utility with a desktop flyout panel.

## Project Layout

```text
AudioRoute/
  App.xaml
  App.xaml.cs
  MainWindow.xaml
  MainWindow.xaml.cs
  SessionCardControl.xaml
  SessionCardControl.xaml.cs
  AudioPolicy.cs
  AudioSessionService.cs
  DeviceHelper.cs
  MixerModels.cs
```
