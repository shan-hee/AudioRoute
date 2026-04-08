# AudioRoute

AudioRoute is now a compact WinUI 3 audio routing panel.

## Launch

```powershell
AudioRoute.exe
```

- Launching the app opens a fixed lower-right panel.
- The panel animates in on show and animates out on dismiss.
- Press `Esc` or click outside the panel to close it.

## Current Scope

- Show active output and input sessions.
- Rebind an app session to a specific playback or recording device.
- Adjust per-session volume from the panel.
- Open as a lightweight desktop flyout instead of a tray utility.

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
