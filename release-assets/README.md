Update `VERSION` first, then rebuild the local release payload in
`release-assets/win-x64/` before committing. The release workflow rejects a
payload whose executable or DLL file version does not match `VERSION`.

Recommended flow:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Update-ReleasePayload.ps1
```

Then verify this executable runs correctly:

```text
release-assets/win-x64/AudioRoute.exe
```

The GitHub Actions `Release` workflow will not rebuild the app. It only validates and packages that directory.
