# DualSlide development

## Build from source

Run these commands in PowerShell from the repository root:

```powershell
.\build.ps1
```

The executable is written to `dist/DualSlide.exe`. The build uses the Windows .NET Framework C# compiler and WPF assemblies at `%WINDIR%\Microsoft.NET\Framework64\v4.0.30319`; no NuGet restore or Office interop package is required. PowerPoint is needed to run a presentation, but not to compile the app.

To build a distributable ZIP and SHA-256 checksum:

```powershell
.\package.ps1
```

This rebuilds the executable and produces `dist/DualSlide-0.5.4.zip` and its `.sha256` file, including documentation and icon assets. The version is read from the executable metadata.

### Icon

The editable icon source is [assets/DualSlide.svg](../assets/DualSlide.svg). Windows uses the derived, multi-resolution `assets/DualSlide.ico` for the executable and window. Both are committed, so normal builds need no image tools. After editing the SVG, regenerate the ICO with Node.js and the `sharp` package available:

```powershell
node tools/build-icon.cjs
```

## Automated builds and releases

[GitHub Actions](https://github.com/silviuk/DualSlide/actions) builds on pushes to `main`, pull requests, version tags, and manual dispatch. The Windows runner compiles the app, runs the animation replay regression without PowerPoint, and uploads the ZIP and checksum as workflow artifacts.

Pushing a tag such as `v0.5.4` also publishes a GitHub release with those artifacts after the build and regression pass. The workflow rejects tags that do not match the executable version. Release publishing uses the repository's `GITHUB_TOKEN`; no personal access token is required.

## Verification and limits

Run the animation replay regression without opening presentations:

```powershell
New-Item -ItemType Directory -Force .build | Out-Null
.\test-build.ps1 -Source AnimationReplayTest.cs -Main AnimationReplayTest -Output .build\AnimationReplayTest.exe
& .\.build\AnimationReplayTest.exe
if ($LASTEXITCODE -ne 0) { throw 'Animation replay regression failed' }
```

PowerPoint and two displays are required for the interactive tests documented in [VERIFICATION.md](../VERIFICATION.md). `AppearVisualTest.cs` creates disposable test decks. Other session tests use a local matching pair named `test-master.pptx` and `test-follower.pptx`; local presentations and diagnostic outputs are excluded from Git.

Synchronization follows PowerPoint's slide/build state and is not frame-locked. Shape-specific animation triggers, ink mirroring, and synchronized audio/video clocks are outside scope. A hard Office or graphics-driver hang may not accept the watchdog's cancellation or minimization requests.

## Source layout

| File | Responsibility |
| --- | --- |
| `DualSlide.cs` | Setup window, PowerPoint sessions, navigation, and animation synchronization |
| `PointerRenderer.cs` | Native cursor copying, WPF laser rendering, and display mapping |
| `ModernWindows.cs` | Theme, accent color, display discovery, and identification |
| `DeckHistory.cs` | Recent presentations and paired selections |
| `SessionSafety.cs` | Single-instance handling, input hooks, recovery, and logging |
