# DualSlide 0.5.2



DualSlide runs two native PowerPoint slideshows on separate Windows displays. Operate the presenter show normally; the audience show follows its slide and animation step. An antialiased audience pointer mirrors the presenter pointer or laser.



Extract `DualSlide-0.5.2.zip` to a folder and run `DualSlide.exe`. Requires Windows 11, desktop PowerPoint, and two extended displays. No installer or .NET SDK download is required.



## Changes in 0.5.2

- Added a two-display synchronization icon to the executable, taskbar, and window. The editable vector source is `assets/DualSlide.svg`; Windows uses the derived multi-resolution `assets/DualSlide.ico`.
- To regenerate the ICO after editing the SVG, run `node tools/build-icon.cjs` with the `sharp` Node package available. Normal builds use the committed ICO and need no additional dependency.

## Changes in 0.5.1

- Fixed missing audience content for ordinary Appear/entrance animations. PowerPoint's GotoClick could update the audience click index while leaving its entrance objects invisible. Audience synchronization now plays native Next steps; backwards movement resets the audience slide and replays to the desired step.
- Catch-up is bounded to eight native steps per poll, with retries when PowerPoint is transitioning. Ordinary animation steps no longer use GotoClick.
- Added a regression that distinguishes reported animation index from rendered content, and an explicit colored-object Appear visual test.

## Changes in 0.5

- Fixed mixed-DPI cropping with the LG as primary and the internal display as audience. PowerPoint runs native windowed slideshows with borders/scrollbars removed and the canvas fitted to each display. Both slide rendering and animations remain in PowerPoint. Old maximized bounds and MDI canvas padding are cleared, and delayed DPI placement changes are corrected.
- Each slideshow frame is obtained from its exact PowerPoint COM window, rather than guessing from global window creation or foreground focus.
- Input hooks run on their own message-pump thread. The pointer has no cross-process owner. Presentation windows are not forced always on top, and synchronization cannot recursively enter its timer.
- A 15-second stall watchdog removes input hooks, hides the replicated pointer, requests asynchronous minimization of the paired windows, and requests cancellation of the blocked COM call. Coordinated no-save cleanup runs when control returns. It never kills PowerPoint or closes unrelated documents. A hard Office/driver hang is not guaranteed to accept cancellation or minimization.
- Only one controller can run per Windows session, including copies in other folders. A second launch exits and asks the existing window to activate.
- Diagnostic startup/error/recovery records are written to `%LOCALAPPDATA%\DualSlide\DualSlide.log`.

## Changes in 0.4

- Escape is consumed on either paired slideshow and routes through one coordinated shutdown. Ending either native slideshow also closes its paired show, even while paused or after focus moves away.
- Tool-owned presentations are marked for discard before exiting their shows and closed without saving. Minimized read-only editor windows prevent PowerPoint's automatic hidden-document close/save path. Busy close requests retry while the original-file write locks remain held.
- Originals are opened directly read-only with their actual filenames. No presentation copies are created. A read-only file handle denies writes, replacement, and deletion for the session. Files already open for editing must be closed before starting; normally opened read-only presentations can be reused without discarding their existing unsaved edits.
- Added the copyright notice and GitHub link to the window footer and copyright metadata to the executable.
- Pointer shutdown no longer waits synchronously on the renderer, and pointer motion does not repeatedly reorder the foreign PowerPoint owner.

## Changes in 0.3



- Replaced the missing laser overlay with a WPF transparent surface owned by the audience slideshow. Its independent render thread handles alpha blending, cursor sampling, and monitor mapping.

- Wheel steps explicitly call Next/Previous on the presenter view, then synchronize the audience. The original wheel message is consumed while aimed at the active presenter, avoiding independent follower wheel actions. Other PowerPoint navigation remains native.

- Added recent-file dropdowns and paired selection in either direction. The latest used pairing wins. History and display choices persist under `%LOCALAPPDATA%\DualSlide`, including across updates or running from a different folder.

- Display selectors automatically resolve collisions in either direction.

- Refreshed the setup layout, rounded selectors and buttons, light/dark colors, and Windows accent color. Buttons clear their corners instead of leaving native border artifacts.



## Architecture and scope



Two slideshow windows run in one PowerPoint application. PowerPoint renders the decks and animations. DualSlide compares native slide/build state every 50 ms and retries transient busy errors. The pointer runs independently at approximately 60 Hz. Timing is not frame-locked.



Decks must have matching technical structure: slide counts, aspect ratio, and animation logic. Ordinary click/arrow animation sequences are supported. Shape-specific triggers, ink mirroring, and synchronized audio/video clocks are outside scope.



Files are opened directly read-only, retaining their filenames. Already-open read-only presentations can be reused and remain open after Stop, including unsaved edits. Presentations open for editing are refused so original-file write protection is maintained. DualSlide never saves either presentation.



## Build and verification



Run `build.ps1`. It uses the Windows .NET Framework compiler and WPF assemblies already present on Windows. Output: `dist/DualSlide.exe`.



Sources: `DualSlide.cs`, `PointerRenderer.cs`, `ModernWindows.cs`, `DeckHistory.cs`, `SessionSafety.cs`.



`test-build.ps1` builds an individual diagnostic harness. Current checks are `ShutdownTest.cs`, `SessionTest.cs` (optional argument swaps displays), `SafetyTest.cs`, `OwnershipTest.cs`, and `FeatureCheck.cs`; `VisualSessionTest.cs` adds disposable in-memory corner markers for interactive geometry/laser inspection. Tests isolate their preferences from the application's normal settings. `AnimationReplayTest.cs` checks animation replay without opening decks; `AppearVisualTest.cs` creates its own disposable test decks. Tests that reference `test-master.pptx` and `test-follower.pptx` require a local pair of matching test presentations; presentation files and generated diagnostics are excluded from Git.



See `USER_GUIDE.md` for controls and `VERIFICATION.md` for evidence and limits. Development decks and harnesses are not required by the packaged application.


Copyright (C) 2026 Silviu Vlasceanu. Project: https://github.com/silviuk/DualSlide
