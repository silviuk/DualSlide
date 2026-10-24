# Version 0.5.4 (2026-10-24)

- End-of-deck handling retains both live shows on the final slide. Forward keyboard commands and wheel steps are guarded before calling native Next; terminal clicks are suppressed. A native done/end screen returns to Last without resetting its completed animations. Escape still uses coordinated no-save shutdown.
- AnimationReplayTest passed its new final-build clamp, end-screen recovery, and backward-navigation cases.
- EndOfDeckTest passed with two disposable PowerPoint decks, each containing two slides and two Appear builds per slide: retained final build 2, repeated forward clamp, actual native done-state recovery, backward/forward navigation, and Escape exit. Test decks use a unique .build/end-deck-* directory.

# Version 0.5.3 (2026-10-24)

- Build and AnimationReplayTest passed.
- PointerRegressionTest passed: native cursor transparency and hotspot, 1,000 bitmap conversions without GDI handle growth, and rendered laser ring/center/halo assertions. Uses the system Hand cursor because desktop capture tooling may replace the system Arrow with transparent artwork.
- CleanupRegressionTest passed: owned empty application quits, pre-existing application and newly opened user documents are preserved, busy Quit retries, repeated cleanup is safe.
- PowerPointLifecycleTest passed on desktop PowerPoint with two displays: generated disposable decks in a separate process, started DualSlide from a closed PowerPoint baseline, stopped both shows, and verified POWERPNT exited while the controller process remained alive.
- Existing ShutdownTest passed all five exit paths, source hashes/timestamps and write-lock checks, unrelated-document preservation, and busy-close retry. OwnershipTest passed borrowed read-only document/unsaved-edit preservation.
- Inspected PowerPoint's native red laser on a disposable white slide. GetCursorInfo reports a null cursor while the laser is visibly drawn in the slide surface. Normal cursors therefore use native bitmap copying; the laser remains a separately drawn ring/halo. No claim of pixel-identical laser matching or custom-color mirroring is made.

Run the new Office-free regressions with test-build.ps1 and the corresponding test main classes. PowerPointLifecycleTest requires PowerPoint closed and two displays, and creates its own unique .build/lifecycle-* folder for disposable decks.

# Version 0.5.2 (2026-10-24)

- Built the executable with the SVG-derived native Windows icon and embedded window icon. Both resources load successfully through System.Drawing.Icon.
- Checked the SVG rendering and all nine ICO sizes: 16, 20, 24, 32, 40, 48, 64, 128, and 256 pixels.
- Rebuilt the application and ran AnimationReplayTest successfully: forward/backward rendering, reset, bounded catch-up, slide changes, and transient busy retry.
- Added a repeatable ZIP/checksum packaging script and Windows build/release workflow. Package contents were checked locally. PowerPoint synchronization behavior is unchanged.

# Verification — DualSlide 0.5.1, 2026-10-24

## Missing audience Appear content

Reproduced on generated matching decks with three high-contrast Appear entrance objects. In 0.5, native presenter Next revealed the red APPEAR 1 rectangle. The audience reported the same slide/build 1/1 but remained blank apart from its static title. This was directly observed in both PowerPoint windows. The older index-only checks did not establish that animated audience content was rendered.

The adapter now uses native View.Next for ordinary audience animation playback. Backward movement resets the audience slide and replays to the target click, with an eight-step per-poll budget and transition retry. GotoClick remains only for special negative automatic-animation states.

After the fix, the first audience red Appear rectangle was visibly rendered at the matching step. A separate three-step run visibly showed all red, green, and blue objects on the audience screen at 1/3. Pressing native Left on the presenter returned both shows to 1/2; the audience visibly retained red/green and hid blue. Presenter Escape closed the paired test without save prompts. The display geometry, file protection, singleton, input thread, and recovery implementation are retained.

- `animation-replay-results.txt`: regression with separate reported/rendered state, including forward/backward steps, reset, slide jump, bounded catch-up, and busy retry. It deliberately models GotoClick changing the index without revealing content.
- `appear-before-results.txt`: failing 0.5 reproduction; equal indices did not imply equal visible content.
- `appear-fixed-results.txt`: fixed first-step visual session.
- `appear-all-results.txt`: three-step native playback and actual Left-key backward navigation; the rendered audience was visually checked at steps 3 and 2.
- `session-051-results.txt`, `session-051-swapped-results.txt`: native integration regression in both display assignments.
- `shutdown-051-results.txt`: original-file hash/timestamp protection and coordinated no-save shutdown.

## Historical 0.5 evidence and limitations

The earlier animation checks below compared indices. The 0.5.1 reproduction supersedes any implication that those index checks proved visible entrance content.

# Verification — DualSlide 0.5, 2026-10-24

## Mixed-DPI display rendering

The paired speaker-mode fullscreen test reproduced clipping on the internal monitor with the LG primary, even with correct native window bounds. Resizing through native calls, changing activation order, and COM width/height recalculation were insufficient in the paired case.

The final approach uses PowerPoint's native windowed slideshow with scrollbars/frame removed and full physical display bounds. Old maximized/minimized styles and MDI canvas frame padding are cleared; delayed placement changes are corrected. Both the outer frame and actual slide canvas client bounds are asserted in regression tests. No Windows display or Office global preferences were changed.

In a paired visual run with LG primary/presenter and internal audience, all four red in-memory corner markers were visible around the complete slide. Black bands were centered for the 16:9 slide on the 3:2 internal screen. The replicated red antialiased laser was visibly present within the audience slide. Native presenter Escape closed both shows and the test application; no save dialog remained.

## Freeze protection and singleton

The prior user's exact crash has no saved dump, so its precise origin is not proven. Removed two risky dependencies: low-level input hooks sharing the Office COM/UI thread, and the WPF pointer owning/being owned across PowerPoint process windows. The controller timer also rejects recursive entry. Neither presentation window is forced always on top.

`SafetyTest.cs` starts both native shows and deliberately blocks the controller thread. A shortened watchdog timeout verifies independent recovery: both windows minimize while the controller is blocked, hooks are removed, then the no-save cleanup closes both shows when pumping resumes. Production uses 15 seconds. The test also starts a competing process against the same mutex and verifies exclusion and release for a later launch.

Recovery requests COM cancellation and asynchronous minimization; a hard Office/graphics-driver hang may refuse either. The tool never automatically kills PowerPoint. Startup, errors, and recovery records go to `%LOCALAPPDATA%\DualSlide\DualSlide.log`.

## Current regression evidence

- `session-05-results.txt` and `session-05-swapped-results.txt`: manifest-bearing integration checks in both display-role assignments.
- `shutdown-05-results.txt`: five coordinated exit paths, including presenter/audience End Show, paused exit, busy-close retry, and Escape callback suppression. Both source files stayed locked during the session, and hashes/timestamps were unchanged after cleanup. An unrelated presentation remained open.
- `ownership-05-results.txt`: original paths used directly, existing read-only presenter reused, unsaved edits retained, only the owned deck closed, source hashes unchanged.
- `safety-05-results.txt`: independent watchdog recovery and competing-process singleton exclusion.
- `singleton-release-05-results.txt`: launched the actual release executable twice; the second process exited and one controller remained. Its visible window was verified as DualSlide 0.5.
- `visual-05-results.txt`: interactive corner-marker/laser session; screenshots were inspected through Computer Use. This log's visibility flags alone are not the visual evidence.
- Actual wheel-down plus native Right advanced both shows to slide/build 2/0. Wheel-up returned both to 1/1. `input-05-results.txt` records the two matching native states and exact frame/canvas bounds after reverse navigation. Presenter Escape then closed both shows with no save dialog.

Exploratory runs before the final implementation failed window discovery, retained sample decks after harness assertions, and exposed delayed maximized-window geometry. Those failures were corrected and the read-only sample leftovers discarded explicitly. Final results above supersede those intermediate runs.

## Historical 0.4 evidence



## Shutdown and original-file protection

- `shutdown-04-results.txt`: five shutdown paths passed with unsaved in-memory changes deliberately introduced in both read-only sample decks: native presenter End Show, native audience End Show, presenter End Show while paused, a rejected close call followed by retry, and the actual Escape-hook callback.
- The Escape-hook check asserted a nonzero suppression result, so its keydown cannot continue to PowerPoint's independent close path. This was a callback integration test; final physical-key UI automation was not completed after user input was detected in the release window.
- Each path verified both slideshow windows and both tool-owned presentations closed, with unrelated presentations preserved.
- Original paths were used directly and both PowerPoint ReadOnly flags were true. Opening either original for writing was denied during the session. SHA-256 hashes and last-write timestamps stayed unchanged; the file locks were released after cleanup.
- `ownership-04-results.txt`: an already-open read-only presenter and its unsaved edits were preserved, with no extra presentation copy and unchanged source hashes.
- `session-04-results.txt`: 12 native paired slide/build advances, blackout/whiteout, backward navigation, geometry, pointer-thread independence/hide, and transient-call recovery passed.
- `features-04-results.txt`: paired history, stable focus validation, monitor separation, restart restoration, and light/dark button rendering passed.
- Visually inspected the actual DualSlide 0.4 release footer: copyright notice and accent-colored github.com/silviuk/DualSlide link are present and fit the window.

One initial broad regression hung PowerPoint. That test run was cancelled, its sample-only PowerPoint process was closed with approval, and the pointer's synchronous shutdown wait/repeated foreign-owner reordering were removed. The subsequent broad regression and shutdown checks passed. The record does not claim that every possible Office hang can be recovered automatically.

The original files are never saved, copied, renamed, or assigned a read-only filesystem attribute by the tool. Read-only opening plus an OS write-denying file handle provides protection. An original already open for editing must be closed before starting. Existing normally opened read-only documents are borrowed and preserved; only tool-owned documents are discarded and closed.

The application copyright is (C) 2026 Silviu Vlasceanu. The footer's destination is exactly https://github.com/silviuk/DualSlide; no license or publishing action was added.

## Earlier 0.3 evidence

## Laser and input



The previous release reported its overlay as visible without demonstrating actual on-screen laser rendering. The user confirmed that it was missing. The native bitmap upload path has been replaced with a WPF transparent window, a separate dispatcher thread, and explicit audience-window ownership.



- Visually observed a red antialiased laser on the actual audience slideshow with the new renderer. This is actual PowerPoint-window capture evidence, not just a Visible flag.

- Repeated with presenter/audience monitors swapped: the laser remained visible. Exact PowerPoint rendering geometry in the swapped high-DPI window capture was inconclusive; this record does not claim a complete physical-screen scaling measurement for that arrangement.

- Tested the actual manifest-bearing release executable, title DualSlide 0.3: Start with laser pointer produced a visible audience laser. Presenter slide center mapped to audience center (1920,1080). Escape stopped both shows cleanly.

- Ctrl+L / Ctrl+A laser selection and native N navigation were exercised. N revealed the corresponding audience animation in the actual release.

- `fluent-live-results.txt`: real wheel down advanced both shows from slide/build 1/0 to 1/1; wheel up returned both to 1/0; two wheel-down steps advanced both to 2/0. Presenter foreground focus was retained, and renderer error stayed empty.

- The original wheel message is consumed only when aimed at the active presenter slideshow. Navigation is queued through that presenter's native View.Next/Previous. The follower is updated from presenter state, avoiding a second independent native wheel action.



## Regression and ownership



- `fluent-session-results.txt`: 12 paired slide/build advances, black/white modes, backward navigation, display geometry, independent pointer thread, hide, and recovery from an injected rejected COM call passed.

- `fluent-ownership-results.txt`: an existing normally opened presenter was reused without an untitled copy; its unsaved shape edit and dirty flag survived Stop; only the owned audience deck closed; source SHA-256 hashes remained unchanged.

- PowerPoint version tested: 16.0.17932.21056. Displays have differing resolution/scaling: 3120x2080 at (3840,70), and 3840x2160 at (0,0).



## Recent files, displays, and appearance



- `fluent-features-final-results.txt`: persisted recent pairs, newest-partner priority, reverse lookup, both dropdown associations, focus-only validation, automatic display collision resolution in both directions, and restoration after restart passed.

- Button corner pixels match the parent surface in both light and dark mode; the native Button border is disabled. Rounded combo-box frames and chevrons use the current palette.

- Visually inspected the actual release's light interface: grouped cards, clean button corners, rounded selectors, and orange primary/focus accent matching Windows (RGB 231,92,11).

- A dark interface preview was inspected. Tests also rendered both theme states. Theme/accent notifications re-read Windows preferences; Windows preferences are not modified.

- Monitor labels and Identify displays retain the previously tested QueryDisplayConfig / DisplayConfigGetDeviceInfo implementation.



## Limits



Ordinary native navigation and matching technical decks are the supported scope. State polling is 50 ms and pointer sampling about 60 Hz; neither provides a synchronized media clock or frame lock. Rapid input can bypass intermediate follower animation frames. Interactive shape triggers and ink remain excluded. Physical remote hardware and other Office versions were not tested.



The zip contains the release executable and documentation. Test preferences are isolated under .build; production preferences use the local Windows user profile.
