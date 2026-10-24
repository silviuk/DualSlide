# DualSlide implementation and maintenance plan — 0.5.1

The implementation is delivered for two technically matching decks with ordinary click/arrow animation sequences. PowerPoint owns slide rendering and animations. Content can differ for any reason.

## Architecture

- `DualSlide.cs`: WinForms setup, read-only PowerPoint COM integration, paired native slideshow windows, display placement, and slide/build synchronization.
- `SessionSafety.cs`: exact Office slideshow HWND interface, singleton mutex/activation message, input message-pump thread, watchdog, and local diagnostic log.
- `PointerRenderer.cs`: antialiased WPF laser/arrow on a separate STA dispatcher, independent of COM. Its transparent click-through nonactivating window has **no foreign HWND owner**. Do not attach it to PowerPoint: ownership couples window lifetimes and input queues across processes.
- `ModernWindows.cs`: light/dark appearance, Windows accent color, monitor names from QueryDisplayConfig, and Identify overlays.
- `DeckHistory.cs`: 20 recent pairs, bidirectional association, and atomic local preferences.

Two native show windows share one PowerPoint application. Background command forwarding previously caused focus and delivery failures; keep the native state adapter for ordinary animation navigation.

## Display placement

Use `ShowType=2` (PowerPoint windowed slideshow) and `ShowScrollbar=0`. Retrieve the exact show HWND through the stable SlideShowWindow dual-interface vtable, then its top-level frame through GetAncestor. Validate that it is a PowerPoint frame before modifying styles.

Restore the frame before stripping caption/frame, minimize/maximize styles, and applying borderless popup styling. Disable DWM rounded corners/border and fit its client rectangle to the selected display's physical bounds. Fit the native MDI slideshow canvas as well as its outer frame; strip leftover MDI frame/maximized styles so the actual canvas client rectangle fills the selected screen. Correct delayed PowerPoint/DPI placement changes without activating windows. Skip corrections while minimized. Resolve selections by Windows device name, not enumeration order. Neither show is forced always on top.

Speaker-mode fullscreen windows retained a clipped renderer after cross-DPI moves even when HWND bounds were correct. Native resize and COM width/height adjustments did not reliably fix the paired case. Borderless windowed slideshows visibly render all four slide corners on the internal display with the LG primary. Do not replace this with a geometry-only fullscreen fix.

## Synchronization and pointer

Let native PowerPoint handle presenter clicks and keyboard navigation. The dedicated hook thread observes keyboard changes and posts work; it must never call COM or wait for the controller. Consume Escape on either paired show and enqueue coordinated Stop. Wheel input over the active presenter is accumulated in 120-unit steps on the controller thread, applied through presenter View.Next/Previous, then synchronized.

Poll the active presenter's slide index, click index, and state every 50 ms. Re-read its slide index before applying the snapshot. Move the audience slide, then play its native View.Next steps to the presenter click index, then match black/white state. Do not use GotoClick for ordinary steps: in an inactive windowed show it can update the index without rendering entrance objects. To move backwards, reset the audience slide and replay native Next steps. Limit replay/catch-up to eight steps per poll and stop the loop if the click index has not advanced; retry on a later poll. Retry transient rejected COM calls. A guard prevents recursive timer processing while Office pumps messages. Pausing suspends mirroring but must still detect closed shows.

The pointer samples physical cursor coordinates on its own dispatcher at about 60 Hz. Map normalized positions between letterboxed slide rectangles. Cache artwork and positions. Configure its state from cached controller geometry/laser state; never call PowerPoint from the render thread. EndSession posts asynchronously, without waiting for the WPF dispatcher.

## File safety and lifecycle

Open the original paths read-only with Untitled=false and real editor windows. Hold FileAccess.Read/FileShare.Read guards for both source files throughout the session and pending cleanup: no writes, replacement, deletion, copying, renaming, or filesystem read-only attribute changes. Refuse originals already open for editing. Borrow normally opened read-only presentations and preserve their unsaved edits.

Cache borrowed ShowType, ShowPresenterView, and ShowScrollbar. Restore only settings actually configured by DualSlide. Mark **only owned** decks Saved=true before View.Exit and again before Close. Saved=true clears an in-memory dirty flag; it does not save. Close only owned presentations, never quit the shared PowerPoint application. Keep failed close references and guards for retries. Form closure waits for coordinated cleanup.

The watchdog runs independently of COM and the UI message pump. After 15 seconds without a completed controller heartbeat, remove hooks, hide the pointer, request asynchronous minimization, request COM cancellation, and enqueue no-save Stop. Tag tracked windows with a session property to reject recycled HWNDs. Never kill PowerPoint or unrelated documents. Cancellation/minimization cannot be guaranteed against a hard Office/driver hang; log diagnostics honestly.

Acquire one named Local mutex before constructing the main form. Keep it for the process lifetime, recover abandoned ownership, and post an activation broadcast on a second launch. An existing controller restores its setup window or activates its presenter.

## Verification and handoff

Build with `build.ps1` using the installed .NET Framework compiler; no SDK or external package download is needed. `test-build.ps1` includes the same DPI manifest and sources.

- `AnimationReplayTest.cs`: distinct logical/rendered state, forward/backward Appear, reset, bounded catch-up, slide jumps, and rejected-call recovery.
- `AppearVisualTest.cs`: generates new matching decks with three explicitly colored Appear objects on each slide, then runs the real adapter; verify visible shapes, not only indices.
- `SessionTest.cs`: 12 paired advances, native animation builds, backwards navigation, blackout/whiteout, both display assignments (frame and actual slide-canvas bounds), pointer-thread independence, and injected rejected-call recovery. Pass any argument to swap displays.
- `ShutdownTest.cs`: five exit paths with intentional in-memory edits, including busy-close retry and Escape callback suppression; hashes/timestamps, write denial/release, and unrelated-document preservation.
- `OwnershipTest.cs`: borrowed read-only presenter and unsaved edits survive Stop, no duplicate documents, source hashes unchanged.
- `SafetyTest.cs`: competing process excluded, singleton released, deliberately blocked controller with independent watchdog recovery, subsequent no-save cleanup, and hook removal.
- `VisualSessionTest.cs`: adds four disposable corner markers to each sample slide, uses LG primary/presenter and internal audience, and starts the laser. Verify all corners, aspect-correct letterboxing, laser visibility/mapping, actual wheel/keyboard input, and Escape. Its timeout/Stop discards the markers.
- `FeatureCheck.cs`: paired recent-file selections, monitor separation, persistence, and appearance.

Failed exploratory harnesses can leave read-only sample decks open. `CleanupTestDecks.cs` discards only the exact workspace test-master/test-follower paths after checking ReadOnly; it must never close arbitrary presentations. Other older probes are historical diagnostics.

Update README, USER_GUIDE, and VERIFICATION with the final behavior and evidence. Package the manifest-bearing executable and docs as DualSlide-0.5.2.zip. Recheck actual rendering; flags and HWND bounds alone are insufficient evidence. Keep the copyright © 2026 Silviu Vlasceanu and footer link https://github.com/silviuk/DualSlide.
