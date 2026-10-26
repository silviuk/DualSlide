# DualSlide 0.5.5



## Start



1. Configure Windows displays to Extend.

2. Extract the archive and open DualSlide.exe. The title should say DualSlide 0.5.2.

3. Choose the presenter and audience decks with Browse or the recent-file dropdowns. They must have matching slide counts, slide format, and animation logic.

4. Select the presenter and audience displays. If a selection conflicts, the other selector automatically chooses a different display. Identify displays shows their Windows numbers for three seconds.

5. Close any selected deck that is already open for editing in PowerPoint. DualSlide locks the originals against writes and opens them directly read-only.
6. Optionally select Start with laser pointer, then Start presentation.



The setup window can start on either monitor. The selected display roles determine placement. The shows use PowerPoint's native windowed slideshow, made borderless to fill the selected screens; this avoids PowerPoint's mixed-DPI fullscreen cropping. Letterboxing is normal when a display and slide have different aspect ratios.

Only one DualSlide instance runs in a Windows session. Launching it again brings the existing controller or active presenter forward.



## Recent files



The app remembers the latest 20 deck pairs after successful starts. Choosing a recent deck on either side restores its most recently used partner on the other side. You can then browse for a different partner; the next successful start remembers that pairing. Merely moving focus between fields does not change the selected pair.



The last file and display selections are restored on restart. Preferences are stored in your local Windows user profile, so extracting an updated build elsewhere does not discard them.



## Present normally



- Operate the presenter slideshow throughout. PowerPoint renders the slides and animations.

- Clicks, arrows, Space, Enter, Page Up/Down, N/P, and other native PowerPoint slide-navigation commands follow through to the audience. A keyboard-based presentation remote uses the same native path.

- Wheel down advances one PowerPoint animation/slide step; wheel up goes back. The wheel is routed explicitly to the active presenter view while the cursor is over its slideshow.

- Move the mouse over the presenter slide to show the corresponding audience pointer. The normal pointer uses the native cursor image and hotspot. The red laser uses a small ring and soft halo matching PowerPoint's style; it is independently rendered and is not a pixel-identical copy of Office's laser. Custom laser colors are not currently mirrored.

- Ctrl+L selects the laser; Ctrl+A returns to an arrow. Ctrl+left-drag also mirrors a temporary laser gesture.

- B and W mirror black/white screen modes.

- Escape on either slideshow stops both shows without saving or asking to save. Ending either show through PowerPoint's End Show command also stops the paired session, including while synchronization is paused.

- Reaching the end of the deck keeps both shows on the last slide with its completed animations. Further ordinary forward keys, clicks, or wheel steps do not end the session. You can go backward and forward again; press Escape when you are ready to exit.



Switch to the setup window to Pause, Resume, Reset current slide, or Stop. Reset restarts both shows at the beginning of the presenter's current slide. Stop closes the slideshows and only the presentations opened by DualSlide. Normally opened read-only presentations you already had open stay open with their existing unsaved edits preserved. No presentation is saved by the tool, and no presentation copies are created. The original files remain locked against writes until cleanup finishes.



The setup interface follows Windows light/dark and accent-color changes. File and display selection are locked while a session is running.



## Supported workflow



Windows 11, desktop PowerPoint, extended displays, and corresponding ordinary click/arrow animation sequences, including With Previous and After Previous chains. Two native slideshow windows run in one PowerPoint application.



Synchronization follows native slide/animation state every 50 ms and retries temporary PowerPoint busy calls automatically. It is not frame-locked. Rehearse with the actual decks before presenting. Interactive shape triggers, ink mirroring, and synchronized media clocks remain outside scope.


Copyright (C) 2026 Silviu Vlasceanu. The footer links to https://github.com/silviuk/DualSlide.

## Recovery

Input handling runs independently of PowerPoint. If the controller stops responding for 15 seconds, recovery removes its input hooks, hides its pointer, requests that both slideshow windows minimize, and requests cancellation of the stalled Office call. Once control returns, both shows close through the normal no-save path. Original-file write guards remain held until cleanup completes.

Presentation windows are ordinary windows, so other applications can cover them. On Stop, DualSlide quits PowerPoint if it started the process and no presentations remain. A pre-existing PowerPoint process or any remaining presentations are preserved. DualSlide never force-kills PowerPoint. A hard Office or graphics-driver hang may reject cancellation/minimization; retain `%LOCALAPPDATA%\DualSlide\DualSlide.log` when reporting a problem.

## Keep presentations open for editing

Enable **Keep presentations open for editing after slideshow** before starting. After Escape, Stop, or ending either show, DualSlide closes both read-only sessions without saving, releases its write locks, then reopens app-owned original decks as editable PowerPoint presentations. DualSlide never modifies or saves those reopened decks. You can edit and save them yourself. This option is off by default and remembered. Existing borrowed read-only decks stay open in their existing mode so unsaved edits are preserved. Failed starts still clean up app-owned decks.
