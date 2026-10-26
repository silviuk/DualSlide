# DualSlide 0.5.5

- Added **Keep presentations open for editing after slideshow**, off by default and remembered across launches.
- When enabled, ending either slideshow, pressing Escape, or selecting Stop closes both read-only sessions without saving, releases the original-file write locks, and reopens app-owned decks for editing in PowerPoint.
- DualSlide never saves either file. Any subsequent edits and saves are controlled by you in PowerPoint.
- Presentations already open read-only remain in their existing mode, preserving unsaved edits. Failed starts still clean up newly opened decks.
- Added an automated keep-open regression to Windows CI, alongside the animation replay regression.

Download and extract `DualSlide-0.5.5.zip`, then run `DualSlide.exe`. Requires Windows 11 x64, desktop PowerPoint, and two extended displays.
