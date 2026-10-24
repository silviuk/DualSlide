# DualSlide 0.5.3

- Replaced the hand-drawn audience arrow with the presenter's native cursor image, including transparency, hotspot, and proportional DPI scaling.
- Changed the audience laser from a solid dot with an offset highlight to a centered ring with a soft halo. This matches PowerPoint's red-laser style, but remains an independent approximation; custom Office laser colors are not mirrored.
- Quit an empty PowerPoint process started by DualSlide after both shows close. Preserve pre-existing PowerPoint sessions and any remaining presentations; retry busy Quit calls and release Office COM references.
- Added cursor/resource and application-ownership regressions to CI, plus a disposable PowerPoint process-lifecycle integration test.

DualSlide presents two matching PowerPoint decks on separate Windows displays, with synchronized slide navigation, ordinary animation steps, and pointer movements.

- Native PowerPoint rendering, including Appear animations, with replay when navigating backwards.
- Antialiased replicated laser pointer and support for displays with different scaling settings.
- Windows light/dark theme and accent color, display names and identification, paired recent-file selections, and a new SVG app icon.
- Original presentations opened directly read-only, coordinated no-save shutdown, single-instance handling, and recovery for stalled PowerPoint calls.
- Project-focused documentation and automated Windows builds, regression checks, release packaging, and SHA-256 checksums.

Download and extract `DualSlide-0.5.3.zip`, then run `DualSlide.exe`. Requires Windows 11 x64, desktop PowerPoint, and two extended displays. Decks must have matching slide and animation structure.

The executable is not code-signed. See the README and user guide in the ZIP for usage and limitations.
