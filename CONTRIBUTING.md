# Contributing

Contributions should keep the toolkit original and independent of proprietary products. Prefer focused standard VBA modules with `Option Explicit`, clear public `AET_` macro names, defensive validation, and conservative drawing modifications.

Before submitting a change:

1. Import all modules into a supported AutoCAD VBA project and run **Debug > Compile**.
2. Exercise affected macros in a disposable drawing using documented drawing units and UCS.
3. Update the README command list, limitations, and manual verification steps when behavior changes.
4. Do not commit a generated `.dvb`, drawings with user data, or unrelated build artifacts.

Include the AutoCAD version, VBA Enabler version, units, steps to reproduce, and expected/actual results when reporting a defect.
