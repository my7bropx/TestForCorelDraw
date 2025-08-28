# CorelDRAW Vector Fill Tool

This repository contains a simple Python script that demonstrates how to
use the CorelDRAW 2025 COM automation interface to:

1. Trace a bitmap or hand-drawn image into vector lines.
2. Optionally fill the traced paths with a repeated element (such as an SVG or CDR object).
3. Control spacing and rotation of the repeated element. Spacing is applied in
   addition to the element's own width so copies do not overlap, and each copy
   is oriented to the tangent of the path plus the user-specified angle.

The script provides a small GUI built with Tkinter where you can choose
the source bitmap, the element to repeat, and configure optional
operations like tracing and filling.

## Requirements

- CorelDRAW 2025 installed with COM automation enabled.
- Python 3 with the `pywin32` package.
- Windows operating system (COM automation is Windows-only).

## Usage

```bash
python coreldraw_vector_fill.py
```

The GUI will appear allowing you to select an image and an element to
repeat. Enable or disable the tracing and filling options as desired,
choose spacing and angle, then click **Run**.

The script is intended as a starting point for building more advanced
tools; some API calls may need adjustment depending on the content of
your document or the version of CorelDRAW.
