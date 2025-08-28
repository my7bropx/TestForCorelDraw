import math
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import win32com.client as win32
except ImportError:
    win32 = None

class CorelVectorFillTool(tk.Tk):
    """Simple GUI wrapper for CorelDRAW automation.

    The tool demonstrates how to
      1. Trace a bitmap into vector lines.
      2. Optionally duplicate an element along those lines.
      3. Allow the user to specify spacing and rotation.

    It requires CorelDRAW 2025 with COM automation enabled and the
    `pywin32` package installed.
    """
    def __init__(self):
        super().__init__()
        self.title("CorelDRAW Vector Fill Tool")
        self.geometry("420x260")

        self.do_trace = tk.BooleanVar(value=True)
        self.do_fill = tk.BooleanVar(value=True)

        # Input for bitmap
        tk.Label(self, text="Bitmap:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.bmp_path = tk.Entry(self, width=40)
        self.bmp_path.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self, text="Browse", command=self._choose_bitmap).grid(row=0, column=2, padx=5, pady=5)

        # Input for element
        tk.Label(self, text="Element to repeat:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.elm_path = tk.Entry(self, width=40)
        self.elm_path.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(self, text="Browse", command=self._choose_element).grid(row=1, column=2, padx=5, pady=5)

        # Options
        tk.Checkbutton(self, text="Trace bitmap to curves", variable=self.do_trace).grid(row=2, column=0, columnspan=3, sticky="w", padx=5)
        tk.Checkbutton(self, text="Fill paths with element", variable=self.do_fill).grid(row=3, column=0, columnspan=3, sticky="w", padx=5)

        tk.Label(self, text="Spacing:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.spacing = tk.Entry(self, width=10)
        self.spacing.insert(0, "10.0")
        self.spacing.grid(row=4, column=1, sticky="w", padx=5, pady=5)

        tk.Label(self, text="Angle:").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.angle = tk.Entry(self, width=10)
        self.angle.insert(0, "0.0")
        self.angle.grid(row=5, column=1, sticky="w", padx=5, pady=5)

        tk.Button(self, text="Run", command=self.run).grid(row=6, column=0, columnspan=3, pady=15)

    # GUI helpers
    def _choose_bitmap(self):
        fname = filedialog.askopenfilename(filetypes=[("Images", "*.bmp;*.jpg;*.png;*.gif")])
        if fname:
            self.bmp_path.delete(0, tk.END)
            self.bmp_path.insert(0, fname)

    def _choose_element(self):
        fname = filedialog.askopenfilename(filetypes=[("CorelDRAW", "*.cdr"), ("SVG", "*.svg")])
        if fname:
            self.elm_path.delete(0, tk.END)
            self.elm_path.insert(0, fname)

    # Core logic
    def run(self):
        if win32 is None:
            messagebox.showerror("Missing dependency", "pywin32 is required to communicate with CorelDRAW.")
            return

        try:
            app = win32.Dispatch("CorelDRAW.Application")
            app.Visible = True
            doc = app.ActiveDocument
        except Exception as exc:
            messagebox.showerror("CorelDRAW error", f"Unable to connect: {exc}")
            return

        layer = doc.ActiveLayer
        spacing = float(self.spacing.get())
        angle = float(self.angle.get())

        traced_shape = None

        if self.do_trace.get():
            bmp_file = self.bmp_path.get()
            if not bmp_file:
                messagebox.showerror("Input missing", "Please choose a bitmap to trace.")
                return
            bitmap = layer.Import(bmp_file)
            # Use default outline tracing. Parameters might need tweaking for your image.
            trace = bitmap.Bitmap.Trace(0)  # cdrTraceType=0 for outline
            traced_shape = trace.Group
            bitmap.Delete()
        else:
            if doc.Selection.Count == 0:
                messagebox.showinfo("Selection", "Select a curve to use as the path.")
                return
            traced_shape = doc.Selection(1)

        if self.do_fill.get():
            element_file = self.elm_path.get()
            if not element_file:
                messagebox.showerror("Input missing", "Please choose an element to duplicate.")
                return
            element = layer.Import(element_file)
            try:
                self._fill_curve_with_element(traced_shape, element, spacing, angle)
            finally:
                element.Delete()

        messagebox.showinfo("Done", "Operation completed.")

    def _fill_curve_with_element(self, path_shape, element_shape, spacing, angle):
        """Duplicate `element_shape` along `path_shape` at given spacing and rotation."""
        curve = path_shape.Curve

        # Width of the element is used to prevent overlaps. CorelDRAW exposes
        # `SizeWidth` in document units; we step by this width plus user
        # spacing so adjacent duplicates touch only if spacing == 0.
        element_width = element_shape.SizeWidth
        step = element_width + spacing

        dist = 0.0
        total_len = curve.Length
        while dist <= total_len:
            point = curve.PositionAt(dist)
            dup = element_shape.Duplicate()
            dup.SetPosition(point.x, point.y)

            # Align the duplicate to the path's tangent and apply the user's
            # additional rotation. `TangentAt` returns a vector; convert to
            # degrees with atan2. Some versions expose `SlopeAt` instead; fall
            # back gracefully if tangent is unavailable.
            try:
                tangent = curve.TangentAt(dist)
                base_angle = math.degrees(math.atan2(tangent.y, tangent.x))
            except Exception:
                base_angle = 0.0
            dup.Rotation = base_angle + angle

            dist += step

if __name__ == "__main__":
    app = CorelVectorFillTool()
    app.mainloop()
