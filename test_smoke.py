#!/usr/bin/env python3
"""Smoke tests for Batch Image Resizer v11.

Tests the core logic (number conversion, shapes, colors, resize, trim)
without needing a display. Run:  python test_smoke.py
"""
import sys, types, re

# ───────────────────────────── mock tkinter ─────────────────────────────
tk = types.ModuleType("tkinter")
tk.Tk = lambda: (_ for _ in ()).throw(Exception("no display"))
for cls in ["Frame", "Canvas", "Label", "Button", "Entry", "Listbox",
            "Toplevel", "Radiobutton", "Checkbutton", "Spinbox", "Scale"]:
    setattr(tk, cls, type(cls, (), {}))
tk.StringVar = lambda v="": type("V", (), {"get": lambda s: s._v, "set": lambda s, x: setattr(s, "_v", x)})()
tk.ttk = types.ModuleType("tkinter.ttk")
for cls in ["Notebook", "Style", "Progressbar", "Scrollbar"]:
    setattr(tk.ttk, cls, type(cls, (), {}))
tk.filedialog = types.ModuleType("tkinter.filedialog")
tk.filedialog.askopenfilenames = lambda **k: ()
tk.filedialog.askdirectory = lambda **k: ""
tk.messagebox = types.ModuleType("tkinter.messagebox")
tk.messagebox.showinfo = lambda *a: None
tk.messagebox.showwarning = lambda *a: None
tk.messagebox.showerror = lambda *a: None

sys.modules["tkinter"] = tk
sys.modules["tkinter.ttk"] = tk.ttk
sys.modules["tkinter.filedialog"] = tk.filedialog
sys.modules["tkinter.messagebox"] = tk.messagebox

# ───────────────────────────── import module ─────────────────────────────
src = open("batch_image_resizer_v11_en.py", encoding="utf-8").read()
src = src.replace('if __name__=="__main__":', 'if False:')
exec(compile(src, "bir_en", "exec"), globals())

from PIL import Image, ImageDraw


# ───────────────────────────── tests ─────────────────────────────
def test_numbers_to_digits():
    assert _numbers_to_digits("یک") == "1"
    assert _numbers_to_digits("دوازده") == "12"
    assert _numbers_to_digits("سی و چهار") == "34"
    assert _numbers_to_digits("صد و بیست و سه") == "123"
    assert _numbers_to_digits("هزار و پانصد") == "1500"


def test_shape_points():
    assert len(_star_points(0, 0, 80, 60)) == 10
    assert len(_heart_points(0, 0, 80, 60)) == 100
    assert len(_hexagon_points(0, 0, 80, 60)) == 6
    assert len(_triangle_points(0, 0, 80, 60)) == 3


def test_colors():
    assert _lighten("#10b981", 20) == "#24cd95"
    assert _darken("#10b981", 20) == "#00a56d"
    m = _mix("#0891b2", "#7c3aed", 0.5)
    assert m.startswith("#") and len(m) == 7


def test_resize():
    img = Image.new("RGB", (100, 50), "red")
    r1 = apply_resize(img, "fit", 50, 50)
    assert r1.size == (50, 25)  # ratio preserved
    r2 = apply_resize(img, "exact", 50, 50)
    assert r2.size == (50, 50)
    r3 = apply_resize(img, "fill", 50, 50)
    assert r3.size == (50, 50)


def test_trim_background():
    img = Image.new("RGB", (300, 200), (255, 255, 255))
    d = ImageDraw.Draw(img)
    d.rectangle([100, 50, 200, 150], fill=(255, 0, 0))
    trimmed = trim_background(img, (255, 255, 255), 10, 0.05)
    assert trimmed.size[0] < 300  # cropped


def test_crop_shape_returns_pair():
    img = Image.new("RGB", (100, 100), "white")
    result = crop_shape(img, "ellipse", 1.0, 1.0, 0, 0)
    assert isinstance(result, tuple) and len(result) == 2


def test_compress():
    img = Image.new("RGB", (500, 500), (30, 60, 90))
    import os, tempfile
    tmp = tempfile.mktemp(suffix=".jpg")
    compress_to_kb(img, tmp, "JPEG", 100)
    assert os.path.getsize(tmp) // 1024 <= 120
    os.unlink(tmp)


def main():
    tests = [
        test_numbers_to_digits,
        test_shape_points,
        test_colors,
        test_resize,
        test_trim_background,
        test_crop_shape_returns_pair,
        test_compress,
    ]
    passed = 0
    for t in tests:
        t()
        passed += 1
        print(f"  ✓ {t.__name__}")
    print(f"\n✅ All {passed} tests passed!")


if __name__ == "__main__":
    main()
