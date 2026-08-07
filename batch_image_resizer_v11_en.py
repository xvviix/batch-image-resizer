import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os, io, time, math
from pathlib import Path
from PIL import Image, ImageDraw, ImageTk
from concurrent.futures import ThreadPoolExecutor, as_completed
import queue, numpy as np, cv2

# ── Optional libraries for the voice tab ────────────────────────────────
try:
    import speech_recognition as sr
    SR_AVAILABLE = True
except ImportError:
    SR_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import docx
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from pptx import Presentation
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

try:
    import vosk
    import json as _json
    VOSK_AVAILABLE = True
except ImportError:
    VOSK_AVAILABLE = False

# Available languages for speech recognition (Google Web Speech API codes)
VOICE_LANGS = [
    ("Persian (Iran)",      "fa-IR"),
    ("English (US)",        "en-US"),
    ("English (UK)",        "en-GB"),
    ("Arabic",              "ar-SA"),
    ("Turkish",              "tr-TR"),
    ("French",               "fr-FR"),
    ("Deutsch",             "de-DE"),
    ("Spanish",              "es-ES"),
    ("Russian",              "ru-RU"),
]

# ── Emerald color palette ──────────────────────────────────────────────────────────────
DARK_BG  = "#061a10"
PANEL_BG = "#0a2318"
CARD_BG  = "#0f2e1f"
ACCENT   = "#10b981"   # emerald-500
ACCENT_H = "#059669"   # emerald-600
ACCENT_L = "#34d399"   # emerald-400
SUCCESS  = "#6ee7b7"   # emerald-300
WARNING  = "#fbbf24"
ERROR    = "#f87171"
TEXT_P   = "#ecfdf5"   # emerald-50
TEXT_S   = "#86efac"   # green-300
BORDER   = "#14532d"   # green-900
BORDER_L = "#166534"   # green-800

# ── Font setup: Vazir with fallback ──────────────────────────────────────────
import tkinter as _tk_check
def _best_persian_font():
    """Returns the best Persian-compatible font available on the system."""
    candidates = ["Vazir", "Vazirmatn", "Sahel", "B Nazanin", "IranSans",
                  "Tahoma", "Arial Unicode MS", "Segoe UI"]
    try:
        root_tmp = _tk_check.Tk(); root_tmp.withdraw()
        available = set(_tk_check.font.families(root_tmp))
        root_tmp.destroy()
        for f in candidates:
            if f in available:
                return f
    except Exception:
        pass
    return "Tahoma"

_PFONT = _best_persian_font()

FT  = (_PFONT, 18, "bold")
FH  = (_PFONT, 12, "bold")
FB  = (_PFONT, 10)
FS  = (_PFONT, 9)
FM  = ("Consolas", 9)

FORMATS = ["JPEG", "PNG", "WEBP", "BMP", "TIFF"]
FMT_EXT = {"JPEG":".jpg","PNG":".png","WEBP":".webp","BMP":".bmp","TIFF":".tif"}

SHAPES = ["ellipse", "rectangle", "star", "heart", "hexagon", "triangle"]
SHAPE_LABELS = {
    "ellipse": "Ellipse",
    "rectangle": "Rectangle",
    "star": "Star",
    "heart": "Heart",
    "hexagon": "Hexagon",
    "triangle": "Triangle",
}

# ─────────────────────────────────────────────────────────────────────────────
#  Spoken-number → digit conversion (Persian + English)
# ─────────────────────────────────────────────────────────────────────────────
_FA_UNITS = {
    "صفر":0, "یک":1, "دو":2, "سه":3, "چهار":4, "پنج":5, "شش":6, "شیش":6,
    "هفت":7, "هشت":8, "نه":9, "ده":10,
    "یازده":11, "دوازده":12, "سیزده":13, "چهارده":14, "پانزده":15,
    "شانزده":16, "هفده":17, "هجده":18, "نوزده":19,
}
_FA_TENS = {
    "بیست":20, "سی":30, "چهل":40, "پنجاه":50, "شصت":60,
    "هفتاد":70, "هشتاد":80, "نود":90,
}
_FA_HUNDREDS = {
    "صد":100, "یکصد":100, "دویست":200, "سیصد":300, "چهارصد":400,
    "پانصد":500, "ششصد":600, "هفتصد":700, "هشتصد":800, "نهصد":900,
}
_FA_SCALES = {"هزار":1000, "میلیون":1000000, "میلیارد":1000000000}
_FA_AND = "و"

_FA_NUMBER_WORDS = set(_FA_UNITS) | set(_FA_TENS) | set(_FA_HUNDREDS) | set(_FA_SCALES) | {_FA_AND}

_EN_UNITS = {
    "zero":0,"one":1,"two":2,"three":3,"four":4,"five":5,"six":6,"seven":7,
    "eight":8,"nine":9,"ten":10,"eleven":11,"twelve":12,"thirteen":13,
    "fourteen":14,"fifteen":15,"sixteen":16,"seventeen":17,"eighteen":18,"nineteen":19,
}
_EN_TENS = {"twenty":20,"thirty":30,"forty":40,"fifty":50,"sixty":60,"seventy":70,"eighty":80,"ninety":90}
_EN_HUNDRED = {"hundred":100}
_EN_SCALES = {"thousand":1000,"million":1000000,"billion":1000000000}
_EN_NUMBER_WORDS = set(_EN_UNITS) | set(_EN_TENS) | set(_EN_HUNDRED) | set(_EN_SCALES) | {"and"}


def _convert_number_group(words, units, tens, hundreds, scales, and_word):
    """Converts a list of spoken-number tokens (without scale words except within them) to an integer."""
    total = 0
    current = 0
    i = 0
    while i < len(words):
        w = words[i].lower()
        if w == and_word:
            i += 1; continue
        if w in scales:
            mult = scales[w]
            if current == 0:
                current = 1
            total += current * mult
            current = 0
        elif w in hundreds:
            current += hundreds[w]
        elif w in tens:
            current += tens[w]
        elif w in units:
            current += units[w]
        else:
            # Unknown token, ignored
            pass
        i += 1
    total += current
    return total


def _numbers_to_digits(text):
    """Replaces runs of spoken-number words (Persian or English) in `text`
    with their digit representation. Everything else is left unchanged."""
    if not text:
        return text

    tokens = text.split(" ")
    out = []
    i = 0
    n = len(tokens)
    while i < n:
        tok = tokens[i]
        bare = tok.strip("،.!؟,.!؟")
        lower = bare.lower()

        if bare in _FA_NUMBER_WORDS or lower in _EN_NUMBER_WORDS:
            is_fa = bare in _FA_NUMBER_WORDS
            group = []
            j = i
            while j < n:
                b = tokens[j].strip("،.!؟,.!؟")
                l = b.lower()
                if is_fa and b in _FA_NUMBER_WORDS:
                    group.append(b); j += 1
                elif (not is_fa) and l in _EN_NUMBER_WORDS:
                    group.append(l); j += 1
                else:
                    break
            # Drop a lone trailing "and" (likely belongs to the next phrase)
            while group and group[-1] in (_FA_AND, "and"):
                group.pop(); j -= 1
            if group:
                if is_fa:
                    value = _convert_number_group(group, _FA_UNITS, _FA_TENS,
                                                    _FA_HUNDREDS, _FA_SCALES, _FA_AND)
                else:
                    value = _convert_number_group(group, _EN_UNITS, _EN_TENS,
                                                    {}, {**_EN_HUNDRED, **_EN_SCALES}, "and")
                # Preserve any trailing punctuation attached to the last consumed token
                last_tok = tokens[j-1]
                trail = last_tok[len(last_tok.rstrip("،.!؟,.!؟")):]
                out.append(str(value) + trail)
                i = j
                continue

        out.append(tok)
        i += 1

    return " ".join(out)




def _star_points(cx, cy, ew, eh, n=5, inner_ratio=0.5, rotation=-90):
    """Generates points for a star polygon inside an ew x eh box."""
    pts = []
    rx, ry = ew/2, eh/2
    for i in range(n*2):
        ang = math.radians(rotation) + i*math.pi/n
        r = 1.0 if i % 2 == 0 else inner_ratio
        pts.append((cx + r*rx*math.cos(ang), cy + r*ry*math.sin(ang)))
    return pts

def _heart_points(cx, cy, ew, eh, n=100):
    """Generates approximate points for a heart shape inside an ew x eh box."""
    pts = []
    rx, ry = ew/2, eh/2
    for i in range(n):
        t = (i/n) * 2*math.pi
        x = 16*(math.sin(t)**3)
        y = -(13*math.cos(t) - 5*math.cos(2*t) - 2*math.cos(3*t) - math.cos(4*t))
        # Normalize to [-1,1]
        x_n = x/17.0
        y_n = y/17.0
        pts.append((cx + x_n*rx, cy + y_n*ry))
    return pts

def _hexagon_points(cx, cy, ew, eh, rotation=-90):
    pts = []
    rx, ry = ew/2, eh/2
    for i in range(6):
        ang = math.radians(rotation) + i*math.pi/3
        pts.append((cx + rx*math.cos(ang), cy + ry*math.sin(ang)))
    return pts

def _triangle_points(cx, cy, ew, eh, rotation=-90):
    pts = []
    rx, ry = ew/2, eh/2
    for i in range(3):
        ang = math.radians(rotation) + i*2*math.pi/3
        pts.append((cx + rx*math.cos(ang), cy + ry*math.sin(ang)))
    return pts


def draw_shape_mask(draw, shape, cx, cy, ew, eh, fill=255, **kw):
    """Draws the given shape onto a PIL ImageDraw mask, centered at (cx,cy)."""
    ew = max(2, ew); eh = max(2, eh)
    if shape == "ellipse":
        draw.ellipse((cx-ew/2, cy-eh/2, cx+ew/2, cy+eh/2), fill=fill)
    elif shape == "rectangle":
        draw.rectangle((cx-ew/2, cy-eh/2, cx+ew/2, cy+eh/2), fill=fill)
    elif shape == "star":
        pts = _star_points(cx, cy, ew, eh,
                           n=kw.get("points", 5),
                           inner_ratio=kw.get("inner_ratio", 0.5))
        draw.polygon(pts, fill=fill)
    elif shape == "heart":
        pts = _heart_points(cx, cy, ew, eh)
        draw.polygon(pts, fill=fill)
    elif shape == "hexagon":
        pts = _hexagon_points(cx, cy, ew, eh)
        draw.polygon(pts, fill=fill)
    elif shape == "triangle":
        pts = _triangle_points(cx, cy, ew, eh)
        draw.polygon(pts, fill=fill)
    else:
        draw.ellipse((cx-ew/2, cy-eh/2, cx+ew/2, cy+eh/2), fill=fill)


# ─────────────────────────────────────────────────────────────────────────────
#  Animation utilities and dynamic buttons
# ─────────────────────────────────────────────────────────────────────────────
def _lighten(hex_c, n=18):
    """Lightens a hex color by n."""
    try:
        r,g,b = int(hex_c[1:3],16), int(hex_c[3:5],16), int(hex_c[5:7],16)
        return f'#{min(255,r+n):02x}{min(255,g+n):02x}{min(255,b+n):02x}'
    except Exception:
        return hex_c

def _darken(hex_c, n=18):
    """Darkens a hex color by n."""
    try:
        r,g,b = int(hex_c[1:3],16), int(hex_c[3:5],16), int(hex_c[5:7],16)
        return f'#{max(0,r-n):02x}{max(0,g-n):02x}{max(0,b-n):02x}'
    except Exception:
        return hex_c

def _mix(c1, c2, t=0.5):
    """Blends two hex colors with ratio t."""
    try:
        r1,g1,b1 = int(c1[1:3],16),int(c1[3:5],16),int(c1[5:7],16)
        r2,g2,b2 = int(c2[1:3],16),int(c2[3:5],16),int(c2[5:7],16)
        r = int(r1*(1-t)+r2*t)
        g = int(g1*(1-t)+g2*t)
        b = int(b1*(1-t)+b2*t)
        return f'#{r:02x}{g:02x}{b:02x}'
    except Exception:
        return c1


class AnimatedButton(tk.Button):
    """An advanced button with hover/press animation, fixing the click-feedback issue."""

    def __init__(self, parent, hover_bg=None, press_bg=None, **kwargs):
        self._normal_bg  = kwargs.get('bg', CARD_BG)
        self._normal_fg  = kwargs.get('fg', TEXT_P)
        self._hover_bg   = hover_bg  or _lighten(self._normal_bg, 22)
        self._press_bg   = press_bg  or _darken(self._normal_bg, 10)
        self._disabled   = False

        # Make sure default values are set correctly
        kwargs.setdefault('relief',         'flat')
        kwargs.setdefault('bd',             0)
        kwargs.setdefault('cursor',         'hand2')
        kwargs.setdefault('activeforeground', kwargs.get('fg', TEXT_P))
        kwargs['activebackground'] = self._hover_bg

        super().__init__(parent, **kwargs)

        self.bind('<Enter>',          self._on_enter)
        self.bind('<Leave>',          self._on_leave)
        self.bind('<Button-1>',       self._on_press)
        self.bind('<ButtonRelease-1>',self._on_release)

    def _on_enter(self, _):
        if str(self.cget('state')) != 'disabled':
            self.config(bg=self._hover_bg)

    def _on_leave(self, _):
        if str(self.cget('state')) != 'disabled':
            self.config(bg=self._normal_bg)

    def _on_press(self, _):
        if str(self.cget('state')) != 'disabled':
            self.config(bg=self._press_bg)

    def _on_release(self, _):
        if str(self.cget('state')) != 'disabled':
            self.config(bg=self._hover_bg)


def _attach_hover(btn, normal_bg, hover_bg=None, press_bg=None):
    """Adds hover/press animation to an existing tk.Button."""
    h = hover_bg or _lighten(normal_bg, 22)
    p = press_bg or _darken(normal_bg, 10)
    btn.bind('<Enter>',           lambda e: btn.winfo_exists() and
             str(btn.cget('state')) != 'disabled' and btn.config(bg=h))
    btn.bind('<Leave>',           lambda e: btn.winfo_exists() and
             str(btn.cget('state')) != 'disabled' and btn.config(bg=normal_bg))
    btn.bind('<Button-1>',        lambda e: btn.winfo_exists() and
             str(btn.cget('state')) != 'disabled' and btn.config(bg=p))
    btn.bind('<ButtonRelease-1>', lambda e: btn.winfo_exists() and
             str(btn.cget('state')) != 'disabled' and btn.config(bg=h))


def _glow_frame(entry_widget, host_frame):
    """Adds a glow animation to entry_widget on focus."""
    def _in(_):
        try: host_frame.config(bg=ACCENT)
        except Exception: pass
    def _out(_):
        try: host_frame.config(bg=BORDER)
        except Exception: pass
    entry_widget.bind('<FocusIn>',  _in)
    entry_widget.bind('<FocusOut>', _out)


def _styled_entry(parent, textvariable=None, width=None, font=None, justify="left",
                  ipady=6, fill="x", expand=True, side="left", padx=0, pady=0):
    """Creates an entry with a glowing emerald border and returns (widget, frame)."""
    frm = tk.Frame(parent, bg=BORDER, bd=0, padx=1, pady=1)
    kw = dict(bg=PANEL_BG, fg=TEXT_P, insertbackground=ACCENT,
              relief='flat', bd=0, font=font or FM)
    if textvariable: kw['textvariable'] = textvariable
    if width:        kw['width'] = width
    if justify:      kw['justify'] = justify
    ent = tk.Entry(frm, **kw)
    ent.pack(fill='both', expand=True, ipady=ipady)
    _glow_frame(ent, frm)
    frm.pack(side=side, fill=fill, expand=expand, padx=padx, pady=pady)
    return ent, frm


def shape_outline_points(shape, cx, cy, ew, eh, **kw):
    """Returns the list of (x,y) points describing the shape outline,
    for drawing on a tk.Canvas (used by the overlay preview)."""
    ew = max(2, ew); eh = max(2, eh)
    if shape == "ellipse":
        return None  # Handled separately (ellipse)
    if shape == "rectangle":
        return None  # Handled separately (rectangle)
    if shape == "star":
        return _star_points(cx, cy, ew, eh,
                            n=kw.get("points", 5),
                            inner_ratio=kw.get("inner_ratio", 0.5))
    if shape == "heart":
        return _heart_points(cx, cy, ew, eh)
    if shape == "hexagon":
        return _hexagon_points(cx, cy, ew, eh)
    if shape == "triangle":
        return _triangle_points(cx, cy, ew, eh)
    return None


# ─────────────────────────────────────────────────────────────────────────────
#  Core: Shape Crop
# ─────────────────────────────────────────────────────────────────────────────
def crop_shape(img_pil, shape="ellipse", rx=1.0, ry=1.0, ox=0.0, oy=0.0,
              bg=(255,255,255), **kw):
    """
    Crops the given shape from the image.
    rx / ry : fraction of width/height the shape's bounding box covers (0-1)
    ox /oy : center offset as a fraction of width/height (-0.5..0.5)
    bg      : RGB background color outside the shape
    """
    img = img_pil.convert("RGBA")
    w, h = img.size
    ew = max(2, w * min(max(rx,0.01),1.0))
    eh = max(2, h * min(max(ry,0.01),1.0))
    cx = w*(0.5+ox)
    cy = h*(0.5+oy)
    # Clamp the center so the bounding box stays within the image
    cx = max(ew/2, min(w-ew/2, cx))
    cy = max(eh/2, min(h-eh/2, cy))

    mask = Image.new("L", (w, h), 0)
    draw_shape_mask(ImageDraw.Draw(mask), shape, cx, cy, ew, eh, fill=255, **kw)

    mn = np.array(mask, np.uint8)
    mn = cv2.GaussianBlur(mn, (5,5), 0)
    ms = Image.fromarray(mn, "L")

    bg_img = Image.new("RGBA", (w,h), (bg or (255,255,255))+(255,))
    return Image.composite(img, bg_img, ms).convert("RGB"), mask


# ─────────────────────────────────────────────────────────────────────────────
#  Core: Remove surrounding background (flexible bounding-box crop)
# ─────────────────────────────────────────────────────────────────────────────
def trim_background(img_pil, bg=(255,255,255), tolerance=10, padding_pct=0.0):
    """
    Finds the bounding box of non-background pixels and crops to it.
    tolerance : allowed color distance to count as background (0-255)
    padding_pct : extra margin around the bounding box, as a percent of its size
    Returns the cropped image. If everything is background, returns the original.
    """
    img = img_pil.convert("RGB")
    arr = np.array(img)
    bg_arr = np.array(bg if bg is not None else (255,255,255), dtype=np.int16)

    diff = np.abs(arr.astype(np.int16) - bg_arr).sum(axis=2)
    mask = diff > tolerance * 3  # combined 3-channel threshold

    if not mask.any():
        return img  # everything is background, nothing to remove

    ys, xs = np.where(mask)
    y0, y1 = ys.min(), ys.max()+1
    x0, x1 = xs.min(), xs.max()+1

    if padding_pct > 0:
        bw, bh = x1-x0, y1-y0
        px = int(bw * padding_pct/100)
        py = int(bh * padding_pct/100)
        x0 = max(0, x0-px); y0 = max(0, y0-py)
        x1 = min(img.width, x1+px); y1 = min(img.height, y1+py)

    return img.crop((x0, y0, x1, y1))


def trim_background_custom(img_pil, rect):
    """Crops the image to an explicit rectangle (x0,y0,x1,y1) (manual crop)."""
    x0,y0,x1,y1 = rect
    x0 = max(0, int(x0)); y0 = max(0, int(y0))
    x1 = min(img_pil.width, int(x1)); y1 = min(img_pil.height, int(y1))
    if x1 <= x0 or y1 <= y0:
        return img_pil
    return img_pil.crop((x0,y0,x1,y1))


# ─────────────────────────────────────────────────────────────────────────────
#  Core: Compression
# ─────────────────────────────────────────────────────────────────────────────
def compress_to_kb(img, out_path, fmt, target_kb):
    tb = int(target_kb*1024)
    def _s(img,fmt,q):
        buf=io.BytesIO(); si=img if img.mode=="RGB" else img.convert("RGB")
        if fmt in("JPEG","JPG"): si.save(buf,format="JPEG",quality=q,subsampling=0,optimize=True)
        elif fmt=="WEBP": si.save(buf,format="WEBP",quality=q,method=6)
        else: si.save(buf,format=fmt)
        return buf.getvalue()
    if fmt in("JPEG","JPG","WEBP"):
        lo,hi,best=10,97,None
        for _ in range(13):
            mid=(lo+hi)//2; d=_s(img,fmt,mid)
            if len(d)<=tb: best=d; lo=mid+1
            else: hi=mid-1
            if lo>hi: break
        with open(out_path,"wb") as f: f.write(best or _s(img,fmt,10))
    else:
        for lvl in range(1,10):
            buf=io.BytesIO(); img.save(buf,format=fmt,compress_level=lvl)
            if buf.tell()<=tb:
                with open(out_path,"wb") as f: f.write(buf.getvalue()); return
        with open(out_path,"wb") as f: f.write(_s(img,"JPEG",85))

def prep_mode(img,fmt):
    if fmt in("JPEG","JPG","BMP") and img.mode not in("RGB","L"):
        return img.convert("RGB")
    return img

def save_kw(fmt,info=None):
    info=info or {}
    if fmt in("JPEG","JPG"):
        kw={"quality":97,"subsampling":0,"optimize":True}
        if "dpi" in info: kw["dpi"]=info["dpi"]
        return kw
    if fmt=="PNG": return{"compress_level":1}
    if fmt=="WEBP": return{"quality":97,"method":6}
    return {}


# ─────────────────────────────────────────────────────────────────────────────
#  Core: Resizing
# ─────────────────────────────────────────────────────────────────────────────
def apply_resize(img, mode, tw, th):
    if mode == "no_resize":
        return img
    if mode == "exact":
        return img.resize((tw,th), Image.LANCZOS)
    if mode == "fit":
        img2 = img.copy()
        img2.thumbnail((tw,th), Image.LANCZOS)
        return img2
    if mode == "fill":
        r=max(tw/img.width, th/img.height)
        nw,nh=int(img.width*r),int(img.height*r)
        img2=img.resize((nw,nh),Image.LANCZOS)
        l=(nw-tw)//2; t=(nh-th)//2
        return img2.crop((l,t,l+tw,t+th))
    if mode == "width_only":
        r=tw/img.width
        return img.resize((tw,int(img.height*r)),Image.LANCZOS)
    if mode == "height_only":
        r=th/img.height
        return img.resize((int(img.width*r),th),Image.LANCZOS)
    return img


# ─────────────────────────────────────────────────────────────────────────────
#  Core: Processing a single image (batch pipeline)
#  Pipeline: shape crop ← trim edges ← resize ← save
# ─────────────────────────────────────────────────────────────────────────────
def process_image(args):
    (fp, out_dir, tw, th, mode, keep, base_dir,
     do_shape, shape, rx, ry, ox, oy, bg, shape_kw,
     do_trim, trim_tol, trim_pad, trim_custom,
     lim_size, target_kb, out_fmt, cmap) = args
    try:
        with Image.open(fp) as _i:
            orig_fmt  = (_i.format or "PNG").upper()
            orig_info = _i.info.copy()
            img = _i.copy()

        # Per-image override (manual session edits)
        if cmap and fp in cmap:
            ov = cmap[fp]
            _shape = ov.get("shape", shape)
            _rx,_ry = ov.get("rx", rx), ov.get("ry", ry)
            _ox,_oy = ov.get("ox", ox), ov.get("oy", oy)
            _shape_kw = ov.get("shape_kw", shape_kw)
            _do_shape = True
            _trim_custom = ov.get("trim_rect", None)
            _do_trim = ov.get("do_trim", do_trim)
        else:
            _shape,_rx,_ry,_ox,_oy,_shape_kw = shape,rx,ry,ox,oy,shape_kw
            _do_shape = do_shape
            _trim_custom = trim_custom
            _do_trim = do_trim

        if _do_shape:
            img, _ = crop_shape(img, _shape, _rx/100, _ry/100, _ox, _oy, bg, **_shape_kw)

        # Remove background / custom rectangle
        if _trim_custom:
            img = trim_background_custom(img, _trim_custom)
        elif _do_trim:
            img = trim_background(img, bg, trim_tol, trim_pad)

        img = apply_resize(img, mode, tw, th)

        sfmt = (out_fmt or orig_fmt).upper()
        if sfmt=="JPG": sfmt="JPEG"
        img  = prep_mode(img, sfmt)
        ext  = FMT_EXT.get(sfmt, Path(fp).suffix)
        fname= Path(fp).stem + ext

        if keep and base_dir:
            rel = Path(fp).relative_to(base_dir)
            op  = Path(out_dir)/rel.with_suffix(ext)
        else:
            op  = Path(out_dir)/fname
        op.parent.mkdir(parents=True, exist_ok=True)

        if lim_size and target_kb:
            compress_to_kb(img, str(op), sfmt, target_kb)
        else:
            img.save(str(op), format=sfmt, **save_kw(sfmt, orig_info))

        return (True, fp, str(op), os.path.getsize(str(op))/1024, "")
    except Exception:
        import traceback; return (False, fp, None, 0, traceback.format_exc())

# ─────────────────────────────────────────────────────────────────────────────
#  Interactive Session Editor  (standalone window)
#
#  Two modes, selectable per image session:
#   - "crop"  : choose a shape, draw/resize it on the image, live preview
#   - "trim"  : draw a custom rectangle to remove empty space around it
#
#  Left side  = original image with overlay (shape outline or crop rectangle)
#  Right side = live preview of the result
#  Navigation: previous / next in folder
#  Save All & Close → writes everything to the output folder
# ─────────────────────────────────────────────────────────────────────────────
class SessionEditor(tk.Toplevel):
    PW = 520
    PH = 480

    def __init__(self, parent, files, out_dir,
                 mode="crop",                     # "crop" or "trim"
                 init_shape="ellipse",
                 init_rx=90, init_ry=90,
                 bg_color=(255,255,255),
                 out_fmt=None, lim_size=False, target_kb=None,
                 resize_mode="no_resize", tw=413, th=531,
                 trim_tol=10, trim_pad=0):
        super().__init__(parent)
        self.title("Session Editor")
        self.configure(bg=DARK_BG)
        self.resizable(True, True)
        self.grab_set()

        self.files      = list(files)
        self.out_dir    = out_dir
        self.mode       = mode   # "crop" | "trim"
        self.bg_color   = bg_color
        self.out_fmt    = out_fmt
        self.lim_size   = lim_size
        self.target_kb  = target_kb
        self.resize_mode= resize_mode
        self.tw, self.th= tw, th
        self.trim_tol   = trim_tol
        self.trim_pad   = trim_pad

        # Per-image state for crop mode:
        #   {path: {"shape":str,"ox":f,"oy":f,"rx":pct,"ry":pct}}
        self.crop_states = {
            f: {"shape": init_shape, "ox":0.0, "oy":0.0,
                "rx": init_rx, "ry": init_ry}
            for f in files
        }
        # Per-image state for trim mode:
        #   {path: (x0,y0,x1,y1) in original image pixel coords, or None=auto}
        self.trim_states = {f: None for f in files}
        # Whether the auto-trim suggestion has been computed for this image
        self._auto_rect_cache = {}

        self.idx = 0
        self._dragging = False
        self._drag_mode = None   # "move" | "resize" | "new_rect" | "move_rect" | "resize_rect"
        self._resize_handle = None

        self._build()
        self._load_current()

    # ── Build UI ─────────────────────────────────────────────────────────
    def _build(self):
        top = tk.Frame(self, bg=PANEL_BG, height=44)
        top.pack(fill="x"); top.pack_propagate(False)
        self._nav_lbl = tk.Label(top, text="", font=FH, bg=PANEL_BG, fg=TEXT_P)
        self._nav_lbl.pack(side="left", padx=16)
        self._fname_lbl = tk.Label(top, text="", font=FM, bg=PANEL_BG, fg=TEXT_S)
        self._fname_lbl.pack(side="left", padx=4)
        mode_txt = "✂ Shape Crop" if self.mode=="crop" else "▭ Trim Empty Edges"
        tk.Label(top, text=mode_txt, font=FH, bg=PANEL_BG, fg=ACCENT).pack(side="right", padx=16)

        main = tk.Frame(self, bg=DARK_BG)
        main.pack(fill="both", expand=True)

        lf = tk.Frame(main, bg=DARK_BG)
        lf.pack(side="left", fill="both", expand=True, padx=(12,6), pady=12)
        left_hint = ("Drag to move, drag the corners to resize"
                      if self.mode=="crop" else
                      "Drag to draw/move the crop rectangle, drag the corners to resize")
        tk.Label(lf, text=f"Original Image  —  {left_hint}",
                 font=FS, bg=DARK_BG, fg=TEXT_S).pack(anchor="w")
        self.left_canvas = tk.Canvas(lf, width=self.PW, height=self.PH,
                                      bg="#111", cursor="crosshair",
                                      highlightthickness=1, highlightbackground=BORDER)
        self.left_canvas.pack()
        self.left_canvas.bind("<Button-1>",       self._on_press)
        self.left_canvas.bind("<B1-Motion>",       self._on_drag)
        self.left_canvas.bind("<ButtonRelease-1>", self._on_release)

        rf = tk.Frame(main, bg=DARK_BG)
        rf.pack(side="left", fill="both", expand=True, padx=(6,12), pady=12)
        tk.Label(rf, text="Preview  —  result of this step",
                 font=FS, bg=DARK_BG, fg=TEXT_S).pack(anchor="w")
        self.right_canvas = tk.Canvas(rf, width=self.PW, height=self.PH,
                                       bg="#111", highlightthickness=1, highlightbackground=BORDER)
        self.right_canvas.pack()

        ctrl = tk.Frame(self, bg=PANEL_BG)
        ctrl.pack(fill="x")

        if self.mode == "crop":
            self._build_crop_controls(ctrl)
        else:
            self._build_trim_controls(ctrl)

        # Shared bottom buttons
        btn_row = tk.Frame(ctrl, bg=PANEL_BG)
        btn_row.pack(pady=(4,12), padx=16, fill="x")

        AnimatedButton(btn_row, text="↺  Reset", font=FB,
                  bg=CARD_BG, fg=TEXT_S, padx=14, pady=9,
                  hover_bg=_lighten(CARD_BG,20), press_bg=_darken(CARD_BG,10),
                  command=self._reset_current).pack(side="left", padx=(0,8))
        AnimatedButton(btn_row, text="◀  Previous", font=FB,
                  bg=CARD_BG, fg=TEXT_S, padx=14, pady=9,
                  hover_bg=_lighten(CARD_BG,20), press_bg=_darken(CARD_BG,10),
                  command=self._prev).pack(side="left", padx=(0,8))
        AnimatedButton(btn_row, text="Next  ▶", font=FB,
                  bg=CARD_BG, fg=TEXT_S, padx=14, pady=9,
                  hover_bg=_lighten(CARD_BG,20), press_bg=_darken(CARD_BG,10),
                  command=self._next).pack(side="left", padx=(0,8))

        tk.Frame(btn_row, bg=PANEL_BG).pack(side="left", fill="x", expand=True)

        AnimatedButton(btn_row, text="💾  Save All & Close", font=FH,
                  bg=SUCCESS, fg=DARK_BG,
                  hover_bg=_lighten(SUCCESS,12), press_bg=_darken(SUCCESS,20),
                  padx=22, pady=9,
                  command=self._save_all).pack(side="right")
        AnimatedButton(btn_row, text="✕  Cancel", font=FB,
                  bg=CARD_BG, fg=ERROR,
                  hover_bg=_mix(CARD_BG, ERROR, 0.15), press_bg=_mix(CARD_BG, ERROR, 0.25),
                  padx=14, pady=9,
                  command=self.destroy).pack(side="right", padx=(0,8))

        self._prog_var = tk.DoubleVar()
        self._prog_lbl = tk.Label(ctrl, text="", font=FS, bg=PANEL_BG, fg=TEXT_S)
        self._prog_lbl.pack()
        self._prog_bar = ttk.Progressbar(ctrl, variable=self._prog_var, maximum=100)
        self._prog_bar.pack(fill="x", padx=16, pady=(0,10))

    # ── Crop-mode controls ───────────────────────────────────────────────────
    def _build_crop_controls(self, ctrl):
        top = tk.Frame(ctrl, bg=PANEL_BG); top.pack(fill="x", padx=16, pady=(10,2))
        tk.Label(top, text="Shape:", font=FS, bg=PANEL_BG, fg=TEXT_S).pack(side="left", padx=(0,8))
        self._shape_var = tk.StringVar(value="ellipse")
        for sh in SHAPES:
            tk.Radiobutton(top, text=SHAPE_LABELS[sh], variable=self._shape_var, value=sh,
                           bg=PANEL_BG, fg=TEXT_S, selectcolor=DARK_BG,
                           activebackground=PANEL_BG, activeforeground=ACCENT_L,
                           font=FS, bd=0, cursor="hand2",
                           command=self._on_shape_change).pack(side="left", padx=(0,8))

        sliders = tk.Frame(ctrl, bg=PANEL_BG)
        sliders.pack(fill="x", padx=16, pady=(6,4))

        def _sl(parent, label, var, lo, hi, suffix="%"):
            row = tk.Frame(parent, bg=PANEL_BG); row.pack(fill="x", pady=2)
            tk.Label(row, text=label, font=FS, bg=PANEL_BG,
                     fg=TEXT_S, width=18, anchor="w").pack(side="left")
            lbl = tk.Label(row, text=f"{var.get()}{suffix}",
                           font=(_PFONT,10,"bold"), bg=PANEL_BG, fg=ACCENT_L, width=6)
            lbl.pack(side="right")
            tk.Scale(row, from_=lo, to=hi, variable=var,
                     orient="horizontal", bg=PANEL_BG, fg=ACCENT,
                     troughcolor=BORDER_L, activebackground=ACCENT_L,
                     sliderrelief="flat",
                     highlightthickness=0, bd=0, showvalue=False,
                     command=lambda v,l=lbl,s=suffix:(
                         l.config(text=f"{int(float(v))}{s}"),
                         self._update_state_from_sliders(),
                         self._redraw()
                     )).pack(side="left", fill="x", expand=True, padx=(0,6))
            return row

        self._rx = tk.IntVar(value=90)
        self._ry = tk.IntVar(value=90)
        _sl(sliders, "Shape width  (% of image)", self._rx, 10, 100)
        _sl(sliders, "Shape height (% of image)", self._ry, 10, 100)

        self._coord_lbl = tk.Label(ctrl, text="", font=FS, bg=PANEL_BG, fg=TEXT_S)
        self._coord_lbl.pack(pady=(0,6))

    # ── Trim-mode controls ───────────────────────────────────────────────────
    def _build_trim_controls(self, ctrl):
        top = tk.Frame(ctrl, bg=PANEL_BG); top.pack(fill="x", padx=16, pady=(10,4))
        tk.Label(top,
                 text="Draw a rectangle around the area you want to keep. "
                      "Everything outside it will be removed.",
                 font=FS, bg=PANEL_BG, fg=TEXT_S, justify="left").pack(side="left")

        row = tk.Frame(ctrl, bg=PANEL_BG); row.pack(fill="x", padx=16, pady=(2,4))
        AnimatedButton(row, text="✨ Auto-Suggest Crop", font=FS,
                  bg=CARD_BG, fg=ACCENT_L,
                  hover_bg=_lighten(CARD_BG,20), press_bg=_darken(CARD_BG,10),
                  padx=10, pady=4,
                  command=self._auto_suggest_trim).pack(side="left", padx=(0,8))
        tk.Label(row, text="(automatically detects the background color and suggests a box)",
                 font=FS, bg=PANEL_BG, fg=TEXT_S).pack(side="left")

        self._coord_lbl = tk.Label(ctrl, text="", font=FS, bg=PANEL_BG, fg=TEXT_S)
        self._coord_lbl.pack(pady=(2,6))

    # ── Load image ─────────────────────────────────────────────────────────
    def _load_current(self):
        if not self.files: return
        fp = self.files[self.idx]
        self._orig_pil = Image.open(fp).convert("RGB")
        ow, oh = self._orig_pil.size

        scale = min(self.PW/ow, self.PH/oh, 1.0)
        self._dw = max(1, int(ow*scale))
        self._dh = max(1, int(oh*scale))
        self._scale = scale
        self._thumb = self._orig_pil.resize((self._dw,self._dh), Image.LANCZOS)

        if self.mode == "crop":
            st = self.crop_states[fp]
            self._shape_var.set(st["shape"])
            self._ox_frac = st["ox"]; self._oy_frac = st["oy"]
            self._rx.set(st["rx"]); self._ry.set(st["ry"])
        else:
            # Crop rectangle in display coordinates
            rect = self.trim_states[fp]
            if rect is None:
                self._trim_rect_disp = None
            else:
                x0,y0,x1,y1 = rect
                self._trim_rect_disp = (x0*self._scale, y0*self._scale,
                                        x1*self._scale, y1*self._scale)

        self._nav_lbl.config(text=f"{self.idx+1} / {len(self.files)}")
        self._fname_lbl.config(text=Path(fp).name)
        self._redraw()

    def _img_offset(self):
        return (self.PW-self._dw)//2, (self.PH-self._dh)//2

    # ── Draw ──────────────────────────────────────────────────────────────────
    def _redraw(self):
        self._draw_left()
        self._draw_right()

    def _draw_left(self):
        c = self.left_canvas; c.delete("all")
        ox_px, oy_px = self._img_offset()
        self._tk_left = ImageTk.PhotoImage(self._thumb)
        c.create_image(ox_px, oy_px, anchor="nw", image=self._tk_left)

        if self.mode == "crop":
            self._draw_left_crop(c, ox_px, oy_px)
        else:
            self._draw_left_trim(c, ox_px, oy_px)

    def _draw_left_crop(self, c, ox_px, oy_px):
        shape = self._shape_var.get()
        ew = self._dw * (self._rx.get()/100)
        eh = self._dh * (self._ry.get()/100)
        cx = ox_px + self._dw*(0.5+self._ox_frac)
        cy = oy_px + self._dh*(0.5+self._oy_frac)
        # Clamp like crop_shape
        cx = max(ox_px+ew/2, min(ox_px+self._dw-ew/2, cx))
        cy = max(oy_px+eh/2, min(oy_px+self._dh-eh/2, cy))

        if shape == "ellipse":
            c.create_oval(cx-ew/2, cy-eh/2, cx+ew/2, cy+eh/2,
                          outline=ACCENT, width=2, dash=(8,4))
        elif shape == "rectangle":
            c.create_rectangle(cx-ew/2, cy-eh/2, cx+ew/2, cy+eh/2,
                               outline=ACCENT, width=2, dash=(8,4))
        else:
            pts = shape_outline_points(shape, cx, cy, ew, eh)
            if pts:
                flat = []
                for px,py in pts: flat += [px,py]
                c.create_polygon(*flat, outline=ACCENT, width=2,
                                 dash=(8,4), fill="")

        # Corner resize handles (use bounding-box corners regardless of shape)
        hs = 6
        for hx,hy in [(cx-ew/2,cy-eh/2),(cx+ew/2,cy-eh/2),
                       (cx-ew/2,cy+eh/2),(cx+ew/2,cy+eh/2)]:
            c.create_rectangle(hx-hs,hy-hs,hx+hs,hy+hs,
                               fill=ACCENT, outline=TEXT_P)

        # Center crosshair
        r=7
        c.create_line(cx-r,cy,cx+r,cy, fill=ACCENT, width=2)
        c.create_line(cx,cy-r,cx,cy+r, fill=ACCENT, width=2)

        ox_pct = round(self._ox_frac*100,1)
        oy_pct = round(self._oy_frac*100,1)
        self._coord_lbl.config(
            text=f"Shape: {SHAPE_LABELS[shape]}   Center offset: X={ox_pct:+.1f}% Y={oy_pct:+.1f}%   "
                 f"Size: {self._rx.get()}% × {self._ry.get()}%")

        self._handle_geo = (cx,cy,ew,eh)

    def _draw_left_trim(self, c, ox_px, oy_px):
        if self._trim_rect_disp:
            x0,y0,x1,y1 = self._trim_rect_disp
            X0,Y0 = ox_px+x0, oy_px+y0
            X1,Y1 = ox_px+x1, oy_px+y1
            c.create_rectangle(X0,Y0,X1,Y1, outline=ACCENT, width=2, dash=(8,4))
            # Corner handles
            hs=6
            for hx,hy in [(X0,Y0),(X1,Y0),(X0,Y1),(X1,Y1)]:
                c.create_rectangle(hx-hs,hy-hs,hx+hs,hy+hs, fill=ACCENT, outline=TEXT_P)
            # Darken the outside area
            self._shade_outside(c, ox_px,oy_px, X0,Y0,X1,Y1)

            rw = (x1-x0)/self._scale
            rh = (y1-y0)/self._scale
            self._coord_lbl.config(
                text=f"Crop box: {int(rw)} × {int(rh)} px "
                     f"(from {self._orig_pil.width} × {self._orig_pil.height})")
        else:
            self._coord_lbl.config(text="No crop box drawn yet — drag on the image")

    def _shade_outside(self, c, ox_px,oy_px, X0,Y0,X1,Y1):
        # Simple 4-rectangle overlay with stipple to create a dark effect
        full_x0,full_y0 = ox_px, oy_px
        full_x1,full_y1 = ox_px+self._dw, oy_px+self._dh
        for (a,b,cc,d) in [
            (full_x0,full_y0,full_x1,Y0),       # top
            (full_x0,Y1,full_x1,full_y1),       # bottom
            (full_x0,Y0,X0,Y1),                  # left
            (X1,Y0,full_x1,Y1),                  # right
        ]:
            if cc>a and d>b:
                c.create_rectangle(a,b,cc,d, fill="#000000", outline="", stipple="gray50")

    def _draw_right(self):
        c = self.right_canvas; c.delete("all")
        try:
            if self.mode == "crop":
                result, _ = crop_shape(
                    self._orig_pil, self._shape_var.get(),
                    self._rx.get()/100, self._ry.get()/100,
                    self._ox_frac, self._oy_frac, self.bg_color)
            else:
                fp = self.files[self.idx]
                rect = self.trim_states[fp]
                if rect:
                    result = trim_background_custom(self._orig_pil, rect)
                else:
                    result = self._orig_pil

            scale  = min(self.PW/result.width, self.PH/result.height, 1.0)
            rw = max(1,int(result.width*scale))
            rh = max(1,int(result.height*scale))
            thumb = result.resize((rw,rh), Image.LANCZOS)
            self._tk_right = ImageTk.PhotoImage(thumb)
            ox = (self.PW-rw)//2; oy = (self.PH-rh)//2
            c.create_image(ox, oy, anchor="nw", image=self._tk_right)
        except Exception as e:
            c.create_text(self.PW//2, self.PH//2, text=str(e), fill=ERROR, font=FS)

    # ── Mouse interaction: crop mode ─────────────────────────────────────────────────
    def _canvas_to_frac(self, event):
        ox_px, oy_px = self._img_offset()
        px = max(0, min(self._dw, event.x - ox_px))
        py = max(0, min(self._dh, event.y - oy_px))
        return px/self._dw - 0.5, py/self._dh - 0.5

    def _hit_handle(self, event):
        if not hasattr(self, "_handle_geo"): return None
        cx,cy,ew,eh = self._handle_geo
        hs = 9
        corners = {
            "tl":(cx-ew/2,cy-eh/2), "tr":(cx+ew/2,cy-eh/2),
            "bl":(cx-ew/2,cy+eh/2), "br":(cx+ew/2,cy+eh/2),
        }
        for name,(hx,hy) in corners.items():
            if abs(event.x-hx)<=hs and abs(event.y-hy)<=hs:
                return name
        return None

    def _on_press(self, event):
        if self.mode == "crop":
            handle = self._hit_handle(event)
            if handle:
                self._drag_mode = "resize"
                self._resize_handle = handle
            else:
                self._drag_mode = "move"
                self._ox_frac, self._oy_frac = self._canvas_to_frac(event)
            self._dragging = True
            self._save_current_state()
            self._redraw()
        else:
            # Trim mode
            ox_px, oy_px = self._img_offset()
            handle = self._hit_trim_handle(event, ox_px, oy_px)
            if handle:
                self._drag_mode = "rect_resize"
                self._resize_handle = handle
            elif self._trim_rect_disp and self._point_in_trim(event, ox_px, oy_px):
                self._drag_mode = "rect_move"
                self._drag_start = (event.x, event.y)
                self._rect_start = self._trim_rect_disp
            else:
                self._drag_mode = "rect_new"
                px = max(0, min(self._dw, event.x-ox_px))
                py = max(0, min(self._dh, event.y-oy_px))
                self._trim_rect_disp = (px,py,px,py)
            self._dragging = True
            self._redraw()

    def _hit_trim_handle(self, event, ox_px, oy_px):
        if not self._trim_rect_disp: return None
        x0,y0,x1,y1 = self._trim_rect_disp
        X0,Y0,X1,Y1 = ox_px+x0, oy_px+y0, ox_px+x1, oy_px+y1
        hs=9
        corners = {"tl":(X0,Y0),"tr":(X1,Y0),"bl":(X0,Y1),"br":(X1,Y1)}
        for name,(hx,hy) in corners.items():
            if abs(event.x-hx)<=hs and abs(event.y-hy)<=hs:
                return name
        return None

    def _point_in_trim(self, event, ox_px, oy_px):
        x0,y0,x1,y1 = self._trim_rect_disp
        X0,Y0,X1,Y1 = ox_px+x0, oy_px+y0, ox_px+x1, oy_px+y1
        return X0<=event.x<=X1 and Y0<=event.y<=Y1

    def _on_drag(self, event):
        if not self._dragging: return
        ox_px, oy_px = self._img_offset()

        if self.mode == "crop":
            if self._drag_mode == "move":
                self._ox_frac, self._oy_frac = self._canvas_to_frac(event)
            elif self._drag_mode == "resize":
                cx,cy,ew,eh = self._handle_geo
                px = max(ox_px, min(ox_px+self._dw, event.x))
                py = max(oy_px, min(oy_px+self._dh, event.y))
                new_ew = abs(px-cx)*2
                new_eh = abs(py-cy)*2
                rx_pct = max(10, min(100, new_ew/self._dw*100))
                ry_pct = max(10, min(100, new_eh/self._dh*100))
                self._rx.set(int(rx_pct))
                self._ry.set(int(ry_pct))
            self._save_current_state()
            self._redraw()
        else:
            px = max(0, min(self._dw, event.x-ox_px))
            py = max(0, min(self._dh, event.y-oy_px))
            if self._drag_mode == "rect_new":
                x0,y0,_,_ = self._trim_rect_disp
                self._trim_rect_disp = (min(x0,px),min(y0,py),max(x0,px),max(y0,py))
            elif self._drag_mode == "rect_move":
                dx = event.x - self._drag_start[0]
                dy = event.y - self._drag_start[1]
                x0,y0,x1,y1 = self._rect_start
                w_,h_ = x1-x0, y1-y0
                nx0 = max(0, min(self._dw-w_, x0+dx))
                ny0 = max(0, min(self._dh-h_, y0+dy))
                self._trim_rect_disp = (nx0,ny0,nx0+w_,ny0+h_)
            elif self._drag_mode == "rect_resize":
                x0,y0,x1,y1 = self._trim_rect_disp
                if self._resize_handle == "tl": x0,y0 = px,py
                elif self._resize_handle == "tr": x1,y0 = px,py
                elif self._resize_handle == "bl": x0,y1 = px,py
                elif self._resize_handle == "br": x1,y1 = px,py
                self._trim_rect_disp = (min(x0,x1),min(y0,y1),max(x0,x1),max(y0,y1))
            self._save_trim_state()
            self._redraw()

    def _on_release(self, event):
        self._dragging = False
        self._drag_mode = None
        self._resize_handle = None

    # ── State management: crop ──────────────────────────────────────────────
    def _save_current_state(self):
        fp = self.files[self.idx]
        self.crop_states[fp] = {
            "shape": self._shape_var.get(),
            "ox": self._ox_frac, "oy": self._oy_frac,
            "rx": self._rx.get(), "ry": self._ry.get(),
        }

    def _update_state_from_sliders(self):
        if not self.files: return
        self._save_current_state()

    def _on_shape_change(self):
        self._save_current_state()
        self._redraw()

    # ── State management: trim ──────────────────────────────────────────────
    def _save_trim_state(self):
        fp = self.files[self.idx]
        if self._trim_rect_disp is None:
            self.trim_states[fp] = None
            return
        x0,y0,x1,y1 = self._trim_rect_disp
        # Convert display coords → original image coords
        ox0 = x0/self._scale; oy0 = y0/self._scale
        ox1 = x1/self._scale; oy1 = y1/self._scale
        self.trim_states[fp] = (ox0,oy0,ox1,oy1)

    def _auto_suggest_trim(self):
        fp = self.files[self.idx]
        trimmed = trim_background(self._orig_pil, self.bg_color, self.trim_tol, self.trim_pad)
        if trimmed.size == self._orig_pil.size:
            # Nothing to remove, or everything is background
            messagebox.showinfo("Auto-Suggest Crop",
                "No background area was detected for removal (the image is already trimmed, "
                "or the background color doesn't match).")
            return
        # Find the cropped bounding box in the original by re-running the mask logic
        arr = np.array(self._orig_pil.convert("RGB"))
        bg_arr = np.array(self.bg_color, dtype=np.int16)
        diff = np.abs(arr.astype(np.int16)-bg_arr).sum(axis=2)
        mask = diff > self.trim_tol*3
        ys,xs = np.where(mask)
        if len(xs)==0:
            return
        x0,x1 = xs.min(), xs.max()+1
        y0,y1 = ys.min(), ys.max()+1
        if self.trim_pad>0:
            bw,bh=x1-x0,y1-y0
            px=int(bw*self.trim_pad/100); py=int(bh*self.trim_pad/100)
            x0=max(0,x0-px); y0=max(0,y0-py)
            x1=min(self._orig_pil.width,x1+px); y1=min(self._orig_pil.height,y1+py)

        self.trim_states[fp] = (x0,y0,x1,y1)
        self._trim_rect_disp = (x0*self._scale,y0*self._scale,
                                x1*self._scale,y1*self._scale)
        self._redraw()

    def _reset_current(self):
        fp = self.files[self.idx]
        if self.mode == "crop":
            self._ox_frac = 0.0; self._oy_frac = 0.0
            self.crop_states[fp] = {
                "shape": self._shape_var.get(),
                "ox":0.0,"oy":0.0,
                "rx": self._rx.get(), "ry": self._ry.get(),
            }
        else:
            self._trim_rect_disp = None
            self.trim_states[fp] = None
        self._redraw()

    # ── Navigation ────────────────────────────────────────────────────────
    def _prev(self):
        if self.mode=="crop": self._save_current_state()
        else: self._save_trim_state()
        if self.idx > 0:
            self.idx -= 1; self._load_current()

    def _next(self):
        if self.mode=="crop": self._save_current_state()
        else: self._save_trim_state()
        if self.idx < len(self.files)-1:
            self.idx += 1; self._load_current()

    # ── Save all ──────────────────────────────────────────────────────────────
    def _save_all(self):
        if self.mode=="crop": self._save_current_state()
        else: self._save_trim_state()

        Path(self.out_dir).mkdir(parents=True, exist_ok=True)
        total = len(self.files); done = 0; errors = 0
        self._prog_bar.config(maximum=total)

        for fp in self.files:
            try:
                with Image.open(fp) as _i:
                    orig_fmt = (_i.format or "PNG").upper()
                    orig_info= _i.info.copy()
                    img = _i.copy()

                if self.mode == "crop":
                    st = self.crop_states[fp]
                    img, _ = crop_shape(img, st["shape"], st["rx"]/100, st["ry"]/100,
                                        st["ox"], st["oy"], self.bg_color)
                else:
                    rect = self.trim_states[fp]
                    if rect:
                        img = trim_background_custom(img, rect)

                img = apply_resize(img, self.resize_mode, self.tw, self.th)

                sfmt = (self.out_fmt or orig_fmt).upper()
                if sfmt=="JPG": sfmt="JPEG"
                img  = prep_mode(img, sfmt)
                ext  = FMT_EXT.get(sfmt, Path(fp).suffix)
                op   = Path(self.out_dir)/(Path(fp).stem+ext)

                if self.lim_size and self.target_kb:
                    compress_to_kb(img, str(op), sfmt, self.target_kb)
                else:
                    img.save(str(op), format=sfmt, **save_kw(sfmt, orig_info))

                done += 1
            except Exception:
                errors += 1

            self._prog_var.set(done+errors)
            self._prog_lbl.config(text=f"Saving {done+errors}/{total}…", fg=TEXT_S)
            self.update_idletasks()

        self._prog_lbl.config(
            text=f"✔  {done} saved,  {errors} error(s)  →  {self.out_dir}",
            fg=SUCCESS if errors==0 else WARNING)
        messagebox.showinfo("Done", f"{done}/{total} image(s) saved to:\n{self.out_dir}")
        if errors == 0:
            self.destroy()


    # ── Export current overrides (for the main app to use if needed) ─────────────
    def get_crop_overrides(self):
        """Return dict suitable for cmap in process_image."""
        out = {}
        for fp, st in self.crop_states.items():
            out[fp] = {
                "shape": st["shape"], "rx": st["rx"], "ry": st["ry"],
                "ox": st["ox"], "oy": st["oy"], "shape_kw": {},
            }
        return out

    def get_trim_overrides(self):
        out = {}
        for fp, rect in self.trim_states.items():
            if rect:
                out[fp] = {"trim_rect": rect, "do_trim": True}
        return out

# ─────────────────────────────────────────────────────────────────────────────
#  Main application
# ─────────────────────────────────────────────────────────────────────────────
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Batch Image Resizer  v8")
        self.root.geometry("1080x780")
        self.root.configure(bg=DARK_BG)
        self.root.minsize(960, 680)

        self.files       = []
        self.output_dir  = tk.StringVar()
        self.width_var   = tk.StringVar(value="413")
        self.height_var  = tk.StringVar(value="531")
        self.mode_var    = tk.StringVar(value="exact")
        self.thread_var  = tk.IntVar(value=min(8, os.cpu_count() or 4))
        self.keep_struct = tk.BooleanVar(value=False)

        # Shape Crop
        self.crop_on     = tk.BooleanVar(value=False)
        self.shape_var   = tk.StringVar(value="ellipse")
        self.c_rx        = tk.IntVar(value=90)
        self.c_ry        = tk.IntVar(value=90)
        self.c_ox        = tk.IntVar(value=0)
        self.c_oy        = tk.IntVar(value=0)
        self.c_bg_rgb    = (255,255,255)
        self.c_bg_hex    = tk.StringVar(value="#ffffff")
        # Extra shape parameters
        self.star_points = tk.IntVar(value=5)
        self.star_inner  = tk.IntVar(value=50)  # percent

        # Trim edges
        self.trim_on     = tk.BooleanVar(value=False)
        self.trim_tol    = tk.IntVar(value=10)
        self.trim_pad    = tk.IntVar(value=0)

        self.manual_map  = {}   # {path: override dict}

        self.size_on     = tk.BooleanVar(value=False)
        self.size_kb_var = tk.StringVar(value="200")
        self.size_unit   = tk.StringVar(value="KB")
        self.out_fmt_var = tk.StringVar(value="Original")

        self.processing  = False
        self.log_q       = queue.Queue()

        # ── Voice tab state ──────────────────────────────────────────────
        self.voice_q          = queue.Queue()
        self.voice_listening  = False
        self.voice_recognizer = None
        self.voice_mic        = None
        self.voice_stop_listening = None
        self.voice_lang       = tk.StringVar(value="fa-IR")
        self.voice_engine     = tk.StringVar(value="online")   # "online" or "offline"
        self.voice_model_path = tk.StringVar()
        self.voice_numbers_to_digits = tk.BooleanVar(value=False)
        self.vosk_model       = None
        self.vosk_model_loaded_path = None
        self.voice_export_fmt   = tk.StringVar(value="xlsx")
        self.voice_excel_layout = tk.StringVar(value="column")
        self.voice_pptx_layout  = tk.StringVar(value="bullets")
        self.voice_out_path     = tk.StringVar()

        self._styles()
        self._build()
        self._poll()
        self._voice_poll()

    # ── Styles ────────────────────────────────────────────────────────────────
    def _styles(self):
        s = ttk.Style(); s.theme_use("clam")
        s.configure("TFrame",       background=DARK_BG)
        s.configure("TProgressbar", background=ACCENT, troughcolor=BORDER_L,
                    borderwidth=0, thickness=7)
        s.configure("TCheckbutton", background=CARD_BG, foreground=TEXT_S, font=FS,
                    indicatorcolor=BORDER_L, selectcolor=ACCENT)
        s.map("TCheckbutton",
              background=[("active", CARD_BG)],
              foreground=[("active", TEXT_P)],
              indicatorcolor=[("selected", ACCENT), ("!selected", BORDER_L)])
        s.configure("TRadiobutton", background=CARD_BG, foreground=TEXT_S, font=FB,
                    indicatorcolor=BORDER_L)
        s.map("TRadiobutton",
              background=[("active", CARD_BG)],
              foreground=[("active", TEXT_P)],
              indicatorcolor=[("selected", ACCENT), ("!selected", BORDER_L)])
        s.configure("TNotebook",     background=DARK_BG, borderwidth=0)
        s.configure("TNotebook.Tab", background=PANEL_BG, foreground=TEXT_S,
                    font=FB, padding=(14, 8))
        s.map("TNotebook.Tab",
              background=[("selected", CARD_BG)],
              foreground=[("selected", ACCENT_L)])
        s.configure("TScrollbar",    background=PANEL_BG, troughcolor=DARK_BG,
                    borderwidth=0, arrowcolor=TEXT_S)
        s.configure("Vertical.TScrollbar", width=6)

    def _card(self, parent, title="", icon="", pady=(0,12)):
        outer = tk.Frame(parent, bg=CARD_BG,
                         highlightbackground=BORDER_L, highlightthickness=1)
        outer.pack(fill="x", pady=pady)
        if title:
            hdr = tk.Frame(outer, bg=CARD_BG)
            hdr.pack(fill="x", padx=0, pady=(0,0))
            # Emerald accent bar on the left
            tk.Frame(hdr, bg=ACCENT, width=3).pack(side="left", fill="y", pady=(8,0))
            txt = f"{icon}  {title}" if icon else title
            tk.Label(hdr, text=txt,
                     font=FH, bg=CARD_BG, fg=TEXT_P).pack(
                         anchor="w", padx=12, pady=(10,4), side="left")
        inner = tk.Frame(outer, bg=CARD_BG)
        inner.pack(fill="x", padx=16, pady=(4,14))
        return inner

    def _sl(self, parent, label, var, lo, hi, suffix="%", w=20, cb=None):
        row = tk.Frame(parent, bg=CARD_BG); row.pack(fill="x", pady=3)
        tk.Label(row, text=label, font=FS, bg=CARD_BG,
                 fg=TEXT_S, width=w, anchor="w").pack(side="left")
        lbl = tk.Label(row, text=f"{var.get()}{suffix}",
                       font=(_PFONT,9,"bold"), bg=CARD_BG,
                       fg=ACCENT_L, width=7)
        lbl.pack(side="right")
        def _cmd(v):
            lbl.config(text=f"{int(float(v))}{suffix}")
            if cb: cb()
        tk.Scale(row, from_=lo, to=hi, variable=var,
                 orient="horizontal", bg=CARD_BG, fg=ACCENT,
                 troughcolor=BORDER_L, activebackground=ACCENT_L,
                 sliderrelief="flat", highlightthickness=0,
                 bd=0, showvalue=False,
                 command=_cmd).pack(side="left", fill="x", expand=True, padx=(0,6))

    # ── Build UI ─────────────────────────────────────────────────────────
    def _build(self):
        hdr = tk.Frame(self.root, bg=PANEL_BG, height=58)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        # Emerald accent stripe at top of header
        tk.Frame(hdr, bg=ACCENT, height=3).place(relx=0, rely=0, relwidth=1)
        inner = tk.Frame(hdr, bg=PANEL_BG)
        inner.place(relx=.5, rely=.55, anchor="center")
        tk.Label(inner, text="🌿", font=(_PFONT,17),
                 bg=PANEL_BG, fg=ACCENT).pack(side="left", padx=(0,8))
        tk.Label(inner, text="Batch Image Resizer Tool",
                 font=FT, bg=PANEL_BG, fg=TEXT_P).pack(side="left")
        tk.Label(inner, text="  v10",
                 font=(_PFONT,10), bg=PANEL_BG, fg=TEXT_S).pack(side="left")
        tk.Frame(self.root, bg=BORDER_L, height=1).pack(fill="x")

        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True)

        tabs = {}
        for key, label in [
            ("images",  "  🖼  Images  "),
            ("resize",  "  ↔  Resize  "),
            ("crop",    "  ✂  Shape Crop  "),
            ("trim",    "  ▭  Trim Edges  "),
            ("session", "  🎨  Editor  "),
            ("voice",   "  🎙  Voice Dictation  "),
        ]:
            outer = tk.Frame(nb, bg=DARK_BG)
            nb.add(outer, text=label)
            canvas = tk.Canvas(outer, bg=DARK_BG, highlightthickness=0)
            sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
            canvas.configure(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")
            canvas.pack(side="left", fill="both", expand=True)
            inner2 = tk.Frame(canvas, bg=DARK_BG)
            win_id = canvas.create_window((0,0), window=inner2, anchor="nw")
            def _resize(e, c=canvas, wid=win_id):
                c.itemconfig(wid, width=e.width)
            canvas.bind("<Configure>", _resize)
            inner2.bind("<Configure>",
                lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
            # Scroll with mouse wheel
            def _wheel(e, c=canvas):
                c.yview_scroll(int(-1*(e.delta/120)), "units")
            canvas.bind_all("<MouseWheel>", _wheel) if key=="images" else None
            tabs[key] = inner2

        self._build_images(tabs["images"])
        self._build_resize(tabs["resize"])
        self._build_crop(tabs["crop"])
        self._build_trim(tabs["trim"])
        self._build_session(tabs["session"])
        self._build_voice(tabs["voice"])

        sep = tk.Frame(self.root, bg=BORDER, height=1)
        sep.pack(fill="x", side="bottom")
        bot = tk.Frame(self.root, bg=PANEL_BG, height=76)
        bot.pack(fill="x", side="bottom"); bot.pack_propagate(False)
        self._build_bottom(bot)

    # ── Tab: Images ───────────────────────────────────────────────────────────
    def _build_images(self, p):
        wrap = tk.Frame(p, bg=DARK_BG)
        wrap.pack(fill="both", expand=True, padx=20, pady=14)

        fc = self._card(wrap, "Images", "🖼")
        row = tk.Frame(fc, bg=CARD_BG); row.pack(fill="x", pady=(0,10))
        # Primary button (filled emerald)
        AnimatedButton(row, text="＋  Add Images", font=FB,
                       bg=ACCENT, fg=TEXT_P,
                       hover_bg=ACCENT_L, press_bg=ACCENT_H,
                       padx=14, pady=7,
                       command=self._add_files).pack(side="left", padx=(0,6), ipadx=2)

        # Folder button (secondary)
        b2 = AnimatedButton(row, text="📁  Add Folder", font=FB,
                            bg=BORDER_L, fg=TEXT_P,
                            hover_bg=_lighten(BORDER_L, 18),
                            press_bg=BORDER,
                            padx=12, pady=7, command=self._add_folder)
        b2.pack(side="left", padx=(0,6))

        # Clear button (danger)
        b3 = AnimatedButton(row, text="✕  Clear All", font=FB,
                            bg=CARD_BG, fg=ERROR,
                            hover_bg=_mix(CARD_BG, ERROR, 0.15),
                            press_bg=_mix(CARD_BG, ERROR, 0.25),
                            padx=12, pady=7, command=self._clear_files)
        b3.pack(side="left")

        lf = tk.Frame(fc, bg=PANEL_BG,
                      highlightbackground=BORDER_L, highlightthickness=1)
        lf.pack(fill="x")
        sb = ttk.Scrollbar(lf, orient="vertical")
        sb.pack(side="right", fill="y")
        self.file_lb = tk.Listbox(lf, bg=PANEL_BG, fg=TEXT_S,
                                   selectbackground=ACCENT,
                                   selectforeground=DARK_BG,
                                   font=FM, bd=0, height=8,
                                   yscrollcommand=sb.set, activestyle="none")
        self.file_lb.pack(fill="x")
        sb.config(command=self.file_lb.yview)
        self.fc_lbl = tk.Label(fc, text="No images selected",
                                font=FS, bg=CARD_BG, fg=TEXT_S)
        self.fc_lbl.pack(anchor="w", pady=(6,0))

        oc = self._card(wrap, "Output Folder", "📂")
        row2 = tk.Frame(oc, bg=CARD_BG); row2.pack(fill="x")
        _styled_entry(row2, textvariable=self.output_dir, ipady=6)
        AnimatedButton(row2, text="Browse…", font=FS,
                       bg=BORDER_L, fg=TEXT_P,
                       hover_bg=ACCENT, press_bg=ACCENT_H,
                       padx=10, pady=6,
                       command=self._pick_out).pack(side="left", padx=(7,0))
        ttk.Checkbutton(oc, text="Preserve original folder structure",
                        variable=self.keep_struct).pack(anchor="w", pady=(8,0))

    # ── Tab: Shape Crop ───────────────────────────────────────────────────────
    def _build_crop(self, p):
        wrap = tk.Frame(p, bg=DARK_BG)
        wrap.pack(fill="both", expand=True, padx=20, pady=14)

        tog = tk.Frame(wrap, bg=CARD_BG,
                       highlightbackground=ACCENT, highlightthickness=2)
        tog.pack(fill="x", pady=(0,10))
        ti  = tk.Frame(tog, bg=CARD_BG)
        ti.pack(fill="x", padx=14, pady=12)
        ttk.Checkbutton(ti, text="  ✂  Enable shape crop (batch)",
                        variable=self.crop_on,
                        command=self._tog_crop).pack(side="left")
        tk.Label(ti,
                 text="\nCrops each image to the chosen shape and fills the area\n"
                      "outside the shape with the background color. For fine-tuning\n"
                      "each image individually, use the \"Editor\" tab.",
                 font=FS, bg=CARD_BG, fg=TEXT_S, justify="left").pack(anchor="w")

        self.c_frame = tk.Frame(wrap, bg=DARK_BG)
        if self.crop_on.get():
            self.c_frame.pack(fill="x")

        # Shape selector
        sh = self._card(self.c_frame, "Shape", "🔷")
        sr = tk.Frame(sh, bg=CARD_BG); sr.pack(anchor="w")
        for shp in SHAPES:
            tk.Radiobutton(sr, text=SHAPE_LABELS[shp], variable=self.shape_var, value=shp,
                           bg=CARD_BG, fg=TEXT_S, selectcolor=DARK_BG,
                           activebackground=CARD_BG, activeforeground=ACCENT_L,
                           font=FB, bd=0, cursor="hand2",
                           command=self._draw_prev).pack(side="left", padx=(0,10))

        # Star-only options
        self.star_opts = tk.Frame(sh, bg=CARD_BG)
        self.star_opts.pack(fill="x", pady=(8,0))
        self._sl(self.star_opts, "Number of star points", self.star_points, 5, 10, suffix="", cb=self._draw_prev)
        self._sl(self.star_opts, "Star inner ratio", self.star_inner, 20, 80, cb=self._draw_prev)

        # Size
        sz = self._card(self.c_frame, "Shape Size", "↔")
        self._sl(sz, "Width coverage",  self.c_rx, 10,100, cb=self._draw_prev)
        self._sl(sz, "Height coverage", self.c_ry, 10,100, cb=self._draw_prev)

        # Offset
        off = self._card(self.c_frame, "Center Offset  (0 = image center)", "🎯")
        tk.Label(off, text="Negative = left/top   Positive = right/bottom",
                 font=FS, bg=CARD_BG, fg=TEXT_S).pack(anchor="w", pady=(0,6))
        self._sl(off, "Horizontal X", self.c_ox, -40,40, cb=self._draw_prev)
        self._sl(off, "Vertical Y",   self.c_oy, -40,40, cb=self._draw_prev)
        AnimatedButton(off, text="↺  Reset to Center", font=FS,
                  bg=PANEL_BG, fg=TEXT_S,
                  hover_bg=_lighten(PANEL_BG,20), press_bg=_darken(PANEL_BG,10),
                  padx=8, pady=3,
                  command=lambda:(self.c_ox.set(0),self.c_oy.set(0),
                                  self._draw_prev())).pack(anchor="w",pady=(6,0))

        # Preview (thumbnail)
        prev = self._card(self.c_frame, "Preview", "👁")
        self.prev_c = tk.Canvas(prev, width=240, height=150,
                                bg=PANEL_BG, highlightthickness=0)
        self.prev_c.pack(anchor="w")
        self._draw_prev()

        # Presets
        pre = self._card(self.c_frame, "Presets", "📐")
        pr = tk.Frame(pre, bg=CARD_BG); pr.pack(anchor="w")
        for nm,rx,ry in [("Circle/Square",100,100),("Wide",95,75),
                          ("Tall",75,95),("Compact",80,80)]:
            def _a(rx=rx,ry=ry):
                self.c_rx.set(rx); self.c_ry.set(ry); self._draw_prev()
            AnimatedButton(pr,text=nm,font=FS,bg=PANEL_BG,fg=TEXT_S,
                      hover_bg=ACCENT, press_bg=ACCENT_H,
                      padx=8,pady=4,
                      command=_a).pack(side="left",padx=(0,6))

        # Background color
        bgc = self._card(self.c_frame, "Background Color (outside shape)", "🎨")
        br  = tk.Frame(bgc, bg=CARD_BG); br.pack(anchor="w")
        self.bg_sw = tk.Label(br, bg=self.c_bg_hex.get(),
                              width=4, relief="solid", bd=1)
        self.bg_sw.pack(side="left",ipady=8,padx=(0,8))
        tk.Label(br,text="Hex:",font=FS,bg=CARD_BG,fg=TEXT_S).pack(side="left")
        tk.Entry(br,textvariable=self.c_bg_hex,width=9,
                 font=FM,bg=PANEL_BG,fg=ACCENT_L,
                 insertbackground=ACCENT,bd=1,relief="solid",
                 justify="center").pack(side="left",ipady=4,padx=(4,8))
        AnimatedButton(br,text="Apply",font=FS,bg=BORDER_L,fg=TEXT_S,
                  activebackground=ACCENT,activeforeground=TEXT_P,
                  bd=0,padx=8,pady=4,cursor="hand2",
                  command=self._apply_bg).pack(side="left")

        self._tog_crop()
        self._on_shape_var_change()

    def _on_shape_var_change(self, *_):
        if self.shape_var.get() == "star":
            self.star_opts.pack(fill="x", pady=(8,0))
        else:
            self.star_opts.pack_forget()

    # ── Tab: Trim Edges ───────────────────────────────────────────────────────
    def _build_trim(self, p):
        wrap = tk.Frame(p, bg=DARK_BG)
        wrap.pack(fill="both", expand=True, padx=20, pady=14)

        tog = tk.Frame(wrap, bg=CARD_BG,
                       highlightbackground=ACCENT, highlightthickness=2)
        tog.pack(fill="x", pady=(0,10))
        ti  = tk.Frame(tog, bg=CARD_BG); ti.pack(fill="x", padx=14, pady=12)
        ttk.Checkbutton(ti, text="  ▭  Enable automatic edge trimming (batch)",
                        variable=self.trim_on,
                        command=self._tog_trim).pack(side="left")
        tk.Label(ti,
                 text="\nWhen you crop to a shape, an empty border in the background color remains around it.\n"
                      "This option finds that extra space and removes it — before resizing.\n"
                      "To manually draw a crop box on each image, use the \"Editor\" tab.",
                 font=FS, bg=CARD_BG, fg=TEXT_S, justify="left").pack(anchor="w")

        opts_outer = tk.Frame(wrap, bg=CARD_BG,
                       highlightbackground=BORDER_L, highlightthickness=1)
        opts_outer.pack(fill="x", pady=(0,10))
        tk.Label(opts_outer, text="⚙  Trim Settings",
                 font=FH, bg=CARD_BG, fg=TEXT_P).pack(
                     anchor="w", padx=14, pady=(10,0))
        opts = tk.Frame(opts_outer, bg=CARD_BG)
        opts.pack(fill="x", padx=14, pady=(4,12))
        self.trim_opts_frame = opts_outer
        self._sl(opts, "Color sensitivity", self.trim_tol, 0, 100, suffix="")
        tk.Label(opts, text="Higher sensitivity → colors close to the background get removed too.",
                 font=FS, bg=CARD_BG, fg=TEXT_S).pack(anchor="w", pady=(0,6))
        self._sl(opts, "Extra padding", self.trim_pad, 0, 50)
        tk.Label(opts, text="Adds a small margin around the detected area (percent).",
                 font=FS, bg=CARD_BG, fg=TEXT_S).pack(anchor="w")

        info_outer = tk.Frame(wrap, bg=CARD_BG,
                       highlightbackground=BORDER_L, highlightthickness=1)
        info_outer.pack(fill="x", pady=(0,0))
        info = tk.Frame(info_outer, bg=CARD_BG)
        info.pack(fill="x", padx=14, pady=(4,12))
        self.trim_info_frame = info_outer
        tk.Label(info,
                 text="ℹ  This tab uses the same background color as the \"Shape Crop\" tab.\n"
                      "   Processing order: shape crop ← trim edges ← resize ← save.",
                 font=FS, bg=CARD_BG, fg=WARNING, justify="left").pack(anchor="w")

        self._tog_trim()

    # ── Tab: Session Editor ───────────────────────────────────────────────────
    def _build_session(self, p):
        wrap = tk.Frame(p, bg=DARK_BG)
        wrap.pack(fill="both", expand=True, padx=20, pady=14)

        info = self._card(wrap, "Image Editor", "🎨")
        tk.Label(info,
                 text="Open the images in a folder one by one and work on each manually.\n"
                      "You have two modes:\n"
                      "  • Shape crop — choose a shape, draw and resize it on the image\n"
                      "  • Trim edges — draw a custom box to remove extra space around the image\n"
                      "After editing all images, you save them all at once.",
                 font=FS, bg=CARD_BG, fg=TEXT_S, justify="left").pack(anchor="w", pady=(0,10))

        # Input folder
        r1 = tk.Frame(info, bg=CARD_BG); r1.pack(fill="x", pady=(0,6))
        tk.Label(r1,text="Input folder:",font=FB,bg=CARD_BG,fg=TEXT_S,
                 width=14,anchor="w").pack(side="left")
        self.ses_in  = tk.StringVar()
        _e1, _f1 = _styled_entry(r1, textvariable=self.ses_in, ipady=5)
        AnimatedButton(r1,text="Browse…",font=FS,bg=BORDER_L,fg=TEXT_S,
                  hover_bg=ACCENT, press_bg=ACCENT_H,
                  padx=10,pady=5,
                  command=lambda:self.ses_in.set(
                      filedialog.askdirectory(title="Input Folder") or self.ses_in.get()
                  )).pack(side="left",padx=(6,0))

        # Output folder
        r2 = tk.Frame(info, bg=CARD_BG); r2.pack(fill="x", pady=(0,6))
        tk.Label(r2,text="Output Folder:",font=FB,bg=CARD_BG,fg=TEXT_S,
                 width=14,anchor="w").pack(side="left")
        self.ses_out = tk.StringVar()
        _e2, _f2 = _styled_entry(r2, textvariable=self.ses_out, ipady=5)
        AnimatedButton(r2,text="Browse…",font=FS,bg=BORDER_L,fg=TEXT_S,
                  hover_bg=ACCENT, press_bg=ACCENT_H,
                  padx=10,pady=5,
                  command=lambda:self.ses_out.set(
                      filedialog.askdirectory(title="Output Folder") or self.ses_out.get()
                  )).pack(side="left",padx=(6,0))

        # Mode selection
        mr = tk.Frame(info, bg=CARD_BG); mr.pack(fill="x", pady=(6,0))
        tk.Label(mr, text="Edit mode:", font=FB, bg=CARD_BG, fg=TEXT_S,
                 width=14, anchor="w").pack(side="left")
        self.ses_mode = tk.StringVar(value="crop")
        tk.Radiobutton(mr, text="✂  Shape Crop", variable=self.ses_mode, value="crop",
                       bg=CARD_BG, fg=TEXT_S, selectcolor=DARK_BG,
                       activebackground=CARD_BG, activeforeground=ACCENT_L,
                       font=FB, bd=0, cursor="hand2",
                       command=self._tog_session_opts).pack(side="left", padx=(0,16))
        tk.Radiobutton(mr, text="▭  Trim Edges", variable=self.ses_mode, value="trim",
                       bg=CARD_BG, fg=TEXT_S, selectcolor=DARK_BG,
                       activebackground=CARD_BG, activeforeground=ACCENT_L,
                       font=FB, bd=0, cursor="hand2",
                       command=self._tog_session_opts).pack(side="left")

        # Options row (crop defaults)
        self.ses_crop_opts = tk.Frame(info, bg=CARD_BG)
        self.ses_crop_opts.pack(fill="x", pady=(8,0))

        sf = tk.Frame(self.ses_crop_opts, bg=CARD_BG); sf.pack(side="left", padx=(0,16))
        tk.Label(sf,text="Default shape:",font=FS,bg=CARD_BG,fg=TEXT_S).pack(anchor="w")
        self.ses_shape = tk.StringVar(value="ellipse")
        ttk.Combobox(sf,textvariable=self.ses_shape,
                     values=SHAPES, width=11, state="readonly").pack(anchor="w")

        ef = tk.Frame(self.ses_crop_opts, bg=CARD_BG); ef.pack(side="left", padx=(0,16))
        tk.Label(ef,text="Default width %:",font=FS,bg=CARD_BG,fg=TEXT_S).pack(anchor="w")
        self.ses_rx = tk.IntVar(value=90)
        tk.Spinbox(ef,from_=10,to=100,textvariable=self.ses_rx,
                   width=5,font=FB,bg=PANEL_BG,fg=ACCENT,
                   buttonbackground=BORDER,relief="solid",bd=1).pack(anchor="w")

        eff = tk.Frame(self.ses_crop_opts, bg=CARD_BG); eff.pack(side="left", padx=(0,16))
        tk.Label(eff,text="Default height %:",font=FS,bg=CARD_BG,fg=TEXT_S).pack(anchor="w")
        self.ses_ry = tk.IntVar(value=90)
        tk.Spinbox(eff,from_=10,to=100,textvariable=self.ses_ry,
                   width=5,font=FB,bg=PANEL_BG,fg=ACCENT,
                   buttonbackground=BORDER,relief="solid",bd=1).pack(anchor="w")

        # Format and resize (shared)
        common = tk.Frame(info, bg=CARD_BG); common.pack(fill="x", pady=(8,0))
        fmf = tk.Frame(common, bg=CARD_BG); fmf.pack(side="left", padx=(0,16))
        tk.Label(fmf,text="Output format:",font=FS,bg=CARD_BG,fg=TEXT_S).pack(anchor="w")
        self.ses_fmt = tk.StringVar(value="Original")
        ttk.Combobox(fmf,textvariable=self.ses_fmt,
                     values=["Original"]+FORMATS,
                     width=9,state="readonly").pack(anchor="w")

        rmf = tk.Frame(common, bg=CARD_BG); rmf.pack(side="left", padx=(0,16))
        tk.Label(rmf,text="Resize mode (after editing):",
                 font=FS,bg=CARD_BG,fg=TEXT_S).pack(anchor="w")
        self.ses_resize_mode = tk.StringVar(value="no_resize")
        ttk.Combobox(rmf,textvariable=self.ses_resize_mode,
                     values=["no_resize","exact","fit","fill","width_only","height_only"],
                     width=12,state="readonly").pack(anchor="w")

        # Launch button
        AnimatedButton(info, text="🎨  Open Editor",
                  font=(_PFONT,13,"bold"),
                  bg=ACCENT, fg=TEXT_P,
                  hover_bg=ACCENT_L, press_bg=ACCENT_H,
                  padx=24, pady=12,
                  command=self._launch_session).pack(anchor="w", pady=(14,0))

        tk.Label(info,
                 text="ℹ  The session editor is independent — it has its own input/output folders,\n"
                      "   does not use the file list from the Images tab, and uses the resize mode /\n"
                      "   dimensions / output format / file size settings below.",
                 font=FS, bg=CARD_BG, fg=WARNING, justify="left").pack(
                     anchor="w", pady=(10,0))

        self._tog_session_opts()

    def _tog_session_opts(self):
        if self.ses_mode.get() == "crop":
            self.ses_crop_opts.pack(fill="x", pady=(8,0))
        else:
            self.ses_crop_opts.pack_forget()

    # ── Tab: Resize ───────────────────────────────────────────────────────
    def _build_resize(self, p):
        wrap = tk.Frame(p, bg=DARK_BG)
        wrap.pack(fill="both", expand=True, padx=20, pady=14)

        fc = self._card(wrap, "Output Format", "📄")
        tk.Label(fc,
                 text="Saves images in the chosen format. The file extension changes automatically too.",
                 font=FS,bg=CARD_BG,fg=TEXT_S).pack(anchor="w",pady=(0,6))
        fr = tk.Frame(fc,bg=CARD_BG); fr.pack(anchor="w")
        for v in ["Original"]+FORMATS:
            tk.Radiobutton(fr,text=v,variable=self.out_fmt_var,value=v,
                           bg=CARD_BG,fg=TEXT_S,selectcolor=DARK_BG,
                           activebackground=CARD_BG,activeforeground=ACCENT_L,
                           font=FB,bd=0,cursor="hand2").pack(side="left",padx=(0,10))
        tk.Label(fc,text="⚠ JPEG and BMP formats have no transparency (alpha is removed).",
                 font=FS,bg=CARD_BG,fg=WARNING).pack(anchor="w",pady=(4,0))

        px = self._card(wrap,"Target Dimensions (pixels)","↔")
        pr = tk.Frame(px,bg=CARD_BG); pr.pack(anchor="w")
        def _pf(p,lbl,var):
            f=tk.Frame(p,bg=CARD_BG); f.pack(side="left",padx=(0,16))
            tk.Label(f,text=lbl,font=FS,bg=CARD_BG,fg=TEXT_S).pack(anchor="w")
            _styled_entry(f, textvariable=var, width=7,
                         font=(_PFONT,14,"bold"), justify="center",
                         ipady=5, fill=None, expand=False)
            tk.Label(f,text="px",font=FS,bg=CARD_BG,fg=TEXT_S).pack(anchor="e")
        _pf(pr,"Width",self.width_var)
        tk.Label(pr,text="×",font=(_PFONT,16),
                 bg=CARD_BG,fg=TEXT_S).pack(side="left",padx=4,pady=(14,0))
        _pf(pr,"Height",self.height_var)

        mc = self._card(wrap,"Resize Mode (batch)","⚙")
        modes=[("exact","Exact","Stretch to W×H"),
               ("fit","Fit","Aspect ratio kept"),
               ("fill","Fill","Fill + crop from center"),
               ("width_only","Width Only","Based on width"),
               ("height_only","Height Only","Based on height"),
               ("no_resize","No Resize","Save as-is")]
        mr=tk.Frame(mc,bg=CARD_BG); mr.pack(fill="x")
        for val,lbl,tip in modes:
            col=tk.Frame(mr,bg=CARD_BG); col.pack(side="left",padx=(0,8))
            tk.Radiobutton(col,text=lbl,variable=self.mode_var,value=val,
                           bg=CARD_BG,fg=TEXT_S,selectcolor=DARK_BG,
                           activebackground=CARD_BG,activeforeground=ACCENT_L,
                           font=FS,bd=0,cursor="hand2").pack(anchor="w")
            tk.Label(col,text=tip,font=(_PFONT,7),
                     bg=CARD_BG,fg=TEXT_S).pack(anchor="w")

        tc=self._card(wrap,"Processing Speed","⚡")
        tk.Label(tc,text="Number of parallel processes (more = faster):",font=FS,bg=CARD_BG,fg=TEXT_S).pack(anchor="w")
        tr=tk.Frame(tc,bg=CARD_BG); tr.pack(fill="x",pady=(2,0))
        self.th_lbl=tk.Label(tr,text=str(self.thread_var.get()),
                              font=(_PFONT,14,"bold"),
                              bg=CARD_BG,fg=ACCENT_L,width=3)
        self.th_lbl.pack(side="right")
        tk.Scale(tr,from_=1,to=min(32,(os.cpu_count()or 4)*2),
                 variable=self.thread_var,orient="horizontal",
                 bg=CARD_BG,fg=ACCENT,troughcolor=BORDER_L,
                 activebackground=ACCENT_L,
                 sliderrelief="flat",highlightthickness=0,bd=0,showvalue=False,
                 command=lambda v:self.th_lbl.config(
                     text=str(int(float(v))))).pack(
                         side="left",fill="x",expand=True)

        so=tk.Frame(wrap,bg=CARD_BG,highlightbackground=BORDER,highlightthickness=1)
        so.pack(fill="x",pady=(0,10))
        sh=tk.Frame(so,bg=CARD_BG); sh.pack(fill="x",padx=14,pady=(12,4))
        ttk.Checkbutton(sh,text="  📦  Limit output file size",
                        variable=self.size_on,
                        command=self._tog_size).pack(side="left")
        self.sz_f=tk.Frame(so,bg=CARD_BG)
        if self.size_on.get():
            self.sz_f.pack(fill="x",padx=14,pady=(0,12))
        tk.Label(self.sz_f,text="Compresses each image below a set size.",
                 font=FS,bg=CARD_BG,fg=TEXT_S).pack(anchor="w")
        tk.Frame(self.sz_f,bg=BORDER_L,height=1).pack(fill="x",pady=4)
        sr=tk.Frame(self.sz_f,bg=CARD_BG); sr.pack(anchor="w")
        tk.Label(sr,text="Max size:",font=FB,bg=CARD_BG,fg=TEXT_S).pack(side="left",padx=(0,6))
        self.sz_e, _szf = _styled_entry(sr, textvariable=self.size_kb_var, width=7,
                         font=(_PFONT,13,"bold"), justify="center",
                         ipady=5, fill=None, expand=False)
        self.kb_btn=AnimatedButton(sr,text="KB",font=FB,bg=ACCENT,fg=TEXT_P,
                               hover_bg=ACCENT_L, press_bg=ACCENT_H,
                               padx=8,pady=5,
                               command=lambda:self._set_unit("KB"))
        self.kb_btn.pack(side="left",padx=(6,2))
        self.mb_btn=AnimatedButton(sr,text="MB",font=FB,bg=BORDER_L,fg=TEXT_S,
                               hover_bg=_lighten(BORDER_L,18), press_bg=BORDER,
                               padx=8,pady=5,
                               command=lambda:self._set_unit("MB"))
        self.mb_btn.pack(side="left")
        self._tog_size()

    # ── Bottom bar ────────────────────────────────────────────────────────────
    def _build_bottom(self, parent):
        left=tk.Frame(parent,bg=PANEL_BG)
        left.pack(side="left",fill="both",expand=True,padx=16,pady=14)
        pr=tk.Frame(left,bg=PANEL_BG); pr.pack(fill="x")
        self.prog_var=tk.DoubleVar()
        ttk.Progressbar(pr,variable=self.prog_var,
                        maximum=100,style="TProgressbar").pack(
                            fill="x",expand=True,side="left")
        self.pct_lbl=tk.Label(pr,text="0%",font=FB,
                               bg=PANEL_BG,fg=ACCENT_L,width=5)
        self.pct_lbl.pack(side="left",padx=(6,0))
        sr=tk.Frame(left,bg=PANEL_BG); sr.pack(fill="x",pady=(3,0))
        self.done_lbl =tk.Label(sr,text="Ready",font=FS,bg=PANEL_BG,fg=TEXT_S)
        self.done_lbl.pack(side="left")
        self.speed_lbl=tk.Label(sr,text="",font=FS,bg=PANEL_BG,fg=TEXT_S)
        self.speed_lbl.pack(side="left",padx=(10,0))
        self.time_lbl =tk.Label(sr,text="",font=FS,bg=PANEL_BG,fg=TEXT_S)
        self.time_lbl.pack(side="right")

        right=tk.Frame(parent,bg=PANEL_BG)
        right.pack(side="right",padx=16,pady=14)
        self.start_btn=AnimatedButton(right,
            text="▶  Start Batch Processing",font=FH,
            bg=ACCENT,fg=TEXT_P,
            hover_bg=ACCENT_L, press_bg=ACCENT_H,
            padx=24,pady=12,
            command=self._start)
        self.start_btn.pack()

    # ── Toggle switches ───────────────────────────────────────────────────────
    def _tog_crop(self):
        if self.crop_on.get(): self.c_frame.pack(fill="x")
        else:                  self.c_frame.pack_forget()

    def _tog_trim(self):
        try:
            if self.trim_on.get():
                self.trim_opts_frame.pack(fill="x", pady=(0,10))
                self.trim_info_frame.pack(fill="x", pady=(0,0))
            else:
                self.trim_opts_frame.pack_forget()
                self.trim_info_frame.pack_forget()
        except Exception:
            pass

    def _tog_size(self):
        try:
            if self.size_on.get():
                self.sz_f.pack(fill="x", padx=14, pady=(0,12))
            else:
                self.sz_f.pack_forget()
        except Exception:
            pass

    def _set_unit(self,u):
        self.size_unit.set(u)
        kb_bg = ACCENT if u=="KB" else BORDER_L
        mb_bg = ACCENT if u=="MB" else BORDER_L
        self.kb_btn.config(bg=kb_bg, fg=TEXT_P if u=="KB" else TEXT_S)
        self.kb_btn._normal_bg = kb_bg
        self.kb_btn._hover_bg  = ACCENT_L if u=="KB" else _lighten(BORDER_L,18)
        self.kb_btn._press_bg  = ACCENT_H if u=="KB" else BORDER
        self.mb_btn.config(bg=mb_bg, fg=TEXT_P if u=="MB" else TEXT_S)
        self.mb_btn._normal_bg = mb_bg
        self.mb_btn._hover_bg  = ACCENT_L if u=="MB" else _lighten(BORDER_L,18)
        self.mb_btn._press_bg  = ACCENT_H if u=="MB" else BORDER

    # ── Preview ───────────────────────────────────────────────────────────────
    def _draw_prev(self):
        self._on_shape_var_change()
        c=self.prev_c; c.delete("all")
        CW,CH=240,150
        c.create_rectangle(8,4,CW-8,CH-4,fill=PANEL_BG,outline=BORDER,width=1)
        rw=(CW-16)*(self.c_rx.get()/100)
        rh=(CH-8) *(self.c_ry.get()/100)
        cx=CW/2+(CW-16)*(self.c_ox.get()/100)
        cy=CH/2+(CH-8) *(self.c_oy.get()/100)
        shape = self.shape_var.get()
        kw = {}
        if shape=="star":
            kw["points"]=self.star_points.get()
            kw["inner_ratio"]=self.star_inner.get()/100
        if shape in ("ellipse",):
            c.create_oval(cx-rw/2,cy-rh/2,cx+rw/2,cy+rh/2,
                          fill=ACCENT,outline=TEXT_P,width=1)
        elif shape=="rectangle":
            c.create_rectangle(cx-rw/2,cy-rh/2,cx+rw/2,cy+rh/2,
                               fill=ACCENT,outline=TEXT_P,width=1)
        else:
            pts = shape_outline_points(shape,cx,cy,rw,rh,**kw)
            if pts:
                flat=[]
                for px,py in pts: flat+=[px,py]
                c.create_polygon(*flat, fill=ACCENT, outline=TEXT_P, width=1)
        c.create_line(cx-5,cy,cx+5,cy,fill=TEXT_P,width=1)
        c.create_line(cx,cy-5,cx,cy+5,fill=TEXT_P,width=1)

    def _apply_bg(self):
        h=self.c_bg_hex.get().strip()
        if not h.startswith("#"): h="#"+h
        self.c_bg_hex.set(h)
        try:
            r,g,b=int(h[1:3],16),int(h[3:5],16),int(h[5:7],16)
            self.c_bg_rgb=(r,g,b)
            self.bg_sw.config(bg=h)
        except Exception:
            messagebox.showerror("Color Error","Enter a valid hex code, e.g. #ffffff")

    # ── File management ───────────────────────────────────────────────────────
    def _add_files(self):
        fs=filedialog.askopenfilenames(
            title="Select Images",
            filetypes=[("Images","*.jpg *.jpeg *.png *.webp *.bmp *.tiff *.tif"),
                       ("All Files","*.*")])
        for f in fs:
            if f not in self.files: self.files.append(f)
        self._ref_list()

    def _add_folder(self):
        d=filedialog.askdirectory(title="Select Folder")
        if not d: return
        exts={".jpg",".jpeg",".png",".webp",".bmp",".tiff",".tif"}
        for rd,_,fs in os.walk(d):
            for f in fs:
                if Path(f).suffix.lower() in exts:
                    fp=os.path.join(rd,f)
                    if fp not in self.files: self.files.append(fp)
        self._ref_list()

    def _clear_files(self):
        self.files.clear(); self._ref_list()

    def _ref_list(self):
        self.file_lb.delete(0,"end")
        for f in self.files:
            tag=" ✏" if f in self.manual_map else ""
            self.file_lb.insert("end",f"  {Path(f).name}{tag}")
        n=len(self.files)
        self.fc_lbl.config(
            text=f"{n} image(s) selected",
            fg=SUCCESS if n>0 else TEXT_S)

    def _pick_out(self):
        d=filedialog.askdirectory(title="Output Folder")
        if d: self.output_dir.set(d)

    # ── Session launch ────────────────────────────────────────────────────────
    def _launch_session(self):
        in_dir  = self.ses_in.get().strip()
        out_dir = self.ses_out.get().strip()
        if not in_dir:
            messagebox.showwarning("No Folder","Please select an input folder.")
            return
        if not out_dir:
            messagebox.showwarning("No Output","Please select an output folder.")
            return
        exts={".jpg",".jpeg",".png",".webp",".bmp",".tiff",".tif"}
        files=[str(p) for p in sorted(Path(in_dir).iterdir())
               if p.suffix.lower() in exts]
        if not files:
            messagebox.showwarning("No Images","No images found in that folder.")
            return

        sfmt=self.ses_fmt.get()
        if sfmt=="Original": sfmt=None

        SessionEditor(
            self.root, files, out_dir,
            mode=self.ses_mode.get(),
            init_shape=self.ses_shape.get(),
            init_rx=self.ses_rx.get(),
            init_ry=self.ses_ry.get(),
            bg_color=self.c_bg_rgb,
            out_fmt=sfmt,
            lim_size=self.size_on.get(),
            target_kb=self._get_target_kb(),
            resize_mode=self.ses_resize_mode.get(),
            tw=self._get_dim(self.width_var,413),
            th=self._get_dim(self.height_var,531),
            trim_tol=self.trim_tol.get(),
            trim_pad=self.trim_pad.get(),
        )

    def _get_target_kb(self):
        if not self.size_on.get(): return None
        try:
            v=float(self.size_kb_var.get())
            if v<=0: return None
            return v*1024 if self.size_unit.get()=="MB" else v
        except (ValueError, TypeError): return None

    def _get_dim(self,var,default):
        try:
            v=int(var.get())
            return v if v>0 else default
        except (ValueError, TypeError):
            return default

    # ── Logging ───────────────────────────────────────────────────────────────────
    def _log(self,msg,tag=""):
        self.log_q.put((msg,tag))

    def _poll(self):
        while not self.log_q.empty():
            msg,tag=self.log_q.get_nowait()
            c={"ok":SUCCESS,"err":ERROR,"info":ACCENT}.get(tag,TEXT_S)
            self.done_lbl.config(text=msg[:72],fg=c)
        self.root.after(80,self._poll)

    # ── Start batch processing ───────────────────────────────────────────────────
    def _start(self):
        if self.processing: return
        if not self.files:
            messagebox.showwarning("No Images","Add at least one image."); return
        out=self.output_dir.get().strip()
        if not out:
            messagebox.showwarning("No Output","Select an output folder."); return
        tw=self._get_dim(self.width_var,413)
        th=self._get_dim(self.height_var,531)
        if self.mode_var.get()!="no_resize" and (tw<=0 or th<=0):
            messagebox.showerror("Invalid Size","Enter valid pixel values."); return
        target_kb=self._get_target_kb()
        out_fmt=self.out_fmt_var.get()
        if out_fmt=="Original": out_fmt=None

        Path(out).mkdir(parents=True,exist_ok=True)
        self.processing=True
        self.start_btn.config(state="disabled",text="⏳  Processing…")
        self.prog_var.set(0); self.pct_lbl.config(text="0%")
        self.speed_lbl.config(text=""); self.time_lbl.config(text="")

        threading.Thread(target=self._run,
            args=(self.files[:],out,tw,th,target_kb,out_fmt),
            daemon=True).start()

    def _run(self,files,out_dir,tw,th,target_kb,out_fmt):
        mode      =self.mode_var.get()
        n_threads =self.thread_var.get()
        keep      =self.keep_struct.get()
        base_dir  =str(Path(files[0]).parent) if keep and files else None

        do_shape  =self.crop_on.get()
        shape     =self.shape_var.get()
        crx,cry   =self.c_rx.get(),self.c_ry.get()
        cox,coy   =self.c_ox.get()/100, self.c_oy.get()/100
        cbg       =self.c_bg_rgb
        shape_kw  ={}
        if shape=="star":
            shape_kw["points"]=self.star_points.get()
            shape_kw["inner_ratio"]=self.star_inner.get()/100

        do_trim   =self.trim_on.get()
        trim_tol  =self.trim_tol.get()
        trim_pad  =self.trim_pad.get()

        lim       =self.size_on.get()
        cmap      = dict(self.manual_map)

        total=len(files)
        self._log(f"▶ {total} images | shape={('on('+shape+')') if do_shape else 'off'} "
                  f"| trim={'on' if do_trim else 'off'} | format={out_fmt or 'original'} "
                  f"| threads={n_threads}","info")

        tasks=[]
        for f in files:
            tasks.append((f,out_dir,tw,th,mode,keep,base_dir,
                          do_shape,shape,crx,cry,cox,coy,cbg,shape_kw,
                          do_trim,trim_tol,trim_pad,None,
                          lim,target_kb,out_fmt,cmap))

        done=errs=0; t0=time.perf_counter()
        results=[]   # جمع‌آوری نتایج برای گزارش نهایی
        with ThreadPoolExecutor(max_workers=n_threads) as ex:
            futs={ex.submit(process_image,t):t for t in tasks}
            for fut in as_completed(futs):
                ok,src,_,kb,msg=fut.result()
                done+=1
                name=Path(src).name
                results.append((ok,name,kb,msg))
                if ok: self._log(f"✓ {name}  ({kb:.0f} KB)","ok")
                else:
                    errs+=1
                    self._log(f"✗ {name}: {(msg or '').split(chr(10))[-1][:55]}","err")
                elapsed=time.perf_counter()-t0
                spd=done/elapsed if elapsed>0 else 0
                eta=(total-done)/spd if spd>0 else 0
                pct=done/total*100
                # Note: this code runs on a background thread; Tk widgets must never
                # be called directly from a non-main thread. All UI updates are
                # dispatched to the main thread via root.after().
                self.root.after(0, self._update_progress_ui, pct, spd, eta, done, total, elapsed)

        elapsed=time.perf_counter()-t0
        self._log(f"✔ {done-errs}/{total} succeeded · {errs} error(s) · {elapsed:.2f}s","info")
        self.root.after(0, self._finish_processing, results, elapsed, out_dir)

    def _update_progress_ui(self, pct, spd, eta, done, total, elapsed):
        """Updates the progress bar and labels; must only be called from the main Tk thread."""
        self.prog_var.set(pct)
        self.pct_lbl.config(text=f"{pct:.0f}%")
        self.speed_lbl.config(text=f"{spd:.1f} img/s")
        self.time_lbl.config(
            text=f"{eta:.0f}s remaining" if done<total
            else f"Done in {elapsed:.1f}s")

    def _finish_processing(self, results=None, elapsed=0.0, out_dir=None):
        """Resets the start button to normal state; must only be called from the main Tk thread."""
        self.start_btn.config(state="normal",text="▶  Start Batch Processing")
        self.processing=False
        if results is not None:
            self.root.after(150, self._show_report, results, elapsed, out_dir)

    def _show_report(self, results, elapsed, out_dir):
        """Shows a final report window with per-file results and a summary."""
        try:
            import tkinter as tk
            from tkinter import ttk
            ok_list = [r for r in results if r[0]]
            err_list = [r for r in results if not r[0]]

            win = tk.Toplevel(self.root)
            win.title("Batch Report")
            win.geometry("620x460")
            win.configure(bg="#0a2318")
            win.transient(self.root)

            # Header
            head = tk.Frame(win, bg="#0f2e1f")
            head.pack(fill="x", padx=10, pady=(10,6))
            total_kb = sum(r[2] or 0 for r in ok_list)
            summary = (f"✔ {len(ok_list)}/{len(results)} succeeded  ·  "
                       f"✗ {len(err_list)} error(s)  ·  {elapsed:.2f}s  ·  "
                       f"{total_kb/1024:.1f} MB output")
            tk.Label(head, text=summary, font=("Segoe UI", 11, "bold"),
                     bg="#0f2e1f", fg="#6ee7b7").pack(anchor="w")

            # Listbox with scrollbar
            frame = tk.Frame(win, bg="#0a2318")
            frame.pack(fill="both", expand=True, padx=10, pady=(4,6))
            lb = tk.Listbox(frame, bg="#0a2318", fg="#ecfdf5",
                            selectbackground="#10b981", font=("Consolas", 9))
            sb = ttk.Scrollbar(frame, orient="vertical", command=lb.yview)
            lb.configure(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")
            lb.pack(side="left", fill="both", expand=True)

            for ok, name, kb, msg in results:
                if ok:
                    lb.insert("end", f"✔ {name}  ({kb/1024:.2f} MB)")
                else:
                    err_txt = (msg or "").strip().split(chr(10))[-1][:60]
                    lb.insert("end", f"✗ {name}: {err_txt}")

            # Buttons
            btns = tk.Frame(win, bg="#0a2318")
            btns.pack(fill="x", padx=10, pady=(0,10))
            tk.Button(btns, text="Close", bg="#166534", fg="#ecfdf5", relief="flat",
                      command=win.destroy).pack(side="left", padx=3)
            if out_dir:
                tk.Button(btns, text="Open Output Folder", bg="#10b981", fg="#061a10",
                          relief="flat", command=lambda: self._open_folder(out_dir)
                          ).pack(side="left", padx=3)
        except Exception:
            pass

    def _open_folder(self, folder):
        """Opens a folder in the OS file explorer."""
        try:
            import os, sys
            if os.name == "nt":
                os.startfile(str(folder))  # type: ignore
            elif sys.platform == "darwin":
                import subprocess as sp
                sp.Popen(["open", str(folder)])
            else:
                import subprocess as sp
                sp.Popen(["xdg-open", str(folder)])
        except Exception:
            pass


    # ── Tab: Voice ───────────────────────────────────────────────────────────
    def _build_voice(self, p):
        wrap = tk.Frame(p, bg=DARK_BG)
        wrap.pack(fill="both", expand=True, padx=20, pady=14)

        # Notice for missing libraries
        missing = []
        if not SR_AVAILABLE:
            missing.append("SpeechRecognition + PyAudio   →   pip install SpeechRecognition pyaudio")
        if not OPENPYXL_AVAILABLE:
            missing.append("openpyxl   →   pip install openpyxl")
        if not DOCX_AVAILABLE:
            missing.append("python-docx   →   pip install python-docx")
        if not PPTX_AVAILABLE:
            missing.append("python-pptx   →   pip install python-pptx")
        if not VOSK_AVAILABLE:
            missing.append("vosk (for offline mode)   →   pip install vosk")
        if missing:
            warn = self._card(wrap, "Required Libraries Not Installed", "⚠")
            for m in missing:
                tk.Label(warn, text="•  "+m, font=FM, bg=CARD_BG, fg=WARNING,
                         justify="left").pack(anchor="w")
            tk.Label(warn, text="After installing the packages above, restart the program.",
                     font=FS, bg=CARD_BG, fg=TEXT_S).pack(anchor="w", pady=(6,0))

        # Intro
        intro = self._card(wrap, "Speech to Text  ←  Word / Excel / PowerPoint", "🎙")
        tk.Label(intro,
                 text="Speak into the microphone; the text is typed live.\n"
                      "Afterwards you can edit the text and export it to a Word, Excel, or PowerPoint file.",
                 font=FS, bg=CARD_BG, fg=TEXT_S, justify="left").pack(anchor="w")

        # Engine (online / offline)
        ec2 = self._card(wrap, "Speech Recognition Mode", "🔌")
        erow = tk.Frame(ec2, bg=CARD_BG); erow.pack(anchor="w")
        tk.Radiobutton(erow, text="Online (Google) — requires internet", font=FB,
                       bg=CARD_BG, fg=TEXT_P, selectcolor=DARK_BG,
                       activebackground=CARD_BG, activeforeground=ACCENT_L,
                       variable=self.voice_engine, value="online",
                       command=self._voice_toggle_engine_opts).pack(side="left", padx=(0,16))
        tk.Radiobutton(erow, text="Offline (Vosk) — no internet needed", font=FB,
                       bg=CARD_BG, fg=TEXT_P, selectcolor=DARK_BG,
                       activebackground=CARD_BG, activeforeground=ACCENT_L,
                       variable=self.voice_engine, value="offline",
                       command=self._voice_toggle_engine_opts).pack(side="left")

        self.voice_offline_opts = tk.Frame(ec2, bg=CARD_BG)
        tk.Label(self.voice_offline_opts,
                 text="Select the Vosk model folder (download the language model you want from\n"
                      "alphacephei.com/vosk/models and unzip it):",
                 font=FS, bg=CARD_BG, fg=TEXT_S, justify="left").pack(anchor="w", pady=(8,4))
        mrow = tk.Frame(self.voice_offline_opts, bg=CARD_BG); mrow.pack(anchor="w", fill="x")
        _styled_entry(mrow, textvariable=self.voice_model_path, ipady=5)
        AnimatedButton(mrow, text="Select Folder…", font=FB, bg=CARD_BG, fg=TEXT_P,
                  hover_bg=_lighten(CARD_BG,20), press_bg=_darken(CARD_BG,10),
                  padx=10, pady=4,
                  command=self._voice_pick_model_dir).pack(side="left", padx=(6,0))
        self._voice_toggle_engine_opts()

        # Language
        lc = self._card(wrap, "Recognition Language", "🌐")
        row = tk.Frame(lc, bg=CARD_BG); row.pack(anchor="w")
        tk.Label(row, text="Language:", font=FB, bg=CARD_BG, fg=TEXT_S).pack(side="left", padx=(0,8))
        self._voice_lang_name = tk.StringVar(value=VOICE_LANGS[0][0])
        cb = ttk.Combobox(row, textvariable=self._voice_lang_name,
                           values=[n for n,_ in VOICE_LANGS], width=20, state="readonly")
        cb.pack(side="left")
        def _on_lang(*_):
            for n,code in VOICE_LANGS:
                if n == self._voice_lang_name.get():
                    self.voice_lang.set(code); break
        cb.bind("<<ComboboxSelected>>", _on_lang)
        tk.Label(lc, text="For better accuracy, speak short and clearly and pause between sentences.",
                 font=FS, bg=CARD_BG, fg=TEXT_S).pack(anchor="w", pady=(6,0))
        tk.Checkbutton(lc, text="Convert spoken numbers to digits  (e.g. \"twenty three\" → 23)",
                       font=FB, bg=CARD_BG, fg=TEXT_P, selectcolor=DARK_BG,
                       activebackground=CARD_BG, activeforeground=ACCENT_L,
                       variable=self.voice_numbers_to_digits).pack(anchor="w", pady=(6,0))

        # Controls
        cc = self._card(wrap, "Recording & Dictation", "🎚")
        btnrow = tk.Frame(cc, bg=CARD_BG); btnrow.pack(anchor="w", pady=(0,8))
        self.start_voice_btn = AnimatedButton(
            btnrow, text="🎙  Start Recording", font=FH,
            bg=ACCENT, fg=TEXT_P,
            hover_bg=ACCENT_L, press_bg=ACCENT_H,
            padx=16, pady=8, command=self._voice_start)
        self.start_voice_btn.pack(side="left", padx=(0,8))

        self.stop_voice_btn = AnimatedButton(
            btnrow, text="⏹  Stop", font=FH,
            bg=ERROR, fg=TEXT_P,
            hover_bg=_lighten(ERROR,15), press_bg=_darken(ERROR,15),
            padx=16, pady=8, state="disabled", command=self._voice_stop)
        self.stop_voice_btn.pack(side="left", padx=(0,8))

        AnimatedButton(btnrow, text="↵  New Line", font=FB, bg=CARD_BG, fg=TEXT_S,
                  hover_bg=_lighten(CARD_BG,20), press_bg=_darken(CARD_BG,10),
                  padx=12, pady=8,
                  command=self._voice_next_line).pack(side="left", padx=(0,8))

        AnimatedButton(btnrow, text="⌫  Delete Last Line", font=FB, bg=CARD_BG, fg=TEXT_S,
                  hover_bg=_lighten(CARD_BG,20), press_bg=_darken(CARD_BG,10),
                  padx=12, pady=8,
                  command=self._voice_undo_last).pack(side="left", padx=(0,8))

        AnimatedButton(btnrow, text="🗑  Clear All", font=FB, bg=CARD_BG, fg=ERROR,
                  hover_bg=_mix(CARD_BG, ERROR, 0.15), press_bg=_mix(CARD_BG, ERROR, 0.25),
                  padx=12, pady=8,
                  command=self._voice_clear).pack(side="left")

        self.voice_status_lbl = tk.Label(cc, text="◼  Ready", font=FS, bg=CARD_BG, fg=TEXT_S)
        self.voice_status_lbl.pack(anchor="w")

        # Text preview
        pc = self._card(wrap, "Text Preview", "👁")
        tk.Label(pc,
                 text="Each line in this box is treated as a separate \"row\" when exporting.",
                 font=FS, bg=CARD_BG, fg=TEXT_S, justify="left").pack(anchor="w", pady=(0,6))
        tf = tk.Frame(pc, bg=PANEL_BG, highlightbackground=BORDER_L, highlightthickness=1)
        tf.pack(fill="x")
        sb = tk.Scrollbar(tf, bg=PANEL_BG, troughcolor=PANEL_BG,
                          activebackground=BORDER_L, width=10)
        sb.pack(side="right", fill="y")
        self.voice_text = tk.Text(tf, height=10, bg=PANEL_BG, fg=TEXT_P,
                                   insertbackground=ACCENT, font=(_PFONT,11),
                                   bd=0, wrap="word", yscrollcommand=sb.set, undo=True)
        self.voice_text.pack(side="left", fill="both", expand=True)
        sb.config(command=self.voice_text.yview)

        # Output
        ec = self._card(wrap, "Export", "📤")
        fr = tk.Frame(ec, bg=CARD_BG); fr.pack(anchor="w")
        tk.Label(fr, text="Output format:", font=FB, bg=CARD_BG, fg=TEXT_S).pack(side="left", padx=(0,10))
        for val, lbl in [("xlsx","📊  Excel"), ("docx","📄  Word"), ("pptx","📽  PowerPoint")]:
            tk.Radiobutton(fr, text=lbl, variable=self.voice_export_fmt, value=val,
                           bg=CARD_BG, fg=TEXT_S, selectcolor=DARK_BG,
                           activebackground=CARD_BG, activeforeground=ACCENT_L,
                           font=FB, bd=0, cursor="hand2",
                           command=self._voice_toggle_export_opts).pack(side="left", padx=(0,12))

        # Excel layout
        self.voice_xlsx_opts = tk.Frame(ec, bg=CARD_BG)
        tk.Label(self.voice_xlsx_opts, text="Excel layout:", font=FS,
                 bg=CARD_BG, fg=TEXT_S).pack(anchor="w", pady=(8,2))
        tk.Radiobutton(self.voice_xlsx_opts,
                       text="Each line in a separate column  (rows side by side in one row)",
                       variable=self.voice_excel_layout, value="column",
                       bg=CARD_BG, fg=TEXT_S, selectcolor=DARK_BG,
                       activebackground=CARD_BG, activeforeground=ACCENT_L,
                       font=FB, bd=0, cursor="hand2").pack(anchor="w")
        tk.Radiobutton(self.voice_xlsx_opts,
                       text="Each line in a separate row  (rows stacked in one column)",
                       variable=self.voice_excel_layout, value="row",
                       bg=CARD_BG, fg=TEXT_S, selectcolor=DARK_BG,
                       activebackground=CARD_BG, activeforeground=ACCENT_L,
                       font=FB, bd=0, cursor="hand2").pack(anchor="w")

        # PowerPoint layout
        self.voice_pptx_opts = tk.Frame(ec, bg=CARD_BG)
        tk.Label(self.voice_pptx_opts, text="PowerPoint layout:", font=FS,
                 bg=CARD_BG, fg=TEXT_S).pack(anchor="w", pady=(8,2))
        tk.Radiobutton(self.voice_pptx_opts,
                       text="All rows as bullets on one slide",
                       variable=self.voice_pptx_layout, value="bullets",
                       bg=CARD_BG, fg=TEXT_S, selectcolor=DARK_BG,
                       activebackground=CARD_BG, activeforeground=ACCENT_L,
                       font=FB, bd=0, cursor="hand2").pack(anchor="w")
        tk.Radiobutton(self.voice_pptx_opts,
                       text="Each row = a separate slide (slide title)",
                       variable=self.voice_pptx_layout, value="slides",
                       bg=CARD_BG, fg=TEXT_S, selectcolor=DARK_BG,
                       activebackground=CARD_BG, activeforeground=ACCENT_L,
                       font=FB, bd=0, cursor="hand2").pack(anchor="w")

        # Output path
        pr2 = tk.Frame(ec, bg=CARD_BG); pr2.pack(fill="x", pady=(10,0))
        tk.Label(pr2, text="Output file path:", font=FB, bg=CARD_BG, fg=TEXT_S,
                 width=14, anchor="w").pack(side="left")
        _styled_entry(pr2, textvariable=self.voice_out_path, ipady=5)
        AnimatedButton(pr2, text="Browse…", font=FS, bg=BORDER_L, fg=TEXT_S,
                  hover_bg=ACCENT, press_bg=ACCENT_H,
                  padx=10, pady=5,
                  command=self._voice_pick_output).pack(side="left", padx=(6,0))

        AnimatedButton(ec, text="💾  Save Output", font=(_PFONT,13,"bold"),
                  bg=ACCENT, fg=TEXT_P,
                  hover_bg=ACCENT_L, press_bg=ACCENT_H,
                  padx=24, pady=12,
                  command=self._voice_export).pack(anchor="w", pady=(12,0))

        self._voice_toggle_export_opts()

    def _voice_toggle_export_opts(self):
        fmt = self.voice_export_fmt.get()
        self.voice_xlsx_opts.pack(fill="x") if fmt=="xlsx" else self.voice_xlsx_opts.pack_forget()
        self.voice_pptx_opts.pack(fill="x") if fmt=="pptx" else self.voice_pptx_opts.pack_forget()

    # ── Voice: engine helpers ────────────────────────────────────────────────
    def _voice_toggle_engine_opts(self):
        if self.voice_engine.get() == "offline":
            self.voice_offline_opts.pack(fill="x")
        else:
            self.voice_offline_opts.pack_forget()

    def _voice_pick_model_dir(self):
        d = filedialog.askdirectory(title="Select Vosk Model Folder")
        if d:
            self.voice_model_path.set(d)

    # ── Voice: status helper ─────────────────────────────────────────────────
    def _voice_set_status(self, text, color=TEXT_S):
        self.voice_status_lbl.config(text=text, fg=color)


    # ── Voice: start / stop listening ───────────────────────────────────────
    def _voice_start(self):
        if not SR_AVAILABLE:
            messagebox.showerror(
                "Library Not Available",
                "To use speech recognition, run the following command in the terminal:\n\n"
                "pip install SpeechRecognition pyaudio")
            return
        if self.voice_listening:
            return

        engine = self.voice_engine.get()

        offline_model = None
        if engine == "offline":
            if not VOSK_AVAILABLE:
                messagebox.showerror(
                    "Library Not Available",
                    "To use offline mode, run the following command in the terminal:\n\n"
                    "pip install vosk")
                return
            model_path = self.voice_model_path.get().strip()
            if not model_path or not os.path.isdir(model_path):
                messagebox.showwarning(
                    "No Model Selected",
                    "First select the Vosk model folder "
                    "(download one from alphacephei.com/vosk/models).")
                return
            try:
                if self.vosk_model is None or self.vosk_model_loaded_path != model_path:
                    self._voice_set_status("◌  Loading offline model…", ACCENT)
                    self.root.update_idletasks()
                    self.vosk_model = vosk.Model(model_path)
                    self.vosk_model_loaded_path = model_path
                offline_model = self.vosk_model
            except Exception as e:
                messagebox.showerror("Offline Model Error", f"Failed to load model:\n{e}")
                self._voice_set_status("◼  Offline model error", ERROR)
                return

        try:
            if self.voice_recognizer is None:
                self.voice_recognizer = sr.Recognizer()
                self.voice_recognizer.pause_threshold = 0.6
                self.voice_recognizer.non_speaking_duration = 0.3
            if self.voice_mic is None:
                self.voice_mic = sr.Microphone()
            self._voice_set_status("◌  Calibrating microphone…", ACCENT)
            with self.voice_mic as source:
                self.voice_recognizer.adjust_for_ambient_noise(source, duration=0.6)
        except Exception as e:
            messagebox.showerror("Microphone Error",
                                  f"Could not access the microphone:\n{e}")
            self._voice_set_status("◼  Microphone error", ERROR)
            return

        lang_code = self.voice_lang.get()

        if engine == "offline":
            def callback(recognizer, audio):
                try:
                    rec = vosk.KaldiRecognizer(offline_model, audio.sample_rate)
                    rec.AcceptWaveform(audio.get_raw_data(convert_rate=audio.sample_rate,
                                                            convert_width=2))
                    result = _json.loads(rec.FinalResult())
                    text = (result.get("text") or "").strip()
                    if text:
                        self.voice_q.put(("text", text))
                except Exception as e:
                    self.voice_q.put(("error", f"⚠ Error (offline): {e}"))
        else:
            def callback(recognizer, audio):
                try:
                    text = recognizer.recognize_google(audio, language=lang_code)
                    if text:
                        self.voice_q.put(("text", text))
                except sr.UnknownValueError:
                    pass  # silence / unrecognizable, simply ignored
                except sr.RequestError as e:
                    self.voice_q.put(("error", f"⚠ Error connecting to speech recognition service: {e}"))
                except Exception as e:
                    self.voice_q.put(("error", f"⚠ Error: {e}"))

        try:
            self.voice_stop_listening = self.voice_recognizer.listen_in_background(
                self.voice_mic, callback, phrase_time_limit=12)
        except Exception as e:
            messagebox.showerror("Error", f"Could not start recording:\n{e}")
            return

        self.voice_listening = True
        self.start_voice_btn.config(state="disabled")
        self.stop_voice_btn.config(state="normal")
        self._voice_set_status("●  Listening… speak now", SUCCESS)

    def _voice_stop(self):
        if self.voice_stop_listening:
            try:
                self.voice_stop_listening(wait_for_stop=False)
            except Exception:
                pass
            self.voice_stop_listening = None
        self.voice_listening = False
        self.start_voice_btn.config(state="normal")
        self.stop_voice_btn.config(state="disabled")
        self._voice_set_status("◼  Stopped", TEXT_S)

    # ── Voice: check recognized text in the background ──────────────────────────────
    def _voice_poll(self):
        while not self.voice_q.empty():
            kind, val = self.voice_q.get_nowait()
            if kind == "text":
                self._voice_append_text(val)
            elif kind == "error":
                self._voice_set_status(val, ERROR)
        self.root.after(150, self._voice_poll)

    def _voice_append_text(self, text):
        if self.voice_numbers_to_digits.get():
            text = _numbers_to_digits(text)
        txt = self.voice_text
        current = txt.get("1.0", "end-1c")
        if current and not current.endswith(("\n", " ")):
            txt.insert("end", " ")
        txt.insert("end", text)
        txt.see("end")
        preview = text if len(text) <= 40 else text[:40] + "…"
        self._voice_set_status(f"✓  \"{preview}\"", SUCCESS)

    # ── Voice: line control ─────────────────────────────────────────────────
    def _voice_next_line(self):
        txt = self.voice_text
        current = txt.get("1.0", "end-1c")
        if current and not current.endswith("\n"):
            txt.insert("end", "\n")
        txt.see("end")
        txt.focus_set()

    def _voice_undo_last(self):
        txt = self.voice_text
        current = txt.get("1.0", "end-1c").rstrip("\n")
        if not current:
            return
        lines = current.split("\n")
        lines = lines[:-1]
        new_text = "\n".join(lines)
        if lines:
            new_text += "\n"
        txt.delete("1.0", "end")
        txt.insert("1.0", new_text)
        txt.see("end")

    def _voice_clear(self):
        if not self.voice_text.get("1.0", "end-1c").strip():
            return
        if messagebox.askyesno("Clear", "Clear all dictated text?"):
            self.voice_text.delete("1.0", "end")
            self._voice_set_status("◼  Ready", TEXT_S)

    # ── Voice: export ─────────────────────────────────────────────────────────
    def _voice_get_lines(self):
        raw = self.voice_text.get("1.0", "end-1c")
        lines = [l.strip() for l in raw.split("\n")]
        return [l for l in lines if l]

    def _voice_pick_output(self):
        fmt = self.voice_export_fmt.get()
        ext   = {"xlsx":".xlsx", "docx":".docx", "pptx":".pptx"}[fmt]
        types = {"xlsx":[("Excel Workbook","*.xlsx")],
                 "docx":[("Word Document","*.docx")],
                 "pptx":[("PowerPoint Presentation","*.pptx")]}[fmt]
        f = filedialog.asksaveasfilename(
            title="Save Output", defaultextension=ext,
            filetypes=types + [("All files","*.*")])
        if f:
            self.voice_out_path.set(f)

    def _voice_export(self):
        lines = self._voice_get_lines()
        if not lines:
            messagebox.showwarning("No Text",
                                    "First speak into the microphone or write some text in the box.")
            return

        out = self.voice_out_path.get().strip()
        fmt = self.voice_export_fmt.get()
        expected_ext = {"xlsx":".xlsx","docx":".docx","pptx":".pptx"}[fmt]
        if not out:
            self._voice_pick_output()
            out = self.voice_out_path.get().strip()
            if not out:
                return
        if Path(out).suffix.lower() != expected_ext:
            out = str(Path(out).with_suffix(expected_ext))
            self.voice_out_path.set(out)

        try:
            if fmt == "xlsx":
                self._voice_export_xlsx(lines, out)
            elif fmt == "docx":
                self._voice_export_docx(lines, out)
            else:
                self._voice_export_pptx(lines, out)
        except Exception as e:
            messagebox.showerror("Error Saving File", str(e))
            self._voice_set_status(f"✗  Error: {e}", ERROR)
            return

        self._voice_set_status(f"✔  Saved: {Path(out).name}", SUCCESS)
        messagebox.showinfo("Done", f"File saved successfully:\n{out}")

    def _voice_export_xlsx(self, lines, out):
        if not OPENPYXL_AVAILABLE:
            raise RuntimeError("openpyxl library is not installed.\npip install openpyxl")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Voice"
        ws.sheet_view.rightToLeft = True

        if self.voice_excel_layout.get() == "row":
            # Each line in a separate row (column A)
            for i, line in enumerate(lines, start=1):
                ws.cell(row=i, column=1, value=line)
            ws.column_dimensions["A"].width = 60
        else:
            # Each line in a separate column (row 1)
            for i, line in enumerate(lines, start=1):
                ws.cell(row=1, column=i, value=line)
                ws.column_dimensions[get_column_letter(i)].width = max(12, min(40, len(line)+2))

        wb.save(out)

    def _voice_export_docx(self, lines, out):
        if not DOCX_AVAILABLE:
            raise RuntimeError("python-docx library is not installed.\npip install python-docx")
        doc = docx.Document()
        for line in lines:
            doc.add_paragraph(line)
        doc.save(out)

    def _voice_export_pptx(self, lines, out):
        if not PPTX_AVAILABLE:
            raise RuntimeError("python-pptx library is not installed.\npip install python-pptx")
        prs = Presentation()
        title_layout   = prs.slide_layouts[1]  # title and content

        if self.voice_pptx_layout.get() == "slides":
            for line in lines:
                slide = prs.slides.add_slide(title_layout)
                slide.shapes.title.text = line
                # Remove the content placeholder's placeholder text
                if len(slide.placeholders) > 1:
                    slide.placeholders[1].text = ""
        else:
            slide = prs.slides.add_slide(title_layout)
            slide.shapes.title.text = "Dictated Text"
            body = slide.placeholders[1]
            tf = body.text_frame
            tf.text = lines[0]
            for line in lines[1:]:
                para = tf.add_paragraph()
                para.text = line

        prs.save(out)


def main():
    root=tk.Tk()
    App(root)
    root.mainloop()

if __name__=="__main__":
    main()
