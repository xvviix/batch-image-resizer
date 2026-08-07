<div align="center">

<br/>

# 🖼️ Batch Image Resizer v11

### Powerful batch image processing — English & Persian

<br/>

> A complete desktop image tool: batch resize, shape crop, edge trim,
> background removal, a visual session editor, and voice dictation —
> all in one Tkinter app. Available in **English** and **Persian**.

<br/>

<img src="https://img.shields.io/badge/Python-3.8%2B-3776AB?style=flat-square&logo=python&logoColor=white" alt="Python"/>
<img src="https://img.shields.io/badge/UI-Tkinter-10B981?style=flat-square" alt="UI"/>
<img src="https://img.shields.io/badge/OpenCV-5C3EE8?style=flat-square&logo=opencv&logoColor=white" alt="OpenCV"/>
<img src="https://img.shields.io/badge/Windows-✓-06B6D4?style=flat-square&logo=windows&logoColor=white" alt="Windows"/>
<img src="https://img.shields.io/badge/Linux-✓-06B6D4?style=flat-square&logo=linux&logoColor=white" alt="Linux"/>
<img src="https://img.shields.io/badge/License-MIT-green?style=flat-square" alt="MIT"/>

</div>

---

## ✨ Features

| Tab | Description |
|-----|-------------|
| 🖼️ **Images** | Batch processing with multi-threading, progress bar, speed & ETA |
| ↔️ **Resize** | Fit / Fill / Stretch modes with custom dimensions |
| ✂️ **Shape Crop** | Crop into star, heart, hexagon, triangle, ellipse + custom shapes |
| ▭ **Trim Edges** | Auto-remove surrounding background (flexible bounding box) |
| 🎨 **Editor** | Visual session editor with live preview, per-image overrides |
| 🎙️ **Voice Dictation** | Speech-to-text + export to Excel / Word / PowerPoint |

### Additional features

- **Compress to KB** — target a specific output size
- **Background removal** — with tolerance & padding controls
- **Batch rename structure** — keep folder structure or flatten
- **Persian number-to-digit conversion** — understands "صد و بیست و سه" → 123
- **Final report** — per-file results, error details, output folder shortcut

---

## 🚀 Quick Start

### Windows — just double-click `run.bat`

This script automatically:
1. Checks that Python is installed
2. Installs required libraries (first run only)
3. Lets you choose **English** or **Persian** version

### Linux / macOS

```bash
pip install -r requirements.txt
python batch_image_resizer_v11_en.py   # English
python batch_image_resizer_v11_fa.py   # Persian
```

---

## 🖥️ Prerequisites

| Requirement | Description |
|-------------|-------------|
| **Python** | 3.8 to 3.12 — **3.12 recommended** |
| **Internet** | Only on first run (to install libraries) |

> ⚠️ Python 3.13+ may cause issues with some libraries — 3.12 recommended.

---

## 📁 Project Structure

```
batch-image-resizer/
├── run.bat                       ← Windows launcher (double-click)
├── batch_image_resizer_v11_en.py ← English version
├── batch_image_resizer_v11_fa.py ← Persian version
├── requirements.txt              ← Dependencies
├── README.md                     ← This file
└── LICENSE                       ← MIT license
```

---

## 🛠️ Troubleshooting

| Problem | Solution |
|---------|----------|
| Program doesn't open | Run `run.bat` again to see the error message |
| Voice tab is disabled | Install optional libs: `pip install SpeechRecognition vosk` |
| Export buttons disabled | Install: `pip install openpyxl python-docx python-pptx` |
| Python 3.13 installed | Install Python 3.12 for best compatibility |

---

## 📄 License

[MIT](LICENSE) — free for personal and commercial use.

---

<div align="center">

**Built with ❤️ — batch processing, zero drama**

</div>
