<div align="center">
  <img src="https://img.shields.io/badge/PDF%20Automation-Toolkit-blueviolet?style=for-the-badge" alt="PDF Automation Toolkit" />
   <h1>PDF Automation Toolkit</h1>
   <p><b>This script lets you use your own PDF certificate design and automatically fills in hundreds of participant names—perfect for events, classes, or workshops!</b></p>
   <!-- You can add a relevant PDF automation image here if you want -->
<br>
<a href="https://github.com/mohammadazri/pdf-automation-toolkit">
   <img src="https://img.shields.io/github/stars/mohammadazri/pdf-automation-toolkit?style=social" alt="GitHub stars"/>
</a>
<a href="https://github.com/mohammadazri/pdf-automation-toolkit">
   <img src="https://img.shields.io/github/forks/mohammadazri/pdf-automation-toolkit?style=social" alt="GitHub forks"/>
</a>
<a href="https://github.com/mohammadazri/pdf-automation-toolkit/blob/master/LICENSE">
   <img src="https://img.shields.io/github/license/mohammadazri/pdf-automation-toolkit?color=brightgreen" alt="License"/>
</a>
</div>


---

## 🚀 Why use this?
- Automates certificates for large lists (hundreds) in minutes.
- Auto mode for one-click exports; Manual mode for per-name tweaks.
- Drag a rectangle to set the name spot; fonts auto-load from `fonts/`.
- Output folder chooser with safe filenames (no OS errors).

---

## ⚡ Quick start
```sh
git clone https://github.com/mohammadazri/pdf-automation-toolkit.git
cd pdf-automation-toolkit/pdf_automation
pip install -r requirements.txt
```
Optional: drop `.ttf/.otf` fonts into `fonts/` (subfolders ok).

---

## 🖥️ How to run
1) Clean names (optional)
```sh
python clean_participant_names.py
```

2) Auto mode (fast)
```sh
python auto_certificate_name_filler.py
```
Pick template → names TXT → font → draw name box → choose output folder → Generate.

3) Manual mode (fine-tune each name)
```sh
python "auto_cert_with_manual verification.py"
```
Same steps, plus adjust size/offset per name before saving.

---

## 📁 Output
- Default: `output/` (auto-created).
- Filenames are sanitized like `certificate_Name.pdf`.

---

## 🛠️ Customize
- Use your own PDF template and fonts.
- Tweak name cleaning or add extra text/QRs if needed.

---

## 🤝 Contributing & License
Suggestions welcome. MIT License.

<div align="center">
  <b>Built by a student to save organizers time.</b>
</div>
