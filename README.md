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
<a href="https://github.com/mohammadazri/pdf-automation-toolkit/blob/main/LICENSE">
   <img src="https://img.shields.io/github/license/mohammadazri/pdf-automation-toolkit?color=brightgreen" alt="License"/>
</a>
</div>


---


## ✨ What can you do with this?

- Make certificates for a whole class or event in just a few minutes.
- Clean up your name list so it looks neat and works with the script.
- No more filename errors—just run and get your PDFs!
- Change the code if you want to use your own template or add more info.

---


## ⚡ How to get started

1. Download or clone this repo:
   ```sh
   git clone https://github.com/yourusername/pdf-automation-toolkit.git
   cd pdf-automation-toolkit/pdf_automation
   ```
2. Install the Python packages you need:
   ```sh
   pip install -r requirements.txt
   ```

3. Add your font files:
   - Place any `.ttf` or `.otf` font files you want to use in the new `fonts` folder inside the project directory.
   - The scripts will automatically detect all fonts in this folder and let you choose which one to use for certificate names.

4. Run the scripts:
   ```sh
   python clean_participant_names.py
   python auto_certificate_name_filler.py
   python auto_cert_with_manual verification.py
   ```

When running the certificate scripts, you'll be prompted to select your desired font from a dropdown menu. No code changes are needed to use new fonts—just add them to the `fonts` folder!

That's it! Your certificates will be saved in the same folder. You can change the scripts to use your own PDF design or adjust how the names are cleaned.

---



<!-- If you want, you can add a screenshot of your certificates or a PDF icon here. For now, no demo is included. -->

---


## 🧩 Can I change stuff?

<details>
   <summary>How to customize</summary>
   <ul>
      <li>Edit the Python files to use your own certificate template or add more info (like event name, date, etc).</li>
      <li>Change the name cleaning logic if your list is different.</li>
      <li>Use these scripts as a starting point for your own automation project.</li>
   </ul>
</details>

---


## 🤝 Want to help?

If you have ideas or improvements, feel free to open a pull request or issue. I'm still learning, so any help or feedback is appreciated!

---


## 📄 License

MIT License. You can use, modify, or share this project as you like.

---


<div align="center">
   <b>Made by a student for students and anyone who needs to automate certificates.<br>
   Fork it, change it, and make your workflow easier!</b>
</div>
## 🤝 Contributing
