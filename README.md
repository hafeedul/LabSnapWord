# LabSnap Word v2.0 — University Lab Screenshot Cropper & Word Report Builder

A native **Windows 10/11 Desktop Application** built with C# (.NET 8 / WPF). Capture screenshots using **`Ctrl + Shift + S`**, crop them on-screen or interactively inside the app, add a **Task / Heading Title** (e.g., `Task 1: Output`), and insert them directly into your Microsoft Word (`.docx`) report — **even while the Word file is open in Microsoft Word!**

---

## Key Features in v2.0

1. **`Ctrl + Shift + S` Global Screenshot Shortcut**
   - Press **`Ctrl + Shift + S`** anytime (even while working inside another program with `LabSnap Word` in the background) to immediately freeze the screen and drag-snip a region into the app.
2. **Works Even While the Word File Is Open in Microsoft Word**
   - If your `.docx` report is closed, `LabSnap Word` writes directly to the file using OpenXML.
   - If your `.docx` report is **open in Microsoft Word**, `LabSnap Word` automatically connects to your open Word window via Live COM Automation, appends the heading, notes, and cropped screenshot right at the bottom of your open document, and saves it automatically!
3. **Standalone Windows Installer (`LabSnapWord-Setup.exe`) & Portable App (`LabSnapWord.exe`)**
   - **Zero Downloads Required:** All runtime components are built directly into the executable.
   - Run **`LabSnapWord-Setup.exe`** (or click **"Install App to PC"** inside the app) to install **LabSnap Word** onto your PC with a Desktop Shortcut and Start Menu icon.
4. **Modern Step-by-Step UI & Interactive Crop Studio**
   - Guided **Step 1 → Step 2 → Step 3** sidebar with a live **Session History Log**.
   - Interactive cropper with darkened outside overlay masks so your selected crop area stands out clearly.
   - Automatic task numbering (`Task 1:` → `Task 2:` → `Task 3:`).
