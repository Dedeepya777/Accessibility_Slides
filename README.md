# Universal PowerPoint Accessibility Remediation Engine

An automated, non-destructive, and visually-lossless accessibility remediation engine for Microsoft PowerPoint (`.pptx`) presentations.

This tool automatically detects and fixes violations identified by Microsoft PowerPoint's native **Accessibility Checker**, ensuring compliance with **WCAG 2.1 AA**, **Section 508**, and **PDF/UA** standards while keeping the visual layout and slide designs 100% intact.

---

## 🌟 Key Features

- **Dynamic Image Classification & Alt Text**: Differentiates decorative elements (borders, icons, backgrounds) from informative images (charts, diagrams, photos) and automatically generates contextual Alt Text or marks items as decorative.
- **Visually-Lossless Slide Titles**: Promotes visual headings into official PowerPoint Title placeholders (preserving font, color, size, and position) or injects an off-canvas title so screen readers can navigate slides seamlessly without altering visual design.
- **Duplicate Title Disambiguation**: Adds clear enumerators (e.g., `(1 of 2)`) to repeated slide titles to pass screen-reader navigation checks.
- **Layer-Safe Reading Order**: Sequences non-overlapping elements in natural order (top-to-bottom, left-to-right) while using bounding-box collision detection to protect layered graphics, cards, and backgrounds.
- **Accessible Data Tables**: Enforces header row flags (`FirstRow = True`) and applies structural description summaries for screen readers.
- **Meaningful Hyperlinks**: Converts bare URLs into readable anchor text based on domain and path context, preserving font styles and hyperlinks.
- **Tagged Accessible PDF Export**: Automatically produces a tagged, accessible PDF (`.pdf`) with Document Structure Tags enabled.
- **Live Pre & Post Audit Verification**: Launches PowerPoint's native Accessibility Assistant to record violation counts before and after remediation, printing a compliance report in your console.

---

## 💻 System Requirements

Because this script uses Windows COM Automation and PowerShell UI Automation to interact directly with PowerPoint's native accessibility engine, the following environment is required:

| Requirement | Supported Version |
| :--- | :--- |
| **Operating System** | Windows 10 or Windows 11 |
| **Microsoft Office** | Desktop Microsoft PowerPoint (Microsoft 365, Office 2016, 2019, or 2021) |
| **Python** | Python 3.8 or newer (ensure Python is added to your system `PATH`) |

> **Note**: This tool runs via Windows COM APIs (`win32com.client`). It cannot run natively on macOS or Linux without a Windows virtual machine.

---

## 🚀 Quick Start Guide

Follow these simple steps on any Windows laptop or PC:

### Step 1: Clone or Download the Repository

Open **PowerShell** or **Command Prompt** and clone this repository:

```powershell
git clone https://github.com/Dedeepya777/Accessibility_Slides.git
cd Accessibility_Slides
```

*(Alternatively, click **Code > Download ZIP** on GitHub, extract the folder, and open PowerShell inside the extracted folder).*

---

### Step 2: (Recommended) Create a Virtual Environment

Creating an isolated environment keeps your Python packages clean:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

> **Tip**: If PowerShell gives an execution policy error when activating the virtual environment, run:
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> .\venv\Scripts\Activate.ps1
> ```

---

### Step 3: Install Required Dependencies

Install the project dependencies using `pip`:

```powershell
pip install -r requirements.txt
```

*(The primary dependency is `pywin32`, which allows Python to automate Microsoft PowerPoint).*

---

### Step 4: Run the Remediation Script

Place your `.pptx` presentation in the project folder, then run:

```powershell
python universal.py "YourPresentation.pptx"
```

#### Automatic Mode (No Arguments)
If you run without arguments, `universal.py` automatically scans the current directory and processes the first `.pptx` file found:

```powershell
python universal.py
```

---

## ⚙️ Command-Line Options

| Command | Description |
| :--- | :--- |
| `python universal.py "deck.pptx"` | Remediates `deck.pptx` and exports an accessible tagged PDF. |
| `python universal.py "deck.pptx" --no-pdf` | Remediates `deck.pptx` without generating a PDF file. |
| `python universal.py "deck.pptx" --pdf` | Explicitly enables tagged PDF export (enabled by default). |
| `python universal.py --help` | Displays command usage and built-in documentation. |

---

## 📁 Generated Output Files

When `universal.py` finishes, you will see three files in your directory:

1. **`YourPresentation.pptx`**: The remediated presentation file saved with all accessibility fixes applied.
2. **`YourPresentation.pptx.backup.pptx`**: A safety copy of your original, untouched presentation created before any modifications.
3. **`YourPresentation.pdf`**: A tagged, WCAG/PDF-UA compliant PDF export ready for distribution.

---

## 📊 Sample Audit Output

At completion, `universal.py` prints a before-and-after audit report directly from PowerPoint's Accessibility Checker:

```text
==============================================================================
          UNIVERSAL POWERPOINT ACCESSIBILITY COMPLIANCE REPORT
==============================================================================
Target File     : sample_slides.pptx
Total Slides    : 24
Visual Changes  : ZERO (100% Visually Lossless)
------------------------------------------------------------------------------
Accessibility Metric                | Before     | After      | Status
------------------------------------------------------------------------------
Missing slide title                 | 8          | 0          | RESOLVED (0)
Missing alt text                    | 15         | 0          | RESOLVED (0)
Unclear hyperlink text              | 4          | 0          | RESOLVED (0)
Hard-to-read text contrast          | 0          | 0          | RESOLVED (0)
Duplicate slide title               | 3          | 0          | RESOLVED (0)
Missing table header                | 2          | 0          | RESOLVED (0)
Use of merged or split cells        | 0          | 0          | RESOLVED (0)
Missing audio or video subtitles    | 0          | 0          | RESOLVED (0)
Check reading order                 | 12         | 0          | RESOLVED (0)
------------------------------------------------------------------------------
Total Errors                        | 25         | 0          | PASSED
Total Warnings                      | 19         | 0          | PASSED
==============================================================================
OVERALL STATUS: ACCESSIBLE (PASSED)
==============================================================================
```

---

## 🔧 Troubleshooting & Tips

- **Close open PowerPoint files before running**: If the target `.pptx` file is currently open in PowerPoint, close it first so the script can access and save the file without permission locks.
- **PowerPoint Dialog Boxes**: Make sure there are no modal popups or unsaved prompts open in PowerPoint in the background.
- **`ModuleNotFoundError: No module named 'win32com'`**:
  Make sure `pywin32` was installed in your active environment:
  ```powershell
  pip install pywin32
  ```
- **Execution Policy Error on Windows**:
  Run PowerShell as Administrator or run:
  ```powershell
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
  ```

---

## 📄 License

This project is licensed under the [MIT License](LICENSE). Feel free to use, modify, and distribute it in your own workflows.
