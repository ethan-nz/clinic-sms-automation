# 🏥 Clinic SMS Appointment Reminder Automation

A desktop automation tool that reads daily clinic booking data from Excel, generates personalised bilingual (Chinese / English) appointment reminder messages, and sends them via the Google Messages — with zero manual copy-pasting.

Built by a practising acupuncturist with 13 years of clinical experience who also writes automation tools.

---

## ✨ Key Features

- **Reads booking Excel files** exported from any clinic management system
- **Generates bilingual reminders** — Chinese by default, English on demand (per patient)
- **SIM card routing** — automatically selects the correct SIM based on practitioner type
- **Human-in-the-loop sending** — you review and click to send each message (no accidental bulk sends)
- **Live message preview** — check all queued messages before sending anything
- **Coordinate calibration UI** — one-time visual setup; no code editing needed when the screen layout changes
- **ESC to stop** — interrupt at any point with the keyboard
- **Local-first, privacy-safe** — no internet connection, no cloud upload, all data stays on the clinic machine

---

## 🖥️ Requirements

| Requirement | Detail |
|---|---|
| OS | **Windows only** (uses `win32gui` and `pygetwindow` for window management) |
| Python | 3.9 or higher |
| Messages app | Google Messages (or any SMS app accessible via keyboard/mouse) |

---

## 📦 Installation

```bash
# 1. Clone the repository
git clone https://github.com/ethan-nz/clinic-sms-automation.git
cd clinic-sms-automation

# 2. Create a virtual environment (recommended)
python -m venv venv
venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt
```

---

## ⚙️ Configuration

### Step 1 — Edit `clinic_data_demo.json`

This file defines your clinic branches, practitioners, and special staff rules.
A template is included — fill in your own data:

```json
{
  "special_staff": {
    "staff_w": "Dr Smith",
    "staff_l": "Dr Lee"
  },
  "branches": {
    "city": {
      "display_name_chi": "城市诊所",
      "display_name_eng": "City Clinic",
      "phone": "09-XXX-XXXX",
      "has_weekend_parking": true,
      "parking_instruction_chi": "周末请停P2停车场。",
      "parking_instruction_eng": "Weekend parking: Level P2.",
      "custom_greeting_chi": "您好，",
      "custom_blessing_chi": "祝您身体健康！",
      "practitioners": [
        { "name": "Dr Smith", "treatment_type": "massage" },
        { "name": "Dr Lee",   "treatment_type": "moxa" }
      ]
    }
  }
}
```

Supported treatment types: `acupuncture`, `massage`, `physio`, `chiropractic`, `moxa`

### Step 2 — Prepare your booking file

Name your Excel export `bookings_demo.xlsx` and place it in the same folder as the script.

Required columns: `Patient`, `PhoneNumber`, `Date`, `Time`, `Doctor`
Optional columns: `Status` (rows marked `Cancelled` are auto-skipped), `Tag`, `Notes`

> Tip: Add `英文` anywhere in the `Notes` field to flag a patient for English reminders.

### Step 3 — Calibrate screen coordinates (first run only)

1. Open your Messages app and position it on screen
2. Launch the tool: `python text_auto_Demo.py`
3. Click **Settings → Calibrate Coordinates**
4. Follow the on-screen prompts to click each UI element once
5. Coordinates are saved to `coordinates.json` automatically

---

## 🚀 Usage

```bash
python text_auto_Demo.py
```

1. The tool loads today's bookings and displays them in a table
2. All patients are pre-checked (✓) — uncheck any you want to skip
3. Toggle the **EN** column to switch a patient to English
4. Click **🔍 Check** to preview all messages before sending
5. Click **▶ Start** — for each patient, the tool opens Messages and pastes the text; you click once to send
6. Press **ESC** or click **⬛ Stop** at any time to interrupt

---

## 🧠 How It Works

```
bookings.xlsx
     │
     ▼
 read_info()           ← parses Excel, groups by phone + date
     │
     ▼
 MessageRouter         ← decides SIM card (Rules 1–3, then default)
     │
     ▼
 MessageGenerator      ← builds Chinese or English reminder text
     │
     ▼
 make_text()           ← drives Messages app via pyautogui
     │
     ▼
 PatientTrackingWindow ← live status table (Pending / Processing / Completed / Error)
```

**SIM routing rules (first match wins):**

| Rule | Condition | SIM |
|---|---|---|
| 1 | Staff W (massage specialist) | SIM 1 |
| 2 | Physio or chiropractic | SIM 1 |
| 3 | Staff L (moxa specialist) | SIM 2 |
| Default | Everything else | SIM 1 |

---

## 🔒 Privacy & Compliance

- All processing is **local** — no data leaves the machine
- Real clinic data files (`clinic_data.json`, `bookings_*.xlsx`) are excluded from this repository via `.gitignore`
- The demo files contain no real patient information
- Designed with the **NZ Health Information Privacy Code 2020** in mind

---

## 📁 Project Structure

```
clinic-sms-automation/
├── text_auto_Demo.py       # Main application
├── clinic_data_demo.json   # Demo clinic configuration (template)
├── bookings_demo.xlsx      # Demo booking data (sample)
├── requirements.txt
├── .gitignore
└── README.md
```

## Demo

<img width="800" height="450" alt="text_auto Demo" src="https://github.com/user-attachments/assets/c0b2deef-208e-4221-869a-5168a1862635" />

---

## 👤 Author

Built by **Ethan** — acupuncturist turned healthcare automation engineer.  
13 years of clinical experience informing every design decision.

---

## 📄 License

MIT License — see [LICENSE](LICENSE) for details.
