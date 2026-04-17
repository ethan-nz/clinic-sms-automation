import time
import json
import pandas as pd
import pyautogui
import pyperclip
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple
from enum import Enum
import tkinter as tk
from tkinter import ttk, messagebox
from threading import Thread, Event
from pynput.mouse import Listener, Button
from pynput import keyboard
import os
from pathlib import Path
import win32gui
import pygetwindow as gw
from PIL import ImageChops


# ============================================================================
# CONSTANTS
# ============================================================================

DAYS_OF_WEEK_CHI: Dict[str, str] = {
    'Monday': '周一', 'Tuesday': '周二', 'Wednesday': '周三',
    'Thursday': '周四', 'Friday': '周五', 'Saturday': '周六', 'Sunday': '周日'
}

# These 0/1 values are LIST INDICES, not labels.
# The UI always displays them as "SIM 1" and "SIM 2" via:  sim_num + 1
SIM_ONE = 0
SIM_TWO = 1


class TreatmentType(Enum):
    CHIROPRACTIC = "chiropractic"
    MASSAGE      = "massage"
    PHYSIO       = "physio"
    ACUPUNCTURE  = "acupuncture"
    MOXA         = "moxa"


TREATMENT_NAMES_CHI = {
    TreatmentType.CHIROPRACTIC: '整脊',
    TreatmentType.MASSAGE:      '推拿',
    TreatmentType.PHYSIO:       '物理',
    TreatmentType.ACUPUNCTURE:  '针灸',
    TreatmentType.MOXA:         '',
}


# ============================================================================
# CONFIGURATION & DATA STRUCTURES
# ============================================================================

@dataclass
class Coordinates:
    """Stores UI coordinates for automation."""
    start_chat:        Tuple[int, int]           = (100, 213)
    phone_input:       Tuple[int, int]           = (536, 275)
    confirm_phone:     Tuple[int, int]           = (569, 400)
    text_input:        Tuple[int, int]           = (800, 980)
    sim_dropdown:      Tuple[int, int]           = (700, 980)
    sim1:              Tuple[int, int]           = (755, 880)
    sim2:              Tuple[int, int]           = (755, 920)
    screenshot_region: Tuple[int, int, int, int] = (412, 422, 588, 515)

    def save(self, filepath: str = "coordinates.json") -> None:
        with open(filepath, 'w') as f:
            json.dump(self.__dict__, f, indent=2)

    @classmethod
    def load(cls, filepath: str = "coordinates.json") -> 'Coordinates':
        if os.path.exists(filepath):
            with open(filepath, 'r') as f:
                data = json.load(f)
            return cls(**data)
        return cls()


@dataclass
class Practitioner:
    name:           str
    treatment_type: TreatmentType

    def matches_name(self, search_name: str) -> bool:
        return self.name.strip() == search_name.strip()


@dataclass
class Branch:
    name:                    str
    display_name_chi:        str
    display_name_eng:        str
    phone:                   str
    practitioners:           List[Practitioner]
    has_weekend_parking:     bool = False
    parking_instruction_chi: str  = ""
    parking_instruction_eng: str  = ""
    custom_greeting_chi:     str  = ""
    custom_blessing_chi:     str  = ""

    def get_parking_message(self, day_of_week_chi: Optional[str],
                            day_of_week_eng: Optional[str],
                            language: str = 'chi') -> str:
        if not self.has_weekend_parking:
            return ""
        if language == 'chi':
            return self.parking_instruction_chi if day_of_week_chi in ['周六', '周日'] else ""
        return self.parking_instruction_eng if day_of_week_eng in ['Saturday', 'Sunday'] else ""


class ClinicConfiguration:
    """Central configuration loaded from an external JSON file."""

    def __init__(self, data_file: str = "clinic_data_demo.json"):
        self.branches:      Dict[str, Branch] = {}
        self.special_staff: Dict[str, str]    = {}
        self.coordinates = Coordinates.load()
        self._practitioner_index: Dict[str, Tuple[Branch, Practitioner]] = {}

        self._load_from_json(data_file)
        self._build_practitioner_index()

    def _load_from_json(self, filepath: str) -> None:
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"Missing configuration file: {filepath}")

        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)

        self.special_staff = data.get("special_staff", {})

        for branch_name, b_data in data.get("branches", {}).items():
            practitioners = []
            for p_data in b_data.get("practitioners", []):
                try:
                    t_type = TreatmentType(p_data["treatment_type"])
                    practitioners.append(Practitioner(name=p_data["name"],
                                                      treatment_type=t_type))
                except ValueError:
                    print(f"Warning: unknown treatment type "
                          f"'{p_data['treatment_type']}' for {p_data['name']}")

            self.branches[branch_name] = Branch(
                name=branch_name,
                display_name_chi=b_data.get("display_name_chi", ""),
                display_name_eng=b_data.get("display_name_eng", ""),
                phone=b_data.get("phone", ""),
                has_weekend_parking=b_data.get("has_weekend_parking", False),
                parking_instruction_chi=b_data.get("parking_instruction_chi", ""),
                parking_instruction_eng=b_data.get("parking_instruction_eng", ""),
                custom_greeting_chi=b_data.get("custom_greeting_chi", ""),
                custom_blessing_chi=b_data.get("custom_blessing_chi", ""),
                practitioners=practitioners,
            )

    def _build_practitioner_index(self) -> None:
        """O(1) lookup index: practitioner name → (Branch, Practitioner)."""
        for branch in self.branches.values():
            for prac in branch.practitioners:
                self._practitioner_index[prac.name.strip()] = (branch, prac)

    def find_practitioner_branch(self, name: str) -> Optional[Tuple[Branch, Practitioner]]:
        return self._practitioner_index.get(name.strip())

    def save_coordinates(self) -> None:
        self.coordinates.save()


# ============================================================================
# GLOBAL THREADING STATE
# ============================================================================

interrupt_event   = Event()
keyboard_listener = None


# ============================================================================
# KEYBOARD LISTENER
# ============================================================================

def on_press(key) -> bool:
    if key == keyboard.Key.esc:
        interrupt_event.set()
        return False
    return True

def start_keyboard_listener() -> None:
    global keyboard_listener
    if keyboard_listener is None or not keyboard_listener.is_alive():
        keyboard_listener = keyboard.Listener(on_press=on_press)
        keyboard_listener.start()

def stop_keyboard_listener() -> None:
    global keyboard_listener
    if keyboard_listener is not None and keyboard_listener.is_alive():
        keyboard_listener.stop()
        keyboard_listener.join()
        keyboard_listener = None


# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def get_greeting() -> str:
    return "Good morning!" if datetime.now().hour < 12 else "Good afternoon!"


def compare_date(date_variable: datetime) -> str:
    """Return a Chinese relative-date label (明天 / 后天 / 今天 / '')."""
    today     = datetime.now().date()
    tomorrow  = today + timedelta(days=1)
    day_after = today + timedelta(days=2)

    if date_variable.date() == tomorrow:   return '明天'
    if date_variable.date() == day_after:  return '后天'
    if date_variable.date() == today:      return '今天'
    return ''


def screenshot_until_change(config: ClinicConfiguration,
                             interval: float = 0.5,
                             timeout:  float = 2.0) -> None:
    """Poll a screen region and return as soon as the UI updates."""
    region = config.coordinates.screenshot_region
    try:
        prev_img = pyautogui.screenshot(region=region)
    except Exception as e:
        print(f"Screenshot error: {e}")
        return

    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(interval)
        curr_img = pyautogui.screenshot(region=region)
        if ImageChops.difference(prev_img, curr_img).getbbox() is not None:
            return


def lock_click(position, interrupt_ev: Event) -> bool:
    """Move cursor to position and click, tolerating Windows DPI scaling drift."""
    if interrupt_ev.is_set():
        return False

    target_x, target_y = int(position[0]), int(position[1])
    pyautogui.moveTo(target_x, target_y, duration=0.5)

    for _ in range(5):
        if interrupt_ev.is_set():
            return False
        cx, cy = pyautogui.position()
        if abs(cx - target_x) <= 3 and abs(cy - target_y) <= 3:
            pyautogui.click()
            return True
        pyautogui.moveTo(target_x, target_y)
        time.sleep(0.1)

    pyautogui.click()  # close enough — click anyway
    return True


# ============================================================================
# COORDINATE CALIBRATION WINDOW
# ============================================================================

class CoordinateCalibrationWindow:
    COORDS_TO_CAPTURE = [
        ("Start Chat",                "start_chat"),
        ("Phone Input",               "phone_input"),
        ("Confirm Phone",             "confirm_phone"),
        ("Text Input",                "text_input"),
        ("SIM Dropdown",              "sim_dropdown"),
        ("SIM 1",                     "sim1"),
        ("SIM 2",                     "sim2"),
        ("Status Area (Top-Left)",    "region_tl"),
        ("Status Area (Bottom-Right)","region_br"),
    ]

    def __init__(self, config: ClinicConfiguration, parent=None):
        self.config          = config
        self.root            = tk.Toplevel(parent) if parent else tk.Tk()
        self.captured_coords: Dict[str, Tuple[int, int]] = {}
        self.current_index   = 0
        self.listener        = None

        self.root.title("Coordinate Calibration")
        self.root.geometry("450x700+1000+100")
        self.root.attributes('-topmost', True)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self._setup_ui()

    def _setup_ui(self) -> None:
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="INSTRUCTIONS:", font=("Arial", 10, "bold")).pack(anchor=tk.W)
        tk.Label(
            main_frame,
            text="1. Click 'Start Capture'.\n"
                 "2. Click the target location in the Messages app.\n"
                 "3. This window stays visible so you can follow each step.",
            justify=tk.LEFT, wraplength=400,
        ).pack(pady=5)

        self.progress_label = tk.Label(main_frame, text="Ready",
                                       font=("Arial", 12, "bold"), fg="blue")
        self.progress_label.pack(pady=10)

        list_frame = ttk.LabelFrame(main_frame, text="Current Captures", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.coord_labels: Dict[str, tk.Label] = {}
        for display_name, key in self.COORDS_TO_CAPTURE:
            row = ttk.Frame(list_frame)
            row.pack(fill=tk.X, pady=2)
            tk.Label(row, text=f"{display_name}:", width=22, anchor=tk.W).pack(side=tk.LEFT)
            lbl = tk.Label(row, text="--", fg="gray")
            lbl.pack(side=tk.LEFT, padx=10)
            self.coord_labels[key] = lbl

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=20)
        self.capture_btn = ttk.Button(btn_frame, text="Start Capture",
                                      command=self.start_capture)
        self.capture_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save & Close",
                   command=self.save_and_close).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel",
                   command=self.on_closing).pack(side=tk.LEFT, padx=5)

    def start_capture(self) -> None:
        self.current_index = 0
        self.capture_btn.config(state=tk.DISABLED)
        self._capture_next()

    def _capture_next(self) -> None:
        if self.current_index >= len(self.COORDS_TO_CAPTURE):
            self.progress_label.config(text="✓ All points captured!", fg="green")
            self.capture_btn.config(state=tk.NORMAL)
            return
        name, key = self.COORDS_TO_CAPTURE[self.current_index]
        self.progress_label.config(text=f"Click on: {name}", fg="red")
        self._start_mouse_listener(key)

    def _start_mouse_listener(self, key: str) -> None:
        def on_click(x, y, button, pressed):
            if button == Button.left and not pressed:
                self.root.after(0, self._process_click, key, int(x), int(y))
                return False

        self.listener = Listener(on_click=on_click)
        self.listener.start()

    def _process_click(self, key: str, x: int, y: int) -> None:
        self.captured_coords[key] = (x, y)
        self.coord_labels[key].config(text=f"({x}, {y})", fg="black")
        self.current_index += 1
        self.root.after(400, self._capture_next)

    def save_and_close(self) -> None:
        if self.listener and self.listener.is_alive():
            self.listener.stop()

        for _, key in self.COORDS_TO_CAPTURE[:-2]:
            if key in self.captured_coords:
                setattr(self.config.coordinates, key, self.captured_coords[key])

        if 'region_tl' in self.captured_coords and 'region_br' in self.captured_coords:
            tl = self.captured_coords['region_tl']
            br = self.captured_coords['region_br']
            self.config.coordinates.screenshot_region = (
                tl[0], tl[1], abs(br[0] - tl[0]), abs(br[1] - tl[1])
            )

        self.config.save_coordinates()
        messagebox.showinfo("Success", "Coordinates updated!")
        self.root.destroy()

    def on_closing(self) -> None:
        if self.listener and self.listener.is_alive():
            self.listener.stop()
        self.root.destroy()


# ============================================================================
# CORE BUSINESS LOGIC — DATA INGESTION
# ============================================================================

def read_info(filepath: str, config: ClinicConfiguration) -> Dict[str, dict]:
    """
    Parse the booking Excel file and return a patient dictionary.
    """
    if not filepath or not os.path.exists(filepath):
        print(f"File not found: {filepath}")
        return {}

    try:
        df = pd.read_excel(filepath, dtype={'PhoneNumber': str})
        df['PhoneNumber'] = df['PhoneNumber'].astype(str).str.replace(r'\.0$', '', regex=True)
        df = df.dropna(subset=['PhoneNumber', 'Date', 'Time'])
        df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)
        df['Time'] = pd.to_datetime(df['Time'], format='%H:%M').dt.strftime('%H:%M')

        # Guard optional columns — demo spreadsheets may not have these
        if 'Status' in df.columns:
            df = df[~df['Status'].eq('Cancelled')]
        if 'Tag' in df.columns:
            df = df[~df['Tag'].eq('QI')]
        if 'Patient' in df.columns:
            df = df[~df['Patient'].str.contains('No No', case=False, na=False)]

        df['date_only'] = df['Date'].dt.date

        data_dict: Dict[str, dict] = {}

        for _group_key, group in df.groupby(['PhoneNumber', 'date_only']):
            is_multi = len(group) > 1
            earliest = group.sort_values('Time').iloc[0]
            patient_name = earliest['Patient']

            notes = str(earliest.get('Notes', '')) if 'Notes' in earliest.index else ''

            data_dict[patient_name] = {
                'phone':      earliest['PhoneNumber'],
                'doctor':     earliest['Doctor'],
                'date':       earliest['Date'],
                'time':       earliest['Time'],
                'use_english': '英文' in notes,
                'is_multi_appointment': is_multi,
            }

        return data_dict

    except Exception as e:
        print(f"Error reading file: {e}")
        return {}


# ============================================================================
# CORE BUSINESS LOGIC — MESSAGE GENERATION
# ============================================================================

class MessageGenerator:
    """Generates appointment reminder messages in Chinese or English."""

    def __init__(self, config: ClinicConfiguration):
        self.config = config

    def generate_message(self, patient_data: dict, use_english: bool) -> str:
        doctor_name     = patient_data['doctor']
        booking_date_dt = patient_data['date']
        booking_time    = patient_data['time']

        result = self.config.find_practitioner_branch(doctor_name)
        if not result:
            return self._generic_message(patient_data, use_english)

        branch, practitioner = result

        formatted_date  = booking_date_dt.strftime('%d/%m')
        day_of_week_eng = booking_date_dt.strftime("%A")
        day_of_week_chi = DAYS_OF_WEEK_CHI[day_of_week_eng]
        relative_date   = compare_date(booking_date_dt)
        treatment_chi   = TREATMENT_NAMES_CHI.get(practitioner.treatment_type, '治疗')

        if use_english:
            return self._english_message(branch, formatted_date, booking_time,
                                         day_of_week_eng, relative_date)
        return self._chinese_message(branch, formatted_date, booking_time,
                                     day_of_week_chi, relative_date, treatment_chi)

    def _chinese_message(self, branch: Branch, formatted_date: str,
                         formatted_time: str, day_of_week_chi: str,
                         relative_date: str, treatment_chi: str) -> str:
        parking = branch.get_parking_message(day_of_week_chi, None, 'chi')
        return (
            f'{branch.custom_greeting_chi}您预约了{relative_date}{day_of_week_chi} '
            f'({formatted_date}) {formatted_time}的{treatment_chi}治疗。{parking}\n'
            f'如有任何问题，请致电诊所座机：{branch.phone}\n'
            f'温馨提示：若不能就诊，请提前24小时联系我们，谢谢。{branch.custom_blessing_chi}\n'
            f'中医诊所{branch.display_name_chi}'
        )

    def _english_message(self, branch: Branch, formatted_date: str,
                         formatted_time: str, day_of_week_eng: str,
                         relative_date: str) -> str:
        date_label = {
            '明天': 'tomorrow', '今天': 'today', '后天': 'the day after tomorrow'
        }.get(relative_date, relative_date)
        parking  = branch.get_parking_message(None, day_of_week_eng, 'eng')
        greeting = get_greeting()
        return (
            f'{greeting} You have booked an appointment on '
            f'{date_label} ({day_of_week_eng} {formatted_date}) at {formatted_time}.{parking}\n'
            f'Please confirm receipt. Thank you!\n'
            f'Contact: {branch.phone}\n'
            f'Note: contact us 24 hrs in advance if unable to attend.\n'
            f'Chinese Medical Clinic {branch.display_name_eng}'
        )

    def _generic_message(self, patient_data: dict, use_english: bool) -> str:
        """Fallback when practitioner is not found in any branch."""
        formatted_date = patient_data['date'].strftime('%d/%m')
        formatted_time = patient_data['time']
        if use_english:
            return f"Appointment reminder: {formatted_date} at {formatted_time}. Please confirm."
        return f"预约提醒: {formatted_date} {formatted_time}。请确认。"


# ============================================================================
# CORE BUSINESS LOGIC — MESSAGE ROUTING
# ============================================================================

class MessageRouter:
    """
    Decides which SIM card to use for each patient.

    SIM CARD INDEX REFERENCE (see SIM_ONE / SIM_TWO constants at the top):
      SIM_ONE (0) → coords.sim1 → physical SIM card 1  (massage, physio, chiro, default)
      SIM_TWO (1) → coords.sim2 → physical SIM card 2  (moxa)

    Rules applied in priority order — first match wins.

    ROUTING SUMMARY:
      Rule 1 — Staff_W (Massage specialist) → SIM 1
      Rule 2 — Physio / Chiropractic        → SIM 1
      Rule 3 — Staff_L (Moxa specialist)    → SIM 2
      Default — everything else             → SIM 1
    """

    def __init__(self, config: ClinicConfiguration):
        self.config = config

    def _rule1_staff_w(self, practitioner_name, **_):
        """Massage specialist always uses SIM 1."""
        if practitioner_name == self.config.special_staff.get('staff_w'):
            return SIM_ONE, "Rule 1: Staff_W (Massage) → SIM 1"
        return None

    def _rule2_physio_chiro(self, practitioner, **_):
        """Physio and chiropractic always use SIM 1."""
        if practitioner.treatment_type in (TreatmentType.PHYSIO, TreatmentType.CHIROPRACTIC):
            return SIM_ONE, f"Rule 2: {practitioner.treatment_type.value} → SIM 1"
        return None

    def _rule3_staff_l(self, practitioner_name, **_):
        """Moxa specialist always uses SIM 2."""
        staff_l_name = self.config.special_staff.get('staff_l', 'Staff_L')
        if practitioner_name.startswith(staff_l_name):
            return SIM_TWO, "Rule 3: Staff_L (Moxa) → SIM 2"
        return None

    # Ordered list — first match wins.
    RULES = [
        _rule1_staff_w,
        _rule2_physio_chiro,
        _rule3_staff_l,
    ]

    def determine_message_params(
        self, patient_data: dict, use_english: bool
    ) -> Tuple[Optional[int], str, str]:
        """
        Returns: (sim_number, message_text, decision_reason)

        sim_number is None only when the practitioner is not found.
        """
        practitioner_name = patient_data.get('doctor', '')

        result = self.config.find_practitioner_branch(practitioner_name)
        if not result:
            return None, "", f"Practitioner '{practitioner_name}' not found"

        branch, practitioner = result

        # Build context — only fields the active rules actually use.
        ctx = dict(
            practitioner=practitioner,
            practitioner_name=practitioner_name,
            branch=branch,
        )

        sim_number, reason = None, ""

        for rule in self.RULES:
            outcome = rule(self, **ctx)
            if outcome is not None:
                sim_number, reason = outcome
                break

        # Default fallback — no rule matched (e.g. plain acupuncture in demo)
        if sim_number is None:
            sim_number = SIM_ONE
            reason = "Default fallback → SIM 1"

        message_text = MessageGenerator(self.config).generate_message(patient_data, use_english)
        return sim_number, message_text, reason


# ============================================================================
# MESSAGE SENDING — AUTOMATION
# ============================================================================

def select_sim(n: int, config: ClinicConfiguration, interrupt_ev: Event) -> bool:
    coords   = config.coordinates
    sim_list = [coords.sim1, coords.sim2]   # index 0 = SIM1, index 1 = SIM2

    if not lock_click(coords.sim_dropdown, interrupt_ev):
        return False
    time.sleep(0.3)
    return lock_click(sim_list[n], interrupt_ev)


def wait_for_mouse_click(interrupt_ev: Event, timeout: float = 300.0) -> bool:
    """Block until the user left-clicks (manual confirmation to send)."""
    clicked = Event()

    def on_click(x, y, button, pressed):
        if button == Button.left and not pressed:
            clicked.set()
            return False
        if interrupt_ev.is_set():
            return False

    listener = Listener(on_click=on_click)
    listener.start()

    deadline = time.time() + timeout
    while not clicked.is_set():
        if interrupt_ev.is_set():
            break
        if time.time() > deadline:
            interrupt_ev.set()
            break
        time.sleep(0.1)

    listener.stop()
    if interrupt_ev.is_set():
        return False
    time.sleep(0.2)
    return True


def make_text(phone: str, text: str, sim_number: int,
              config: ClinicConfiguration, interrupt_ev: Event) -> bool:
    """Drive the Messages app via pyautogui to send a single SMS."""
    if interrupt_ev.is_set():
        return False
    if not phone or len(str(phone)) <= 5:
        return True  # no valid number — skip silently

    coords = config.coordinates

    if not lock_click(coords.start_chat, interrupt_ev): return False
    time.sleep(1)
    if not lock_click(coords.phone_input, interrupt_ev): return False

    pyperclip.copy(str(phone))
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)

    if not lock_click(coords.confirm_phone, interrupt_ev): return False
    time.sleep(0.5)

    screenshot_until_change(config)

    if not select_sim(sim_number, config, interrupt_ev): return False
    if not lock_click(coords.text_input, interrupt_ev): return False

    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')

    if not wait_for_mouse_click(interrupt_ev): return False

    pyautogui.press("enter")
    time.sleep(1)
    return True


# ============================================================================
# GUI — CHECK PREVIEW WINDOW
# ============================================================================

class CheckPreviewWindow:
    """Scrollable preview of all queued messages before sending."""

    def __init__(self, parent: tk.Tk, results: list):
        self.win     = tk.Toplevel(parent)
        self.results = results
        self.win.title(f"Message Preview — {len(results)} patient(s)")
        self.win.geometry("700x600+200+100")
        self.win.attributes('-topmost', True)
        self._build_ui()

    def _build_ui(self) -> None:
        tk.Label(self.win,
                 text=f"Previewing {len(self.results)} checked patient(s)",
                 font=("Arial", 11, "bold")).pack(anchor=tk.W, padx=12, pady=(10, 4))

        container = ttk.Frame(self.win)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        canvas    = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        inner         = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=inner, anchor='nw')

        def _on_canvas_resize(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind('<Configure>', _on_canvas_resize)

        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox('all'))
        inner.bind('<Configure>', _on_frame_configure)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        canvas.bind_all('<MouseWheel>', _on_mousewheel)

        for entry in self.results:
            self._add_card(inner, entry)

        toolbar = ttk.Frame(self.win)
        toolbar.pack(fill=tk.X, padx=10, pady=(4, 10))
        ttk.Button(toolbar, text="Copy All Messages",
                   command=self._copy_all).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="Close",
                   command=self.win.destroy).pack(side=tk.RIGHT, padx=4)

    def _add_card(self, parent: ttk.Frame, entry: dict) -> None:
        card = ttk.LabelFrame(parent, text=entry['name'], padding="8")
        card.pack(fill=tk.X, padx=6, pady=5)

        if 'error' in entry:
            tk.Label(card, text=f"Error: {entry['error']}",
                     fg='red', anchor=tk.W).pack(fill=tk.X)
            return

        meta = (f"Phone: {entry['phone']}    "
                f"SIM {entry['sim']}    "
                f"Rule: {entry['reason']}")
        tk.Label(card, text=meta, fg='#444444', anchor=tk.W,
                 font=("Arial", 9)).pack(fill=tk.X, pady=(0, 4))

        msg_frame   = ttk.Frame(card)
        msg_frame.pack(fill=tk.X)
        text_widget = tk.Text(msg_frame, height=5, wrap=tk.WORD,
                               font=("Arial", 10), relief=tk.FLAT,
                               bg='#f9f9f9', fg='#111111')
        text_widget.insert('1.0', entry['message'])
        text_widget.config(state=tk.DISABLED)
        text_widget.pack(side=tk.LEFT, fill=tk.X, expand=True)

        def _copy(msg=entry['message']):
            pyperclip.copy(msg)
        ttk.Button(msg_frame, text="Copy", width=6,
                   command=_copy).pack(side=tk.RIGHT, padx=(6, 0))

    def _copy_all(self) -> None:
        parts = []
        for entry in self.results:
            if 'message' in entry:
                parts.append(f"[{entry['name']} | SIM {entry['sim']}]\n{entry['message']}")
            else:
                parts.append(f"[{entry['name']}] ERROR: {entry.get('error', '?')}")
        pyperclip.copy("\n\n---\n\n".join(parts))
        messagebox.showinfo("Copied", f"{len(parts)} message(s) copied to clipboard.")


# ============================================================================
# GUI — PATIENT TRACKING WINDOW
# ============================================================================

class PatientTrackingWindow:
    COLORS = {
        'pending':    '#ffffff',
        'processing': '#fff7e6',
        'completed':  '#e6ffe6',
        'error':      '#ffe6e6',
    }

    def __init__(self, patients_dict: Dict[str, dict], config: ClinicConfiguration):
        self.patients_dict = patients_dict or {}
        self.config        = config
        self.should_stop   = False

        self.root = tk.Tk()
        self.root.title("Patient Text Message Tracking")
        self.root.geometry("900x800+1000+100")
        self.root.attributes('-topmost', True)

        self.doctors = ['All'] + sorted({
            data['doctor']
            for data in self.patients_dict.values()
            if 'doctor' in data
        })
        self.selected_doctor = tk.StringVar(value='All')
        self._setup_ui()

    def _setup_ui(self) -> None:
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="Calibrate Coordinates",
                                   command=self.open_calibration)
        settings_menu.add_separator()
        settings_menu.add_command(label="Exit", command=self.root.quit)

        filter_frame = ttk.LabelFrame(main_frame, text="Filter by Doctor", padding="5")
        filter_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        max_cols = 6
        for i, doctor in enumerate(self.doctors):
            ttk.Radiobutton(
                filter_frame, text=doctor, value=doctor,
                variable=self.selected_doctor, command=self.update_table,
            ).grid(row=i // max_cols, column=i % max_cols, padx=5, pady=2, sticky="w")

        self._create_table(main_frame)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=2, column=0, pady=(20, 0), sticky="ew")
        btn_frame.columnconfigure((0, 1, 2, 3), weight=1)

        self.start_button = ttk.Button(btn_frame, text="▶ Start",
                                        command=self.start_process)
        self.start_button.grid(row=0, column=0, padx=5, sticky="ew")

        self.stop_button = ttk.Button(btn_frame, text="⬛ Stop",
                                       command=self.stop_process, state='disabled')
        self.stop_button.grid(row=0, column=1, padx=5, sticky="ew")

        self.check_button = ttk.Button(btn_frame, text="🔍 Check",
                                        command=self.check_process)
        self.check_button.grid(row=0, column=2, padx=5, sticky="ew")

        ttk.Button(btn_frame, text="⚙ Calibrate",
                   command=self.open_calibration).grid(row=0, column=3, padx=5, sticky="ew")

        self.status_bar = ttk.Label(main_frame, text="Ready",
                                     relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=3, column=0, sticky="ew", pady=(10, 0))

    def _create_table(self, parent: ttk.Frame) -> None:
        table_frame = ttk.Frame(parent)
        table_frame.grid(row=1, column=0, sticky="nsew")
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        columns = ('Check', 'Name', 'Time', 'Doctor', 'Phone', 'Status', 'Translate')
        self.tree = ttk.Treeview(table_frame, columns=columns,
                                  show='headings', selectmode='extended')

        col_config = {
            'Check':     (30,  tk.CENTER, tk.NO),
            'Name':      (150, tk.W,      tk.YES),
            'Time':      (70,  tk.CENTER, tk.YES),
            'Doctor':    (120, tk.CENTER, tk.YES),
            'Phone':     (110, tk.CENTER, tk.YES),
            'Status':    (90,  tk.CENTER, tk.YES),
            'Translate': (50,  tk.CENTER, tk.NO),
        }
        display_text = {'Check': '✓', 'Translate': 'EN'}

        for col, (width, anchor, stretch) in col_config.items():
            self.tree.heading(col, text=display_text.get(col, col))
            self.tree.column(col, width=width, anchor=anchor, stretch=stretch)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

        for tag, color in self.COLORS.items():
            self.tree.tag_configure(tag, background=color)

        self.tree.bind("<ButtonRelease-1>", self.on_tree_click)
        self.update_table()

    def update_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        selected_doc = self.selected_doctor.get()
        for name, data in self.patients_dict.items():
            if selected_doc != 'All' and data.get('doctor') != selected_doc:
                continue
            en_mark = '✓' if data.get('use_english', False) else ' '
            self.tree.insert(
                '', 'end', iid=name,
                values=('✓', name, data.get('time', ''), data.get('doctor', ''),
                        data.get('phone', ''), 'Pending', en_mark),
                tags=('pending',),
            )

    def on_tree_click(self, event) -> None:
        item_id = self.tree.identify_row(event.y)
        column  = self.tree.identify_column(event.x)
        if not item_id or not column:
            return
        values    = list(self.tree.item(item_id, 'values'))
        col_index = int(column.replace('#', '')) - 1
        if col_index == 0:
            values[0] = ' ' if values[0] == '✓' else '✓'
        elif col_index == 6:
            values[6] = '✓' if values[6] == ' ' else ' '
            if item_id in self.patients_dict:
                self.patients_dict[item_id]['use_english'] = (values[6] == '✓')
        self.tree.item(item_id, values=tuple(values))

    def update_status(self, item_id: str, status: str) -> None:
        if not self.tree.exists(item_id):
            return
        try:
            values    = list(self.tree.item(item_id, 'values'))
            values[5] = status
            tag = {'Processing': 'processing', 'Completed': 'completed',
                   'Error': 'error', 'Pending': 'pending'}.get(status, 'pending')
            if status == 'Error':
                values[0] = ' '
            self.tree.item(item_id, values=tuple(values), tags=(tag,))
            self.root.update_idletasks()
        except (tk.TclError, IndexError):
            pass

    def start_process(self) -> None:
        self.should_stop = False
        interrupt_event.clear()
        self.start_button['state'] = 'disabled'
        self.stop_button['state']  = 'normal'
        start_keyboard_listener()
        Thread(target=self.process_messages, daemon=True).start()

    def stop_process(self) -> None:
        self.should_stop = True
        interrupt_event.set()

    def process_messages(self) -> None:
        items = [
            iid for iid in self.tree.get_children()
            if self.tree.item(iid)['values'][0] == '✓'
        ]
        total      = len(items)
        router     = MessageRouter(self.config)
        active_iid = None

        try:
            for index, iid in enumerate(items):
                active_iid = iid
                if self.should_stop or interrupt_event.is_set():
                    break

                values = self.tree.item(iid)['values']
                if not values or values[0] != '✓':
                    continue

                name        = values[1]
                use_english = (values[6] == '✓')

                self.update_status(iid, 'Processing')
                self.status_bar.config(text=f"Processing {index + 1}/{total}: {name}")

                if name not in self.patients_dict:
                    self.update_status(iid, 'Error')
                    continue

                patient_data = self.patients_dict[name]
                phone_number = patient_data.get('phone')

                sim_num, text_msg, reason = router.determine_message_params(
                    patient_data, use_english
                )

                if sim_num is None:
                    self.update_status(iid, 'Error')
                    continue

                if not make_text(phone_number, text_msg, sim_num,
                                  self.config, interrupt_event):
                    self.update_status(iid, 'Error')
                    if interrupt_event.is_set():
                        break
                    continue

                self.update_status(iid, 'Completed')
                self.tree.yview_moveto((index + 1) / max(total, 1))

                for _ in range(5):
                    if interrupt_event.is_set():
                        break
                    time.sleep(0.1)

        except Exception as e:
            print(f"Unexpected error: {e}")
            if active_iid:
                self.update_status(active_iid, 'Error')
        finally:
            self.start_button['state'] = 'normal'
            self.stop_button['state']  = 'disabled'
            self.status_bar.config(text="Process finished")
            interrupt_event.clear()
            stop_keyboard_listener()

    def check_process(self) -> None:
        """Preview all checked patients in a scrollable window."""
        checked_iids = [
            iid for iid in self.tree.get_children()
            if self.tree.item(iid)['values'][0] == '✓'
        ]
        if not checked_iids:
            messagebox.showwarning("No Selection", "Please select at least one patient.")
            return

        router  = MessageRouter(self.config)
        results = []

        for iid in checked_iids:
            values      = self.tree.item(iid)['values']
            name        = values[1]
            use_english = (values[6] == '✓')

            if name not in self.patients_dict:
                results.append({'name': name, 'error': 'Not found in patient data'})
                continue

            patient_data = self.patients_dict[name]
            sim_num, text_msg, reason = router.determine_message_params(
                patient_data, use_english
            )

            if sim_num is None:
                results.append({'name': name, 'error': reason})
            else:
                results.append({
                    'name':    name,
                    'phone':   patient_data.get('phone', '—'),
                    'sim':     sim_num + 1,   # display as 1-based ("SIM 1" / "SIM 2")
                    'reason':  reason,
                    'message': text_msg,
                })

        CheckPreviewWindow(self.root, results)

    def open_calibration(self) -> None:
        CoordinateCalibrationWindow(self.config, parent=self.root)

    def run(self) -> None:
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        try:
            self.root.mainloop()
        finally:
            self._cleanup()

    def _on_closing(self) -> None:
        self.should_stop = True
        interrupt_event.set()
        self._cleanup()
        self.root.destroy()

    def _cleanup(self) -> None:
        interrupt_event.set()
        stop_keyboard_listener()


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def find_latest_booking_file() -> Optional[str]:
    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))

    # Priority 1: explicit demo file
    demo_file = script_dir / "bookings_demo.xlsx"
    if demo_file.exists():
        print(f"Using demo file: {demo_file}")
        return str(demo_file)

    # Priority 2: any booking file alongside the script
    matches = list(script_dir.glob("bookings_*.xlsx"))
    if matches:
        latest = max(matches, key=lambda p: p.stat().st_ctime)
        print(f"Found booking file in script folder: {latest.name}")
        return str(latest)

    # Priority 3: Downloads folder fallback for real clinic machine
    download_folder = Path.home() / "Downloads"
    if not download_folder.exists():
        print(f"Downloads folder not found: {download_folder}")
        return None

    matches = list(download_folder.glob("bookings_*.xlsx"))
    if not matches:
        print("No booking files found in Downloads folder.")
        return None

    latest = max(matches, key=lambda p: p.stat().st_ctime)
    print(f"Using latest file from Downloads: {latest.name}")
    return str(latest)


def activate_app_window(title: str) -> None:
    try:
        windows = gw.getWindowsWithTitle(title)
        if windows:
            win = windows[0]
            if win.isMinimized:
                win.restore()
            win.activate()
            time.sleep(0.5)
            win32gui.SetForegroundWindow(win._hWnd)
    except Exception as e:
        print(f"Error activating window: {e}")


def main() -> None:
    pyautogui.FAILSAFE = True

    try:
        config = ClinicConfiguration()
    except FileNotFoundError as e:
        messagebox.showerror("Configuration Error", str(e))
        return

    filepath = find_latest_booking_file()
    if not filepath:
        messagebox.showerror("Error", "No booking file found!")
        return

    clients_dict = read_info(filepath, config)
    if not clients_dict:
        messagebox.showerror("Error", "No valid client data loaded!")
        return

    activate_app_window("Messages")
    PatientTrackingWindow(clients_dict, config).run()


if __name__ == "__main__":
    main()