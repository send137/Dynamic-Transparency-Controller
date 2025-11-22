import pickle
import os
import sys
import threading
import time
import random
import functools
import tkinter as tk
import win32api
import win32con # Import win32con for SW_SHOWNA and MONITOR_DEFAULTTOPRIMARY

import customtkinter
import ahk
import win32gui
import ctypes
import platform
import subprocess
import importlib.util

# --- Windows API Constants and Functions ---
GWL_EXSTYLE = -20
WS_EX_LAYERED = 0x00080000
LWA_ALPHA = 0x00000002
LWA_COLORKEY = 0x00000001

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ = 0x0010

try:
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    if ctypes.sizeof(ctypes.c_void_p) == 8:
        SetWindowLongPtrW = user32.SetWindowLongPtrW
        GetWindowLongPtrW = user32.GetWindowLongPtrW
    else:
        SetWindowLongPtrW = user32.SetWindowLongW
        GetWindowLongPtrW = user32.GetWindowLongW

    EnumWindows = user32.EnumWindows
    GetWindowThreadProcessId = user32.GetWindowThreadProcessId
    SetLayeredWindowAttributes = user32.SetLayeredWindowAttributes
    OpenProcess = kernel32.OpenProcess
    QueryFullProcessImageNameW = kernel32.QueryFullProcessImageNameW
    CloseHandle = kernel32.CloseHandle

except AttributeError as e:
    print(f"Error loading Windows API functions: {e}")
    print("This script is intended for Windows operating systems.")
    sys.exit(1)

# --- GLOBAL CONFIGURATION ---
INITIAL_WINDOW_SIZE = "450x900" # Slightly increased height for new settings
DEBUG_PRINT_MODIFIER_STATE_ON_MOUSE_EVENT = False
SCROLL_SEQUENCE_TIMEOUT_MS = 1000 # Fixed timeout (1s) to reset scroll sequence

# --- Default Settings for Persistence ---
DEFAULT_SETTINGS = {
    'theme_color': 'green',
    'appearance_mode': 'System',
    'hotkeys': {
        'increase_transparency': 'ctrl+wheelup',
        'decrease_transparency': 'ctrl+wheeldown',
        'set_86_percent': 'ctrl+xbutton2',
        'set_100_percent': 'ctrl+shift+xbutton2',
        'set_30_percent': 'ctrl+xbutton1',
        'toggle_script': 'alt+w',
        'kill_script_failsafe': 'ctrl+alt+shift+k',
        'center_window': 'ctrl+rbutton',
        'minimize_others': 'ctrl+shift+rbutton',
        'toggle_focus_mode': 'alt+q',
        'focus_mode_alt_tab': 'alt+tab',
        'increase_brightness': 'alt+wheelup',
        'decrease_brightness': 'alt+wheeldown',
        'set_80_percent_brightness': 'alt+xbutton2',
        'set_0_percent_brightness': 'alt+xbutton1',
    },
    'transparency_levels': {
        'initial': 49,
        'min': 10,
        'max': 100,
        'scroll_increment_slow': 1,
        'scroll_increment_fast': 2,
        'fast_scroll_threshold_ms': 40,
        'preset_xbutton2': 86,
        'preset_xbutton2_shift': 100,
        'preset_xbutton1': 30,
        'reset_on_scroll_start': False, 
    },
    'brightness_levels': {
        'initial': 49,
        'min': 0,
        'max': 100,
        'scroll_increment_slow': 1,
        'scroll_increment_fast': 4,
        'fast_scroll_threshold_ms': 30,
        'preset_xbutton2': 64,
        'preset_xbutton1': 0,
        'reset_on_scroll_start': False,
    },
    'script_enabled': True,
    'tooltip_x_position': 2,
    'tooltip_y_position': -25,
    'tooltip_display_time_ms': 1500,
    'tooltip_alpha': 0.86,
    'ui_always_on_top': False,
    'show_mouse_position_ui': False,
    'apply_transparency_to_new_windows': False,
    'new_window_transparency_level': 86,
    'global_transparency_exclusions': 'dsclock, explorer, WorkerW, SideBar_HTMLHostWindow, Sidebar, kv_ds_digitclock_32', # REMOVED ElectricsheepWndClass
    'dynamic_transparency_enabled': False,
    'active_window_transparency': 86,
    'inactive_window_transparency': 64,
    'manage_all_windows_dynamically': False,
    'inactive_window_auto_update': False,    # RESTORED
    'window_monitor_interval_ms': 200,
    'new_window_check_interval_ms': 2000,
    'center_on_first_launch': True,
    'prevent_window_edges_off_screen': False,
    'focus_mode_active': False,
    'focus_tooltip_x_position': -70,
    'focus_tooltip_y_position': -20,
    'focus_mode_alt_tab_delay_ms': 1600,
    'minimize_inactive_windows': True,
    'minimize_inactive_delay_ms': 15000,
    'minimize_inactive_ignore_count': 3,
    'apply_on_script_start': True,
    'center_electricsheep_special': True,
    'enable_hotkey_passthrough': False, # NEW: Setting for Electricsheep crash protection
}

SETTINGS_FILE = 'transparency_settings.pkl'

class TransparencyControllerApp:
    _CUSTOM_KEY_DISPLAY_ORDER = [
        'None',
        'Backspace', 'Tab', 'Enter', 'Esc', 'Space', 'Page Up', 'Page Down', 'End', 'Home',
        'Left', 'Up', 'Right', 'Down', 'Ins', 'Del',
        '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
        'Numpad 0', 'Numpad 1', 'Numpad 2', 'Numpad 3', 'Numpad 4', 'Numpad 5', 'Numpad 6', 'Numpad 7', 'Numpad 8', 'Numpad 9',
        'Num *', 'Num +', 'Num -', 'Num .', 'Num /',
        'F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11', 'F12',
        'Mouse Left Click', 'Mouse Right Click', 'Mouse Middle Click', 'Mouse XButton1', 'Mouse XButton2', 'Mouse Wheel Up', 'Mouse Wheel Down',
        '~', '-', '=', '[', ']', ';', "'", ',', '.', '\\', '/'
    ]
    
    _CHROMA_KEY_COLOR_HEX = "#00FF00"

    def __init__(self, root):
        self.root = root
        
        # Fix for clicking out of variable boxes
        self.root.bind_all("<Button-1>", self._on_click_anywhere)
        
        self.load_settings()
        self.apply_theme_settings()

        self.root.title("Transparency Controller")
        self.root.geometry(INITIAL_WINDOW_SIZE)
        self.root.resizable(True, True)
        self.root.attributes('-topmost', self.settings['ui_always_on_top'])
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.current_transparency_level = self.settings['transparency_levels']['initial']
        self.script_enabled = self.settings['script_enabled']
        self.focus_mode_active = self.settings['focus_mode_active']
        self.last_scroll_time = 0
        self.last_processed_hwnd = None
        self.tooltip_timer = None
        self.hotkey_capture_active = False
        self.changer_window = None # Ensure changer_window is initialized to None early
        self.setting_entries = {}
        self.exclusion_list_entries = {}

        self.mouse_pos_timer = None

        self.tooltip_following = False
        self.tooltip_follow_timer = None

        # NEW: Brightness control state
        self.current_brightness_level = self.settings['brightness_levels']['initial']
        self.is_brightness_scrolling = False
        self.last_brightness_scroll_time = 0
        self.last_brightness_hotkey_press_time = 0 # NEW: For tracking hotkey presses for reset_on_scroll_start
        self._sbc_available = False # Flag to check if screen_brightness_control is available

        # NEW: Transparency scrolling state
        self.is_transparency_scrolling = False # NEW
        self.last_transparency_hotkey_press_time = 0 # NEW

        self.last_foreground_hwnd = None
        self.processed_new_windows = set()
        self.managed_by_script_hwnds = set()
        self.minimized_by_script_hwnds = set()
        self.initial_script_start_hwnds = set()
        self.window_last_active_time = {}

        self.window_monitor_fg_timer = None
        self.window_monitor_new_timer = None
        self.window_monitor_inactivity_timer = None

        self.ahk = ahk.AHK(executable_path='C:\\Program Files\\AutoHotkey\\v2\\AutoHotkey.exe')

        self._initialize_hotkey_maps()

        # --- FIX: Create tooltip window BEFORE populating initial HWNDs ---
        self.setup_tooltip_window() 
        # --- END FIX ---

        self._populate_initial_script_hwnds()

        # Check for screen_brightness_control availability on Windows
        if platform.system() == "Windows":
            self._sbc_available = importlib.util.find_spec("screen_brightness_control") is not None
            if not self._sbc_available:
                self.show_message("Warning: 'screen-brightness-control' library not found. Brightness control will be unavailable. Please install it manually: pip install screen-brightness-control pywin32", "orange")

        self.scrollable_frame = customtkinter.CTkScrollableFrame(self.root)
        self.scrollable_frame.pack(pady=10, padx=10, fill="both", expand=True)
        self.create_widgets(self.scrollable_frame) # create_widgets will now use the existing tooltip_window

        self.ahk.start_hotkeys()

        self.register_hotkeys()
        self.update_status_label()

        if self.settings['show_mouse_position_ui']:
            self.update_mouse_position_label()

        self._start_window_monitoring()

        # NEW: Apply settings on script start if enabled
        if self.settings['apply_on_script_start']:
            self.show_message("Applying initial settings based on 'Apply on Script Start'.", "blue")
            # Populate managed_by_script_hwnds with all initial non-excluded windows if manage_all is ON
            # or if dynamic transparency is enabled and allowed for new windows (which includes initial ones for this purpose)
            if self.settings['manage_all_windows_dynamically'] or self.settings['dynamic_transparency_enabled']:
                for hwnd in self.initial_script_start_hwnds:
                    if not self._is_window_excluded(hwnd):
                        self.managed_by_script_hwnds.add(hwnd)

            # Reapply dynamic transparency to all relevant windows (initial ones)
            if self.settings['dynamic_transparency_enabled']:
                self._reapply_dynamic_transparency_on_all_windows(force_all=self.settings['manage_all_windows_dynamically'])
            
            # Apply centering to initial windows if enabled
            if self.settings['center_on_first_launch']:
                for hwnd in self.initial_script_start_hwnds:
                    if not self._is_window_excluded(hwnd) and hwnd not in self.managed_by_script_hwnds: # Only center if not already managed/processed
                        self._center_window(hwnd, show_tooltip=False)

        # Initialize window_last_active_time for all currently open windows
        current_time_ms = time.time() * 1000
        for hwnd in self.initial_script_start_hwnds:
            self.window_last_active_time[hwnd] = current_time_ms
        # Also for the foreground window
        fg_hwnd = win32gui.GetForegroundWindow()
        if fg_hwnd:
            self.window_last_active_time[fg_hwnd] = current_time_ms

    def _on_click_anywhere(self, event):
        """Handles click events globally to clear focus from entry widgets."""
        # If the clicked widget is NOT an entry, focus the root window (clearing entry focus)
        if not isinstance(event.widget, (customtkinter.CTkEntry, tk.Entry)):
            self.root.focus_set()

    def load_settings(self):
        """Loads settings from file, or uses defaults if file is not found/corrupt."""
        try:
            with open(SETTINGS_FILE, 'rb') as f:
                self.settings = pickle.load(f)
            def merge_dicts(source, destination):
                for key, value in source.items():
                    if key not in destination:
                        destination[key] = value
                    elif isinstance(value, dict) and isinstance(destination[key], dict):
                        merge_dicts(value, destination[key])
            merge_dicts(DEFAULT_SETTINGS, self.settings)

            # Remove deprecated settings
            if 'hotkey_capture_settings' in self.settings:
                del self.settings['hotkey_capture_settings']
            if 'ui_topmost_checkbox_x' in self.settings:
                del self.settings['ui_topmost_checkbox_x']
            if 'show_mouse_pos_checkbox_x' in self.settings:
                del self.settings['show_mouse_pos_checkbox_x']
            if 'new_window_transparency_exclusions' in self.settings:
                del self.settings['new_window_transparency_exclusions']
            if 'dynamic_transparency_exclusions' in self.settings:
                del self.settings['dynamic_transparency_exclusions']
            if 'manage_existing_windows_dynamically' in self.settings:
                del self.settings['manage_existing_windows_dynamically']
            # REMOVED DELETION OF THESE KEYS as they are being restored
            # if 'inactive_window_auto_update' in self.settings:
            #     del self.settings['inactive_window_auto_update']
            if 'brightness_levels' in self.settings and 'scroll_stop_delay_ms' in self.settings['brightness_levels']:
                del self.settings['brightness_levels']['scroll_stop_delay_ms']

        except (FileNotFoundError, EOFError, pickle.UnpickingError):
            self.settings = DEFAULT_SETTINGS.copy()
        self.settings['hotkeys'] = DEFAULT_SETTINGS['hotkeys'].copy() # Reset hotkeys for test
        self.original_hotkeys = self.settings['hotkeys'].copy()

    def save_settings(self):
        """Saves current settings to file."""
        with open(SETTINGS_FILE, 'wb') as f:
            pickle.dump(self.settings, f)

    def apply_theme_settings(self):
        """Applies CustomTkinter theme and appearance mode based on settings."""
        customtkinter.set_appearance_mode(self.settings['appearance_mode'])
        customtkinter.set_default_color_theme(self.settings['theme_color'])

    def create_widgets(self, parent_frame):
        """Creates and places all GUI elements within the given parent frame."""
        # CHANGED: Removed fill="x" to allow frame to shrink-wrap (centering the block)
        appearance_frame = customtkinter.CTkFrame(parent_frame)
        appearance_frame.pack(pady=10, padx=10, anchor="center") 

        customtkinter.CTkLabel(appearance_frame, text="Appearance Settings", font=customtkinter.CTkFont(weight="bold")).pack(pady=5, anchor="center")

        # --- Theme Color and Appearance Mode on one line ---
        theme_mode_row_frame = customtkinter.CTkFrame(appearance_frame, fg_color="transparent")
        theme_mode_row_frame.pack(pady=2, anchor="center")

        theme_col_frame = customtkinter.CTkFrame(theme_mode_row_frame, fg_color="transparent")
        theme_col_frame.pack(side="left", padx=10, pady=2, expand=True)
        customtkinter.CTkLabel(theme_col_frame, text="Theme Color:").pack(pady=(0,2), anchor="center")
        self.theme_menu_var = customtkinter.StringVar(value=self.settings['theme_color'])
        self.theme_menu = customtkinter.CTkOptionMenu(theme_col_frame, values=['green', 'blue', 'dark-blue'],
                                                 variable=self.theme_menu_var,
                                                 command=self.change_theme_color)
        self.theme_menu.pack(pady=(2,0), anchor="center")

        appearance_mode_col_frame = customtkinter.CTkFrame(theme_mode_row_frame, fg_color="transparent")
        appearance_mode_col_frame.pack(side="left", padx=10, pady=2, expand=True)
        customtkinter.CTkLabel(appearance_mode_col_frame, text="Mode:").pack(pady=(0,2), anchor="center")
        self.appearance_menu_var = customtkinter.StringVar(value=self.settings['appearance_mode'])
        self.appearance_menu = customtkinter.CTkOptionMenu(appearance_mode_col_frame, values=['System', 'Dark', 'Light'],
                                                      variable=self.appearance_menu_var,
                                                      command=self.change_appearance_mode)
        self.appearance_menu.pack(pady=(2,0), anchor="center")
        # --- End Theme Color and Appearance Mode ---

        # --- Always On Top and Mouse Position on one line ---
        checkbox_row_frame = customtkinter.CTkFrame(appearance_frame, fg_color="transparent")
        checkbox_row_frame.pack(pady=5, anchor="center", fill="x")

        # Frame for "Always On Top" checkbox
        ui_topmost_col_frame = customtkinter.CTkFrame(checkbox_row_frame, fg_color="transparent")
        ui_topmost_col_frame.pack(side="left", padx=10, expand=True, fill="x")
        self.ui_topmost_checkbox = customtkinter.CTkCheckBox(ui_topmost_col_frame,
                                                             text="Always On Top",
                                                             command=self.toggle_ui_topmost)
        self.ui_topmost_checkbox.pack(anchor="center") # Centered within its column frame
        if self.settings['ui_always_on_top']:
            self.ui_topmost_checkbox.select()
        else:
            self.ui_topmost_checkbox.deselect()

        # Frame for "Mouse Position" checkbox
        show_mouse_pos_col_frame = customtkinter.CTkFrame(checkbox_row_frame, fg_color="transparent")
        show_mouse_pos_col_frame.pack(side="left", padx=10, expand=True, fill="x")
        self.show_mouse_pos_checkbox = customtkinter.CTkCheckBox(show_mouse_pos_col_frame,
                                                                 text="Mouse Position",
                                                                 command=self.toggle_show_mouse_position_ui)
        self.show_mouse_pos_checkbox.pack(anchor="center") # Centered within its column frame
        if self.settings['show_mouse_position_ui']:
            self.show_mouse_pos_checkbox.select()
        else:
            self.show_mouse_pos_checkbox.deselect()
        # --- End Always On Top and Mouse Position ---

        self.restart_warning_label = customtkinter.CTkLabel(appearance_frame,
                                                            text="Warning: Restart needed for full theme changes. Click to restart.",
                                                            text_color="yellow",
                                                            cursor="hand2",
                                                            wraplength=300)
        self.restart_warning_label.pack(pady=5, anchor="center")
        self.restart_warning_label.pack_forget()
        self.restart_warning_label.bind("<Button-1>", lambda event: self.restart_app())

        # CHANGED: Removed fill="x"
        hotkey_frame = customtkinter.CTkFrame(parent_frame)
        hotkey_frame.pack(pady=10, padx=10, anchor="center") 
        customtkinter.CTkLabel(hotkey_frame, text="Hotkey Configuration", font=customtkinter.CTkFont(weight="bold")).pack(pady=5, anchor="center")

        self.hotkey_labels = {}
        for action, hotkey in self.settings['hotkeys'].items():
            row_frame = customtkinter.CTkFrame(hotkey_frame, fg_color="transparent")
            # CHANGED: anchor="w" to align rows left within the centered block
            row_frame.pack(pady=2, anchor="w")

            customtkinter.CTkLabel(row_frame, text=f"{action.replace('_', ' ').title()}:", width=150, anchor="w").pack(side="left", padx=5)

            display_text = self._get_hotkey_display_text(action, hotkey)
            self.hotkey_labels[action] = customtkinter.CTkLabel(row_frame, text=display_text, width=120, anchor="w")
            self.hotkey_labels[action].pack(side="left", padx=5)

            change_button = customtkinter.CTkButton(row_frame, text="Change", width=60,
                                                    command=lambda a=action: self.open_manual_hotkey_changer(a))
            change_button.pack(side="right", padx=5)

        # CHANGED: Removed fill="x"
        transparency_settings_frame = customtkinter.CTkFrame(parent_frame)
        transparency_settings_frame.pack(pady=10, padx=10, anchor="center")
        customtkinter.CTkLabel(transparency_settings_frame, text="Transparency Settings", font=customtkinter.CTkFont(weight="bold")).pack(pady=5, anchor="center")

        self.create_setting_entry(transparency_settings_frame, "Initial Level (%):", 'transparency_levels', 'initial')
        self.create_setting_entry(transparency_settings_frame, "Min Level (%):", 'transparency_levels', 'min')
        self.create_setting_entry(transparency_settings_frame, "Max Level (%):", 'transparency_levels', 'max')
        self.create_setting_entry(transparency_settings_frame, "Slow Scroll Inc:", 'transparency_levels', 'scroll_increment_slow')
        self.create_setting_entry(transparency_settings_frame, "Fast Scroll Inc:", 'transparency_levels', 'scroll_increment_fast')
        self.create_setting_entry(transparency_settings_frame, "Fast Scroll Threshold (ms):", 'transparency_levels', 'fast_scroll_threshold_ms')
        self.create_setting_entry(transparency_settings_frame, "XButton2 Preset (%):", 'transparency_levels', 'preset_xbutton2')
        self.create_setting_entry(transparency_settings_frame, "Ctrl+Shift+XButton2 Preset (%):", 'transparency_levels', 'preset_xbutton2_shift')
        self.create_setting_entry(transparency_settings_frame, "XButton1 Preset (%):", 'transparency_levels', 'preset_xbutton1')
        
        # NEW: Reset Transparency on Scroll Start checkbox
        self.transparency_reset_on_scroll_checkbox = customtkinter.CTkCheckBox(transparency_settings_frame,
                                                                          text="Reset Transparency on Scroll Start",
                                                                          command=lambda: self._update_setting_from_checkbox('reset_on_scroll_start', self.transparency_reset_on_scroll_checkbox, category='transparency_levels'))
        self.transparency_reset_on_scroll_checkbox.pack(pady=5, anchor="w", padx=10)
        if self.settings['transparency_levels']['reset_on_scroll_start']:
            self.transparency_reset_on_scroll_checkbox.select()
        else:
            self.transparency_reset_on_scroll_checkbox.deselect()

        # CHANGED: Removed fill="x"
        tooltip_settings_frame = customtkinter.CTkFrame(parent_frame)
        tooltip_settings_frame.pack(pady=10, padx=10, anchor="center")
        customtkinter.CTkLabel(tooltip_settings_frame, text="Tooltip Settings", font=customtkinter.CTkFont(weight="bold")).pack(pady=5, anchor="center")
        self.create_setting_entry(tooltip_settings_frame, "Tooltip X Offset:", 'tooltip_x_position', None, is_top_level=True)
        self.create_setting_entry(tooltip_settings_frame, "Tooltip Y Offset:", 'tooltip_y_position', None, is_top_level=True)
        self.create_setting_entry(tooltip_settings_frame, "Tooltip Alpha (0.0-1.0):", 'tooltip_alpha', None, is_top_level=True, value_type=float, increment=0.01)

        # CHANGED: Removed fill="x"
        brightness_settings_frame = customtkinter.CTkFrame(parent_frame)
        brightness_settings_frame.pack(pady=10, padx=10, anchor="center")
        customtkinter.CTkLabel(brightness_settings_frame, text="Brightness Settings", font=customtkinter.CTkFont(weight="bold")).pack(pady=5, anchor="center")

        self.create_setting_entry(brightness_settings_frame, "Initial Level (%):", 'brightness_levels', 'initial')
        self.create_setting_entry(brightness_settings_frame, "Min Level (%):", 'brightness_levels', 'min')
        self.create_setting_entry(brightness_settings_frame, "Max Level (%):", 'brightness_levels', 'max')
        self.create_setting_entry(brightness_settings_frame, "Slow Scroll Inc:", 'brightness_levels', 'scroll_increment_slow')
        self.create_setting_entry(brightness_settings_frame, "Fast Scroll Inc:", 'brightness_levels', 'scroll_increment_fast')
        self.create_setting_entry(brightness_settings_frame, "Fast Scroll Threshold (ms):", 'brightness_levels', 'fast_scroll_threshold_ms')
        self.create_setting_entry(brightness_settings_frame, "XButton2 Preset (%):", 'brightness_levels', 'preset_xbutton2')
        self.create_setting_entry(brightness_settings_frame, "XButton1 Preset (%):", 'brightness_levels', 'preset_xbutton1')

        self.brightness_reset_on_scroll_checkbox = customtkinter.CTkCheckBox(brightness_settings_frame,
                                                                          text="Reset Brightness on Scroll Start",
                                                                          command=lambda: self._update_setting_from_checkbox('reset_on_scroll_start', self.brightness_reset_on_scroll_checkbox, category='brightness_levels'))
        self.brightness_reset_on_scroll_checkbox.pack(pady=5, anchor="w", padx=10)
        if self.settings['brightness_levels']['reset_on_scroll_start']:
            self.brightness_reset_on_scroll_checkbox.select()
        else:
            self.brightness_reset_on_scroll_checkbox.deselect()

        # CHANGED: Removed fill="x"
        advanced_transparency_frame = customtkinter.CTkFrame(parent_frame)
        advanced_transparency_frame.pack(pady=10, padx=10, anchor="center")
        customtkinter.CTkLabel(advanced_transparency_frame, text="Advanced Window Management", font=customtkinter.CTkFont(weight="bold")).pack(pady=5, anchor="center")

        self.new_window_transparency_checkbox = customtkinter.CTkCheckBox(advanced_transparency_frame,
                                                                          text="Apply transparency to new windows (once)",
                                                                          command=self.toggle_apply_transparency_to_new_windows)
        self.new_window_transparency_checkbox.pack(pady=5, anchor="w", padx=10)
        if self.settings['apply_transparency_to_new_windows']:
            self.new_window_transparency_checkbox.select()
        else:
            self.new_window_transparency_checkbox.deselect()

        self.create_setting_entry(advanced_transparency_frame, "New Window Level (%):", 'new_window_transparency_level', None, is_top_level=True)

        self.dynamic_transparency_checkbox = customtkinter.CTkCheckBox(advanced_transparency_frame,
                                                                       text="Dynamic Transparency Active/Inactive Manual Update",
                                                                       command=self.toggle_dynamic_transparency)
        self.dynamic_transparency_checkbox.pack(pady=5, anchor="w", padx=10)
        if self.settings['dynamic_transparency_enabled']:
            self.dynamic_transparency_checkbox.select()
        else:
            self.dynamic_transparency_checkbox.deselect()

        # New: Controls for granular behavior when "Manage ALL Windows" is OFF

        self.inactive_window_auto_update_checkbox = customtkinter.CTkCheckBox(advanced_transparency_frame,
                                                                                                  text="Inactive Window Auto Update",
                                                                                                  command=lambda: self._update_setting_from_checkbox('inactive_window_auto_update', self.inactive_window_auto_update_checkbox))
        self.inactive_window_auto_update_checkbox.pack(pady=5, anchor="w", padx=10)
        if self.settings['inactive_window_auto_update']:
            self.inactive_window_auto_update_checkbox.select()
        else:
            self.inactive_window_auto_update_checkbox.deselect()

        self.create_setting_entry(advanced_transparency_frame, "Active Window Level (%):", 'active_window_transparency', None, is_top_level=True)
        self.create_setting_entry(advanced_transparency_frame, "Inactive Window Level (%):", 'inactive_window_transparency', None, is_top_level=True)

        # NEW: Manage ALL Windows Switch
        self.manage_all_windows_dynamically_switch = customtkinter.CTkSwitch(advanced_transparency_frame,
                                                                          text="Manage ALL Windows (Existing & New)",
                                                                          command=self.toggle_manage_all_windows_dynamically,
                                                                          onvalue=True, offvalue=False)
        self.manage_all_windows_dynamically_switch.pack(pady=5, anchor="w", padx=10)
        if self.settings['manage_all_windows_dynamically']:
            self.manage_all_windows_dynamically_switch.select()
        else:
            self.manage_all_windows_dynamically_switch.deselect()

        # New: Center on First Launch Checkbox
        self.center_on_first_launch_checkbox = customtkinter.CTkCheckBox(advanced_transparency_frame,
                                                                         text="Center new windows (once)", # Changed text
                                                                         command=self.toggle_center_on_first_launch)
        self.center_on_first_launch_checkbox.pack(pady=5, anchor="w", padx=10)
        if self.settings['center_on_first_launch']:
            self.center_on_first_launch_checkbox.select()
        else:
            self.center_on_first_launch_checkbox.deselect()

        # New: Prevent Window Edges Off Screen Checkbox
        self.prevent_edges_off_screen_checkbox = customtkinter.CTkCheckBox(advanced_transparency_frame,
                                                                           text="Prevent centered window edges off screen",
                                                                           command=self.toggle_prevent_window_edges_off_screen)
        self.prevent_edges_off_screen_checkbox.pack(pady=5, anchor="w", padx=10)
        if self.settings['prevent_window_edges_off_screen']:
            self.prevent_edges_off_screen_checkbox.select()
        else:
            self.prevent_edges_off_screen_checkbox.deselect()

        # NEW: Electricsheep Special Centering Checkbox
        self.center_electricsheep_special_checkbox = customtkinter.CTkCheckBox(advanced_transparency_frame,
                                                                               text="Center Electricsheep (es.exe) with title bar hidden",
                                                                               command=self.toggle_center_electricsheep_special)
        self.center_electricsheep_special_checkbox.pack(pady=5, anchor="w", padx=10)
        if self.settings['center_electricsheep_special']:
            self.center_electricsheep_special_checkbox.select()
        else:
            self.center_electricsheep_special_checkbox.deselect()

        # NEW: Electricsheep Crash Protection Checkbox
        self.enable_hotkey_passthrough_checkbox = customtkinter.CTkCheckBox(advanced_transparency_frame,
                                                                                 text="""Enable Hotkey Passthrough Non-Suppressing Exception
Fixes Electricsheep Hotkey Compatibility """,
                                                                                 command=self.toggle_enable_hotkey_passthrough)
        self.enable_hotkey_passthrough_checkbox.pack(pady=5, anchor="w", padx=10)
        if self.settings['enable_hotkey_passthrough']:
            self.enable_hotkey_passthrough_checkbox.select()
        else:
            self.enable_hotkey_passthrough_checkbox.deselect()

        # NEW: Minimize Inactive Windows Checkbox
        self.minimize_inactive_windows_checkbox = customtkinter.CTkCheckBox(advanced_transparency_frame,
                                                                          text="Minimize inactive windows",
                                                                          command=self.toggle_minimize_inactive_windows)
        self.minimize_inactive_windows_checkbox.pack(pady=5, anchor="w", padx=10)
        if self.settings['minimize_inactive_windows']:
            self.minimize_inactive_windows_checkbox.select()
        else:
            self.minimize_inactive_windows_checkbox.deselect()

        # NEW: Minimize Inactive Windows settings entries
        self.create_setting_entry(advanced_transparency_frame, "Minimize after inactive for (ms):", 'minimize_inactive_delay_ms', None, is_top_level=True)
        self.create_setting_entry(advanced_transparency_frame, "Ignore X oldest inactive windows:", 'minimize_inactive_ignore_count', None, is_top_level=True)

        # NEW: Apply on Script Start checkbox
        self.apply_on_script_start_checkbox = customtkinter.CTkCheckBox(advanced_transparency_frame,
                                                                        text="Apply settings on script start",
                                                                        command=lambda: self._update_setting_from_checkbox('apply_on_script_start', self.apply_on_script_start_checkbox))
        self.apply_on_script_start_checkbox.pack(pady=5, anchor="w", padx=10)
        if self.settings['apply_on_script_start']:
            self.apply_on_script_start_checkbox.select()
        else:
            self.apply_on_script_start_checkbox.deselect()


        # New: Focus Mode Settings Frame
        # CHANGED: Removed fill="x"
        focus_mode_frame = customtkinter.CTkFrame(parent_frame)
        focus_mode_frame.pack(pady=10, padx=10, anchor="center")
        customtkinter.CTkLabel(focus_mode_frame, text="Focus Mode Settings", font=customtkinter.CTkFont(weight="bold")).pack(pady=5, anchor="center")

        self.focus_mode_checkbox = customtkinter.CTkCheckBox(focus_mode_frame,
                                                             text="Enable Focus Mode",
                                                             command=self.toggle_focus_mode_ui)
        self.focus_mode_checkbox.pack(pady=5, anchor="w", padx=10)
        if self.settings['focus_mode_active']:
            self.focus_mode_checkbox.select()
        else:
            self.focus_mode_checkbox.deselect()

        self.create_setting_entry(focus_mode_frame, "Focus Tooltip X Offset:", 'focus_tooltip_x_position', None, is_top_level=True)
        self.create_setting_entry(focus_mode_frame, "Focus Tooltip Y Offset:", 'focus_tooltip_y_position', None, is_top_level=True)
        self.create_setting_entry(focus_mode_frame, "Alt+Tab Delay (ms):", 'focus_mode_alt_tab_delay_ms', None, is_top_level=True)

        self.create_exclusion_list_entry(advanced_transparency_frame, "Global Exclusions (exe/class,exe/class):", 'global_transparency_exclusions')

        # CHANGED: Removed fill="x"
        control_frame = customtkinter.CTkFrame(parent_frame)
        control_frame.pack(pady=10, padx=10, anchor="center")
        self.status_label = customtkinter.CTkLabel(control_frame, text="Status: Initializing...")
        self.status_label.pack(pady=5, anchor="center")

        self.toggle_button = customtkinter.CTkButton(control_frame, text="Toggle Script (Alt+W)", command=self.toggle_script_ui)
        self.toggle_button.pack(pady=5, anchor="center")

        # NEW: Restore All Managed Windows to 100% button
        restore_button = customtkinter.CTkButton(control_frame, text="Restore All Managed Windows to 100%", command=self.restore_all_managed_to_full_opacity)
        restore_button.pack(pady=5, anchor="center")

        reset_button = customtkinter.CTkButton(control_frame, text="Reset to Defaults", command=self.reset_to_defaults)
        reset_button.pack(pady=5, anchor="center")

        self.mouse_pos_label = customtkinter.CTkLabel(self.root, text="")
        self.mouse_pos_label.pack(side="bottom", anchor="s", padx=10, pady=5)

    def create_setting_entry(self, parent_frame, label_text, category, key, is_top_level=False, value_type=int, increment=1):
        """Helper to create a label, entry, and apply button for a setting, with scroll/arrow key support."""
        frame = customtkinter.CTkFrame(parent_frame, fg_color="transparent")
        # CHANGED: anchor="w" to align rows left within the centered block
        frame.pack(pady=2, anchor="w")

        customtkinter.CTkLabel(frame, text=label_text, width=200, anchor="w").pack(side="left")

        if is_top_level:
            initial_value = self.settings[category]
        else:
            initial_value = self.settings[category][key]

        entry = customtkinter.CTkEntry(frame, width=80)
        if value_type == float:
            entry.insert(0, f"{initial_value:.2f}")
        else:
            entry.insert(0, str(initial_value))
        entry.pack(side="left", padx=5)

        if is_top_level:
            self.setting_entries[(category,)] = (entry, value_type, increment)
        else:
            self.setting_entries[(category, key)] = (entry, value_type, increment)

        entry.bind("<MouseWheel>", lambda event, e=entry, c=category, k=key, i_tl=is_top_level, vt=value_type, inc=increment: self._on_entry_scroll(event, e, c, k, i_tl, vt, inc))
        entry.bind("<Up>", lambda event, e=entry, c=category, k=key, i_tl=is_top_level, vt=value_type, inc=increment: self._on_entry_arrow_key(event, e, c, k, i_tl, 1, vt, inc))
        entry.bind("<Down>", lambda event, e=entry, c=category, k=key, i_tl=is_top_level, vt=value_type, inc=increment: self._on_entry_arrow_key(event, e, c, k, i_tl, -1, vt, inc))

        save_button = customtkinter.CTkButton(frame, text="Apply", width=60,
                                              command=lambda: self.apply_setting(entry, category, key, is_top_level, value_type))
        save_button.pack(side="left", padx=5)

    def create_exclusion_list_entry(self, parent_frame, label_text, setting_key):
        """Helper to create a label, entry for an exclusion list, and apply button."""
        # Separate label onto its own line
        # CHANGED: anchor="w" to align label left within the centered block
        customtkinter.CTkLabel(parent_frame, text=label_text, anchor="w").pack(pady=(5, 2), padx=5, anchor="w")

        frame = customtkinter.CTkFrame(parent_frame, fg_color="transparent")
        # CHANGED: anchor="w" to align entry left within the centered block
        frame.pack(pady=2, anchor="w", fill="x")

        initial_value = self.settings[setting_key]
        entry = customtkinter.CTkEntry(frame)
        entry.insert(0, str(initial_value))
        entry.pack(side="left", padx=5, fill="x", expand=True)

        self.exclusion_list_entries[setting_key] = entry

        save_button = customtkinter.CTkButton(frame, text="Apply", width=60,
                                              command=lambda: self.apply_exclusion_list_setting(entry, setting_key))
        save_button.pack(side="left", padx=5)

    def apply_exclusion_list_setting(self, entry_widget, setting_key):
        """Applies an exclusion list setting from an entry widget."""
        try:
            value_str = entry_widget.get()
            cleaned_list = [item.strip().lower() for item in value_str.split(',') if item.strip()]
            self.settings[setting_key] = ", ".join(cleaned_list)
            self.save_settings()
            self.show_message(f"Applied {setting_key.replace('_', ' ').title()}: {self.settings[setting_key]}", "green")
            
            # NEW: Trigger re-evaluation for all managed windows and reset inactivity tracking
            # Crucially, we need to make sure that if an item is removed from the list, it is no longer considered excluded immediately.
            self._reapply_dynamic_transparency_on_all_windows(force_all=True) # Forces re-evaluation of all windows
            self._reset_inactivity_tracking_state() # Re-evaluates for minimization based on new exclusions
            
        except Exception as e:
            self.show_message(f"Error applying exclusion list setting: {e}", "red")

    def toggle_apply_transparency_to_new_windows(self):
        """Toggles the 'apply_transparency_to_new_windows' setting."""
        new_state = self.new_window_transparency_checkbox.get() == 1
        self.settings['apply_transparency_to_new_windows'] = new_state
        self.save_settings()
        self.show_message(f"'Apply transparency to new windows' set to: {new_state}", "blue")

    def toggle_dynamic_transparency(self):
        """Toggles the 'dynamic_transparency_enabled' setting and applies changes."""
        new_state = self.dynamic_transparency_checkbox.get() == 1
        self.settings['dynamic_transparency_enabled'] = new_state
        self.save_settings()
        self.show_message(f"'Dynamic transparency for active/inactive windows' set to: {new_state}", "blue")

        # Reapply dynamic transparency to current windows based on the new state.
        # The _reapply_dynamic_transparency_on_all_windows function will now correctly
        # handle whether to apply transparency or just remove from managed set without restoring.
        self._reapply_dynamic_transparency_on_all_windows(force_all=True) # Force re-evaluation of all windows

    def toggle_manage_all_windows_dynamically(self):
        """
        Toggles the 'manage_all_windows_dynamically' setting (the switch).
        When enabled, the script will manage transparency for ALL open windows (if not excluded).
        When disabled, it will only manage windows that open *after* the script started
        and were initially processed by 'apply_transparency_to_new_windows' (if respective sub-settings allow).
        """
        new_state = self.manage_all_windows_dynamically_switch.get() # CTkSwitch uses .get() directly
        self.settings['manage_all_windows_dynamically'] = new_state
        self.save_settings()
        self.show_message(f"'Manage ALL Windows' set to: {new_state}", "blue")

        if new_state: # If switch is ON
            if self.settings['dynamic_transparency_enabled']:
                self.show_message("Applying dynamic transparency to all currently open windows.", "blue")
                # Add all initial non-excluded windows to the managed set
                for hwnd in self.initial_script_start_hwnds:
                    if not self._is_window_excluded(hwnd):
                        self.managed_by_script_hwnds.add(hwnd)
                self._reapply_dynamic_transparency_on_all_windows(force_all=True)
            else:
                self.show_message("Dynamic transparency is off, 'Manage ALL Windows' switch has limited effect.", "orange")
                # Even if dynamic is off, if 'Manage ALL Windows' is ON, we might still want to track them.
                # But for transparency, no action is taken.
        else: # If switch is OFF
            self.show_message("Stopping automatic dynamic management for windows opened before script started.", "blue")
            # When manage_all is OFF, we don't automatically restore.
            # The _reapply_dynamic_transparency_on_all_windows will handle removing
            # windows from managed_by_script_hwnds that no longer meet criteria,
            # and restoring their transparency if they were managed by this specific setting.
            self._reapply_dynamic_transparency_on_all_windows(force_all=False)

    def toggle_center_on_first_launch(self):
        """Toggles the 'center_on_first_launch' setting."""
        new_state = self.center_on_first_launch_checkbox.get() == 1
        self.settings['center_on_first_launch'] = new_state
        self.save_settings()
        self.show_message(f"'Center new windows (once)' set to: {new_state}", "blue")

    def toggle_prevent_window_edges_off_screen(self):
        """Toggles the 'prevent_window_edges_off_screen' setting."""
        new_state = self.prevent_edges_off_screen_checkbox.get() == 1
        self.settings['prevent_window_edges_off_screen'] = new_state
        self.save_settings()
        self.show_message(f"'Prevent centered window edges off screen' set to: {new_state}", "blue")

    def toggle_center_electricsheep_special(self):
        """Toggles the 'center_electricsheep_special' setting."""
        new_state = self.center_electricsheep_special_checkbox.get() == 1
        self.settings['center_electricsheep_special'] = new_state
        self.save_settings()
        self.show_message(f"'Center Electricsheep (es.exe) with title bar hidden' set to: {new_state}", "blue")

    def toggle_enable_hotkey_passthrough(self):
        """Toggles the 'enable_hotkey_passthrough' setting and re-registers hotkeys."""
        new_state = self.enable_hotkey_passthrough_checkbox.get() == 1
        self.settings['enable_hotkey_passthrough'] = new_state
        self.save_settings()
        self.show_message(f"'Electricsheep Crash Protection' set to: {new_state}. Re-registering hotkeys.", "blue")
        self.register_hotkeys() # Re-register hotkeys to apply new suppression logic
        self._reset_inactivity_tracking_state() # Reset state for minimization exclusion

    def _populate_initial_script_hwnds(self):
        """Populates the set of HWNDs that exist when the script starts."""
        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != "":
                # Ensure our own UI windows are not added to initial_script_start_hwnds
                if hwnd != self.root.winfo_id() and \
                   hwnd != self.tooltip_window.winfo_id() and \
                   not (self.changer_window and hwnd == self.changer_window.winfo_id()):
                    self.initial_script_start_hwnds.add(hwnd)
            return True
        win32gui.EnumWindows(callback, None)
        # print(f"DEBUG: Initial script HWNDs: {len(self.initial_script_start_hwnds)}")

    def _check_foreground_window(self):
        """Periodically checks the foreground window and applies dynamic transparency."""
        if not self.script_enabled:
            self.window_monitor_fg_timer = self.root.after(self.settings['window_monitor_interval_ms'], self._check_foreground_window)
            return

        current_fg_hwnd = win32gui.GetForegroundWindow()

        # Update last active time for the current foreground window
        if current_fg_hwnd and win32gui.IsWindow(current_fg_hwnd):
            self.window_last_active_time[current_fg_hwnd] = time.time() * 1000

        if current_fg_hwnd == self.root.winfo_id() or \
           current_fg_hwnd == self.tooltip_window.winfo_id() or \
           (self.changer_window and current_fg_hwnd == self.changer_window.winfo_id()):
            if self.settings['dynamic_transparency_enabled'] and self.last_foreground_hwnd:
                self._apply_dynamic_transparency(current_fg_hwnd, self.last_foreground_hwnd)
            self.last_foreground_hwnd = current_fg_hwnd
            self.window_monitor_fg_timer = self.root.after(self.settings['window_monitor_interval_ms'], self._check_foreground_window)
            return

        if current_fg_hwnd != self.last_foreground_hwnd:
            if self.settings['dynamic_transparency_enabled']:
                self._apply_dynamic_transparency(current_fg_hwnd, self.last_foreground_hwnd)
            self.last_foreground_hwnd = current_fg_hwnd

        self.window_monitor_fg_timer = self.root.after(self.settings['window_monitor_interval_ms'], self._check_foreground_window)

    def _check_for_new_windows(self):
        """Periodically enumerates all windows to find and process newly opened ones."""
        # Only run if new window transparency or centering is enabled
        if not self.script_enabled and \
           not self.settings['apply_transparency_to_new_windows'] and \
           not self.settings['center_on_first_launch']:
            self.window_monitor_new_timer = self.root.after(self.settings['new_window_check_interval_ms'], self._check_for_new_windows)
            return

        current_visible_hwnds = set()
        def get_current_visible_hwnds_callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != "":
                # Ensure our own UI windows are not considered
                if hwnd != self.root.winfo_id() and \
                   hwnd != self.tooltip_window.winfo_id() and \
                   (not self.changer_window or hwnd != self.changer_window.winfo_id()):
                    current_visible_hwnds.add(hwnd)
            return True
        win32gui.EnumWindows(get_current_visible_hwnds_callback, None)

        # Identify closed windows and remove them from tracking sets
        closed_hwnds = self.processed_new_windows.difference(current_visible_hwnds)
        for hwnd in closed_hwnds:
            self.processed_new_windows.discard(hwnd)
            self.managed_by_script_hwnds.discard(hwnd)
            self.minimized_by_script_hwnds.discard(hwnd)
            if hwnd in self.window_last_active_time:
                del self.window_last_active_time[hwnd]
            # Note: We don't remove from initial_script_start_hwnds as that's a static list of windows present at script start.

        # Now, identify genuinely new windows (not in processed_new_windows)
        for hwnd in current_visible_hwnds:
            if hwnd not in self.processed_new_windows:
                self._process_newly_found_window(hwnd)

        self.window_monitor_new_timer = self.root.after(self.settings['new_window_check_interval_ms'], self._check_for_new_windows)

    def _process_newly_found_window(self, hwnd):
        """Applies transparency and/or centers a newly found window if enabled and not excluded."""
        # Mark as processed immediately to prevent re-processing by this specific check
        self.processed_new_windows.add(hwnd)

        # Ensure our own UI windows are not processed as new windows
        if hwnd == self.root.winfo_id() or \
           hwnd == self.tooltip_window.winfo_id() or \
           (self.changer_window and hwnd == self.changer_window.winfo_id()):
            return

        # If window is in the exclusion list, DO NOT ATTEMPT TO SET TRANSPARENCY OR CENTER.
        # Just ensure it's not in the managed set.
        if self._is_window_excluded(hwnd):
            if hwnd in self.managed_by_script_hwnds:
                self.managed_by_script_hwnds.discard(hwnd)
            return

        # Center on first launch (only if not excluded)
        if self.settings['center_on_first_launch']:
            # Only center if it's a truly new window not already managed by script
            if hwnd not in self.managed_by_script_hwnds:
                self._center_window(hwnd, show_tooltip=False) # No tooltip for auto-center

        # Apply transparency to new windows (only if not excluded)
        if self.settings['apply_transparency_to_new_windows']:
            # Only apply if it's a truly new window not already managed by script
            if hwnd not in self.managed_by_script_hwnds:
                target_level = self.settings['new_window_transparency_level']
                
                # If dynamic transparency is also enabled, and it's the foreground window,
                # apply the active level immediately. Otherwise, apply the new_window_transparency_level.
                # The dynamic transparency monitor will take over from here.
                if self.settings['dynamic_transparency_enabled'] and \
                   hwnd == win32gui.GetForegroundWindow():
                    target_level = self.settings['active_window_transparency']

                set_transparency_for_hwnd(hwnd, target_level)
                self.managed_by_script_hwnds.add(hwnd) # Add to managed set
                # self.show_message(f"Applied new window transparency ({target_level}%) to {get_window_exe_name(hwnd)}", "blue")

    def _set_screen_brightness(self, level):
        """
        Sets the screen brightness to the specified level (0-100).
        This method absorbs the logic from set_brightness.py.
        """
        if not 0 <= level <= 100:
            self.show_message("Error: Brightness level must be between 0 and 100.", "red")
            return

        os_name = platform.system()

        try:
            if os_name == "Windows":
                if not self._sbc_available:
                    self.show_message("Brightness control is unavailable because 'screen-brightness-control' is not installed.", "orange")
                    return
                try:
                    import screen_brightness_control as sbc
                    sbc.set_brightness(level)
                except Exception as e:
                    self.show_message(f"Error setting brightness on Windows: {e}", "red")
                    self.show_message("Ensure you have the necessary permissions and that 'screen-brightness-control' and 'pywin32' are correctly installed.", "red")

            elif os_name == "Darwin": # macOS
                macos_level = level / 100.0
                try:
                    script = f'tell application "System Events" to tell process "ControlCenter" to slider 1 of group 1 of group 1 of group 1 of UI element 1 of row 1 of outline 1 of scroll area 1 of group 1 of window "Control Center" to set value to {macos_level}'
                    subprocess.run(['osascript', '-e', script], check=True)
                except FileNotFoundError:
                    self.show_message("Error: 'osascript' command not found. This script requires macOS.", "red")
                except subprocess.CalledProcessError as e:
                    self.show_message(f"Error executing osascript: {e}. Could not set brightness. Ensure System Events has permission to control ControlCenter.", "red")
                except Exception as e:
                    self.show_message(f"Error setting brightness on macOS: {e}", "red")

            elif os_name == "Linux":
                linux_level = level / 100.0
                try:
                    output = subprocess.run(['xrandr'], capture_output=True, text=True, check=True)
                    lines = output.stdout.splitlines()
                    display = None
                    for line in lines:
                        if ' connected primary' in line:
                            display = line.split(' ')[0]
                            break

                    if display:
                        subprocess.run(['xrandr', '--output', display, '--brightness', str(linux_level)], check=True)
                    else:
                        self.show_message("Could not find primary display using xrandr.", "orange")

                except FileNotFoundError:
                    self.show_message("Error: 'xrandr' command not found. This script requires xrandr.", "red")
                except subprocess.CalledProcessError as e:
                    self.show_message(f"Error executing xrandr: {e}", "red")
                except Exception as e:
                    self.show_message(f"Error setting brightness on Linux: {e}", "red")

            else:
                self.show_message(f"Error: Unsupported operating system for brightness control: {os_name}", "red")

        except Exception as e:
            self.show_message(f"An unexpected error occurred during brightness control: {e}", "red")

    def _update_brightness_gui(self, new_level=None, delta=0):
        """
        Updates the screen brightness and shows a tooltip.
        This function is scheduled to run on the main GUI thread.
        """
        if not self.script_enabled:
            return

        current_brightness_config = self.settings['brightness_levels']

        # Determine new brightness level
        if new_level is not None:
            self.current_brightness_level = new_level
        elif delta != 0:
            current_time = time.time() * 1000
            # Calculate time difference using the dedicated scroll time
            time_diff = current_time - self.last_brightness_scroll_time
            self.last_brightness_scroll_time = current_time # Update *here* for fast/slow increment

            if time_diff < current_brightness_config['fast_scroll_threshold_ms'] and time_diff > 0:
                increment = current_brightness_config['scroll_increment_fast']
            else:
                increment = current_brightness_config['scroll_increment_slow']

            self.current_brightness_level += (increment * delta)

        self.current_brightness_level = max(current_brightness_config['min'],
                                            min(current_brightness_config['max'],
                                                self.current_brightness_level))

        # Apply brightness immediately
        self._set_screen_brightness(self.current_brightness_level)

        # Show tooltip immediately
        self.show_tooltip(f"Brightness: {self.current_brightness_level}%")

    def _ahk_brightness_callback(self, action):
        """
        Generic callback for AHK hotkeys that modify brightness.
        This function runs in the AHK hotkey thread, so it schedules GUI updates.
        """
        if not self.script_enabled:
            return

        hotkey_config_str = self.settings['hotkeys'][action]

        if not self.check_modifiers_match(hotkey_config_str):
            if DEBUG_PRINT_MODIFIER_STATE_ON_MOUSE_EVENT:
                print(f"DEBUG: Modifiers mismatch for {action} with hotkey '{hotkey_config_str}'. Current state: Ctrl={self.ahk.key_state('Ctrl')}, Shift={self.ahk.key_state('Shift')}, Alt={self.ahk.key_state('Alt')}, Win={self.ahk.key_state('LWin') or self.ahk.key_state('RWin')}")
            return

        current_brightness_config = self.settings['brightness_levels']

        # Check if this is the start of a new scrolling sequence AND if reset is enabled
        current_hotkey_time = time.time() * 1000
        time_since_last_hotkey = current_hotkey_time - self.last_brightness_hotkey_press_time
        
        # A longer timeout to detect end of scroll sequence
        if time_since_last_hotkey > SCROLL_SEQUENCE_TIMEOUT_MS:
            self.is_brightness_scrolling = False

        if not self.is_brightness_scrolling and current_brightness_config['reset_on_scroll_start']:
            self.current_brightness_level = current_brightness_config['initial']
            self.is_brightness_scrolling = True
        elif not self.is_brightness_scrolling and not current_brightness_config['reset_on_scroll_start']:
            self.is_brightness_scrolling = True
        
        self.last_brightness_hotkey_press_time = current_hotkey_time # Update last hotkey press time for next check

        if action == 'increase_brightness':
            self.root.after(0, lambda: self._update_brightness_gui(delta=1))
        elif action == 'decrease_brightness':
            self.root.after(0, lambda: self._update_brightness_gui(delta=-1))
        elif action == 'set_80_percent_brightness':
            self.root.after(0, lambda: self._update_brightness_gui(new_level=current_brightness_config['preset_xbutton2']))
            self.is_brightness_scrolling = False # Presets are not part of a scroll sequence
        elif action == 'set_0_percent_brightness':
            self.root.after(0, lambda: self._update_brightness_gui(new_level=current_brightness_config['preset_xbutton1']))
            self.is_brightness_scrolling = False # Presets are not part of a scroll sequence
        else:
            self.show_message(f"Unhandled AHK hotkey action for brightness: {action}", "orange")

    def _should_window_be_dynamically_managed(self, hwnd, is_foreground):
        """
        Determines if a given window should be actively managed for dynamic transparency
        based on current settings and its foreground status.
        If it should be managed, it's added to self.managed_by_script_hwnds.
        """
        if not self.settings['dynamic_transparency_enabled']:
            return False

        if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd) or not win32gui.GetWindowText(hwnd):
            # Invalid or invisible windows should not be managed
            if hwnd in self.managed_by_script_hwnds:
                self.managed_by_script_hwnds.discard(hwnd)
            return False

        if self._is_window_excluded(hwnd):
            # Excluded windows are never dynamically managed
            if hwnd in self.managed_by_script_hwnds:
                self.managed_by_script_hwnds.discard(hwnd)
            return False

        if self.settings['manage_all_windows_dynamically']:
            # If 'Manage ALL' is ON, all non-excluded, valid windows are managed.
            self.managed_by_script_hwnds.add(hwnd)
            return True
        else:
            # If 'Manage ALL' is OFF:
            # 1. If 'Inactive Window Manual Update' is ON, and this window just became foreground,
            #    it should be added to managed_by_script_hwnds.
            if is_foreground and self.settings['inactive_window_auto_update']:
                self.managed_by_script_hwnds.add(hwnd)
                return True
            # 2. Otherwise, it's only managed if it was ALREADY in managed_by_script_hwnds
            #    (e.g., from 'apply_transparency_to_new_windows' or hotkey action).
            return hwnd in self.managed_by_script_hwnds

    def _apply_dynamic_transparency(self, new_fg_hwnd, old_fg_hwnd):
        """Applies active/inactive transparency based on foreground window change."""
        # Handle minimization/restoration based on foreground window change (always run this)
        self._restore_minimized_windows_on_focus_change(new_fg_hwnd, old_fg_hwnd)

        # Only proceed with transparency logic if dynamic transparency is enabled
        if not self.settings['dynamic_transparency_enabled']:
            # If dynamic transparency is OFF, ensure any windows that were managed
            # and are now *not* supposed to be managed (e.g., manage_all was turned off)
            # are removed from the managed set. Do NOT restore transparency here.
            # The _reapply_dynamic_transparency_on_all_windows handles the cleanup.
            return

        # Process the new foreground window for transparency
        if self._should_window_be_dynamically_managed(new_fg_hwnd, is_foreground=True):
            target_level = self.settings['active_window_transparency']
            set_transparency_for_hwnd(new_fg_hwnd, target_level)
            # self.show_message(f"Set {get_window_exe_name(new_fg_hwnd)} to ACTIVE ({target_level}%)", "purple")
        elif new_fg_hwnd in self.managed_by_script_hwnds:
            # If it was managed but now _should_window_be_dynamically_managed returned False
            # (e.g., settings changed, or it's no longer foreground and not managed by other means)
            # Restore to 100% and remove from managed set.
            set_transparency_for_hwnd(new_fg_hwnd, 100)
            self.managed_by_script_hwnds.discard(new_fg_hwnd)


        # Process the old foreground window (now inactive) for transparency
        if old_fg_hwnd and old_fg_hwnd != new_fg_hwnd:
            if self._should_window_be_dynamically_managed(old_fg_hwnd, is_foreground=False):
                target_level = self.settings['inactive_window_transparency']
                set_transparency_for_hwnd(old_fg_hwnd, target_level)
                # self.show_message(f"Set {get_window_exe_name(old_fg_hwnd)} to INACTIVE ({target_level}%)", "purple")
            elif old_fg_hwnd in self.managed_by_script_hwnds:
                # If it was managed but now _should_window_be_dynamically_managed returned False
                # Restore to 100% and remove from managed set.
                set_transparency_for_hwnd(old_fg_hwnd, 100)
                self.managed_by_script_hwnds.discard(old_fg_hwnd)

    def _reapply_dynamic_transparency_on_all_windows(self, force_all=False):
        """
        Re-evaluates and applies dynamic transparency to windows.
        If force_all is True, it enumerates all visible windows and adds them to managed_by_script_hwnds
        (if not excluded). Otherwise, it only processes windows already in managed_by_script_hwnds.
        """
        current_fg_hwnd = win32gui.GetForegroundWindow()
        
        windows_to_check = set()
        if force_all or self.settings['manage_all_windows_dynamically'] or self.settings['inactive_window_auto_update']:
            # Enumerate all visible windows if 'manage_all' is ON, or if 'manual update' is ON (to catch potential new ones), or if forced.
            def enum_all_callback(hwnd, extra):
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != "":
                    # Skip our own UI
                    if hwnd != self.root.winfo_id() and \
                       hwnd != self.tooltip_window.winfo_id() and \
                       not (self.changer_window and hwnd == self.changer_window.winfo_id()):
                        windows_to_check.add(hwnd)
                return True
            win32gui.EnumWindows(enum_all_callback, None)
        
        # Also include any windows currently in managed_by_script_hwnds that might not be visible anymore
        # but we need to process for removal.
        windows_to_check.update(self.managed_by_script_hwnds)

        # This set will hold HWNDs that are actually managed for dynamic transparency in this cycle.
        current_cycle_dynamically_managed_hwnds = set()
        
        # Determine which windows should be managed dynamically in this cycle
        for hwnd in list(windows_to_check):
            # Check if it should be managed (this also updates self.managed_by_script_hwnds)
            if self._should_window_be_dynamically_managed(hwnd, is_foreground=(hwnd == current_fg_hwnd)):
                current_cycle_dynamically_managed_hwnds.add(hwnd)
            elif hwnd in self.managed_by_script_hwnds:
                # If it was managed but now _should_window_be_dynamically_managed returned False,
                # remove it from the managed set. Do NOT restore transparency here,
                # as per user request (unless dynamic_transparency_enabled is OFF, then
                # it should retain its last transparency).
                self.managed_by_script_hwnds.discard(hwnd)

        # Apply dynamic transparency to the determined set of windows
        if self.settings['dynamic_transparency_enabled']:
            for hwnd in current_cycle_dynamically_managed_hwnds:
                if hwnd == current_fg_hwnd:
                    set_transparency_for_hwnd(hwnd, self.settings['active_window_transparency'])
                else:
                    set_transparency_for_hwnd(hwnd, self.settings['inactive_window_transparency'])
        else:
            # If dynamic transparency is OFF, we should not apply any transparency here.
            # Windows should retain their last set transparency.
            pass

        # Final cleanup: Any windows that are still in `self.managed_by_script_hwnds` but
        # were NOT in `current_cycle_dynamically_managed_hwnds` (meaning they are no longer
        # considered managed by the current settings) should be removed from `self.managed_by_script_hwnds`.
        # Their transparency should *not* be restored to 100% here if dynamic is off,
        # unless `force_all` is true and it implies a full reset (e.g. from exclusion list change).
        
        # If dynamic_transparency_enabled is OFF, we just remove from tracking.
        # If it's ON, but a window is no longer managed, we restore it to 100%.
        hwnds_to_cleanup = self.managed_by_script_hwnds.difference(current_cycle_dynamically_managed_hwnds)
        for hwnd in list(hwnds_to_cleanup):
            if win32gui.IsWindow(hwnd) and not self._is_window_excluded(hwnd) and self.settings['dynamic_transparency_enabled']:
                # Only restore to 100% if dynamic is ON and it's no longer managed.
                set_transparency_for_hwnd(hwnd, 100)
            self.managed_by_script_hwnds.discard(hwnd)

    def restore_all_managed_to_full_opacity(self): # NEW: Method for the button
        """Restores all windows currently managed by the script to 100% opacity."""
        self.show_message("Restoring all managed windows to 100% opacity.", "blue")
        self._restore_managed_transparency_to_full_opacity() # Call the internal helper
        self.show_tooltip("All managed windows restored to 100%.")

    def _restore_managed_transparency_to_full_opacity(self):
        """Restores all windows currently managed by the script to 100% opacity,
        unless they are currently in the exclusion list. Excluded windows are simply
        removed from the managed set without their transparency being altered by the script."""
        
        hwnds_to_remove = set()
        for hwnd in list(self.managed_by_script_hwnds): # Iterate a copy for safe modification
            if not win32gui.IsWindow(hwnd):
                hwnds_to_remove.add(hwnd)
                continue

            # This function is explicitly for restoring to full opacity.
            # If a window is excluded, we still remove it from managed_by_script_hwnds,
            # but we don't attempt to set its transparency.
            if not self._is_window_excluded(hwnd):
                set_transparency_for_hwnd(hwnd, 100)
            hwnds_to_remove.add(hwnd) # Always remove from tracking after processing

        # Remove all processed HWNDs from the managed set
        self.managed_by_script_hwnds.difference_update(hwnds_to_remove)


    def _reset_inactivity_tracking_state(self):
        """
        Resets the internal state related to window inactivity tracking.
        Clears tracking data, but does NOT force restore previously minimized windows.
        Restoration will happen naturally if a minimized window gains focus.
        Re-initializes window_last_active_time for all currently visible, non-excluded windows.
        """
        self.show_message("Resetting window inactivity tracking state...", "blue")
        
        # 1. Clear windows minimized by the script from tracking, but do NOT restore them.
        self.minimized_by_script_hwnds.clear()
        
        # 2. Clear all inactivity tracking data
        self.window_last_active_time.clear()

        # 3. Re-populate window_last_active_time for all currently visible, non-excluded windows
        current_time_ms = time.time() * 1000
        
        def callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != "":
                # Skip our own UI and excluded windows
                if hwnd != self.root.winfo_id() and \
                   hwnd != self.tooltip_window.winfo_id() and \
                   not (self.changer_window and hwnd == self.changer_window.winfo_id()) and \
                   not self._is_window_excluded(hwnd):
                    self.window_last_active_time[hwnd] = current_time_ms
            return True
        win32gui.EnumWindows(callback, None)

        # Also ensure the current foreground window is marked active
        fg_hwnd = win32gui.GetForegroundWindow()
        if fg_hwnd and win32gui.IsWindow(fg_hwnd) and not self._is_window_excluded(fg_hwnd):
            self.window_last_active_time[fg_hwnd] = current_time_ms
            
        self.show_message("Inactivity tracking state reset.", "blue")

    def _start_window_monitoring(self):
        """Starts the periodic checks for foreground window changes, new windows, and inactive windows."""
        self._check_foreground_window()
        self._check_for_new_windows()
        self._check_for_inactive_windows() # NEW: Start inactive window monitoring

    def _check_for_inactive_windows(self):
        """Periodically checks for inactive windows and minimizes them based on settings."""
        if not self.settings['minimize_inactive_windows']:
            self.window_monitor_inactivity_timer = self.root.after(self.settings['window_monitor_interval_ms'], self._check_for_inactive_windows)
            return

        current_time_ms = time.time() * 1000
        current_fg_hwnd = win32gui.GetForegroundWindow()
        
        # Clean up window_last_active_time for invalid HWNDs
        for hwnd in list(self.window_last_active_time.keys()):
            if not win32gui.IsWindow(hwnd):
                del self.window_last_active_time[hwnd]
                self.minimized_by_script_hwnds.discard(hwnd)
                self.managed_by_script_hwnds.discard(hwnd) # Also remove from managed if invalid
                self.processed_new_windows.discard(hwnd)
                self.initial_script_start_hwnds.discard(hwnd)
                continue

        # Update last active time for foreground window
        if current_fg_hwnd and win32gui.IsWindow(current_fg_hwnd):
            self.window_last_active_time[current_fg_hwnd] = current_time_ms

        inactive_candidates = []
        for hwnd in list(self.initial_script_start_hwnds.union(self.processed_new_windows)): # Consider all known windows
            if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd) or not win32gui.GetWindowText(hwnd):
                # Clean up invalid/invisible windows
                self.initial_script_start_hwnds.discard(hwnd)
                self.processed_new_windows.discard(hwnd)
                self.managed_by_script_hwnds.discard(hwnd)
                self.minimized_by_script_hwnds.discard(hwnd)
                if hwnd in self.window_last_active_time:
                    del self.window_last_active_time[hwnd]
                continue

            # Skip our own UI, foreground window, and excluded windows
            if hwnd == self.root.winfo_id() or \
               hwnd == self.tooltip_window.winfo_id() or \
               (self.changer_window and hwnd == self.changer_window.winfo_id()) or \
               hwnd == current_fg_hwnd or \
               self._is_window_excluded(hwnd):
                continue

            # NEW: Explicitly exclude Electricsheep from minimization if crash protection is enabled
            if self.settings['enable_hotkey_passthrough']:
                exe_name = get_window_exe_name(hwnd)
                window_class = get_window_class_name(hwnd)
                if (exe_name and exe_name.lower() == 'es') or \
                   (window_class and window_class.lower() == 'electricsheepwndclass'):
                    # self.show_message(f"Skipping inactive minimization for Electricsheep (HWND: {hwnd}) due to crash protection.", "yellow")
                    continue # Skip Electricsheep

            # If already minimized by the script, keep it minimized
            if hwnd in self.minimized_by_script_hwnds:
                continue

            last_active = self.window_last_active_time.get(hwnd, current_time_ms) # Default to current time if not tracked yet
            if (current_time_ms - last_active) > self.settings['minimize_inactive_delay_ms']:
                inactive_candidates.append((last_active, hwnd))
        
        # Sort candidates by last active time (oldest first)
        inactive_candidates.sort()

        # Minimize all but the 'ignore_count' most recently active (still inactive) windows
        num_to_minimize = max(0, len(inactive_candidates) - self.settings['minimize_inactive_ignore_count'])

        for i in range(num_to_minimize):
            _, hwnd_to_minimize = inactive_candidates[i]
            if win32gui.IsWindow(hwnd_to_minimize) and win32gui.IsWindowVisible(hwnd_to_minimize) and not win32gui.IsIconic(hwnd_to_minimize):
                try:
                    win32gui.ShowWindow(hwnd_to_minimize, win32con.SW_MINIMIZE)
                    self.minimized_by_script_hwnds.add(hwnd_to_minimize)
                    # self.show_message(f"Minimized inactive window: {get_window_exe_name(hwnd_to_minimize)}", "yellow")
                except Exception as e:
                    self.show_message(f"Failed to minimize HWND {hwnd_to_minimize}: {e}", "orange")

        self.window_monitor_inactivity_timer = self.root.after(self.settings['window_monitor_interval_ms'], self._check_for_inactive_windows)

    def _restore_managed_windows_to_full_opacity(self):
        """Restores all windows currently managed by the script to 100% opacity,
        unless they are currently in the exclusion list. Excluded windows are simply
        removed from the managed set without their transparency being altered by the script."""
        
        hwnds_to_remove = set()
        for hwnd in list(self.managed_by_script_hwnds): # Iterate a copy for safe modification
            if not win32gui.IsWindow(hwnd):
                hwnds_to_remove.add(hwnd)
                continue

            if self._is_window_excluded(hwnd):
                # If currently excluded, just remove from managed set, DO NOT touch transparency.
                hwnds_to_remove.add(hwnd)
                # print(f"DEBUG: _restore_managed: Excluded window '{get_window_exe_name(hwnd) or get_window_class_name(hwnd)}' (HWND: {hwnd}) not restored to 100% opacity (removed from managed set).")
            else:
                # If not excluded, restore to 100%
                set_transparency_for_hwnd(hwnd, 100)
                hwnds_to_remove.add(hwnd) # Mark for removal after restoration

        # Remove all processed HWNDs from the managed set
        self.managed_by_script_hwnds.difference_update(hwnds_to_remove)

    def _center_window(self, hwnd, show_tooltip=True):
        """
        Centers the specified window on its primary monitor.
        If show_tooltip is True, displays a tooltip message.
        """
        if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd) or not win32gui.GetWindowText(hwnd):
            if show_tooltip:
                self.show_message("Cannot center window: not visible or invalid.", "red")
            self.show_message(f"Could not center HWND {hwnd}: not visible or invalid.", "red")
            return False

        # IMPORTANT: If Electricsheep is in global_transparency_exclusions, this check will prevent centering.
        # It has been removed from DEFAULT_SETTINGS for this purpose.
        if self._is_window_excluded(hwnd):
            if show_tooltip:
                self.show_message("Cannot center excluded window.", "red")
            self.show_message(f"Could not center HWND {hwnd}: window is excluded.", "red")
            return False

        try:
            monitor_info = win32api.GetMonitorInfo(win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTOPRIMARY))
            # 'Monitor' gives the physical screen bounds. 'Work' gives the usable area (excluding taskbar).
            # For Electricsheep, we want to size it to the *physical* screen width, but position it relative to the work area.
            monitor_rect = monitor_info['Monitor'] # (left, top, right, bottom) - physical screen
            work_area = monitor_info['Work']     # (left, top, right, bottom) - usable area

            physical_monitor_left = monitor_rect[0]
            physical_monitor_top = monitor_rect[1]
            physical_monitor_width = monitor_rect[2] - monitor_rect[0]
            physical_monitor_height = monitor_rect[3] - monitor_rect[1]

            work_area_left = work_area[0]
            work_area_top = work_area[1]
            work_area_width = work_area[2] - work_area[0]
            work_area_height = work_area[3] - work_area[1]

            # NEW: Special handling for Electricsheep
            if self.settings['center_electricsheep_special']:
                exe_name = get_window_exe_name(hwnd)
                window_class = get_window_class_name(hwnd)
                if (exe_name and exe_name.lower() == 'es') or \
                   (window_class and window_class.lower() == 'electricsheepwndclass'):
                    
                    # Based on user's provided metrics for Electricsheep for a 1920x1200 monitor with 23px taskbar:
                    # Desired Client area: (1920, 1177) which matches work_area_width, work_area_height
                    # Desired Window area: (1936, 1216)
                    # Desired Window position: (-8, -31)
                    
                    # New window dimensions (including borders and title bar)
                    # The goal is for the client area to fill the work area, with the window title bar hidden above.
                    # This means the window's total width should be physical_monitor_width (1920) + 16 (borders) = 1936
                    # The window's total height should be physical_monitor_height (1200) + 16 (borders) = 1216
                    # (assuming 31px for title bar + 8px for bottom border, client height = 1200 - 31 - 8 = 1161, which is not 1177)
                    # Let's re-evaluate based on the desired "Screen: x: -8 y: -31 w: 1936 h: 1216" for a 1920x1200 screen.

                    # To achieve a screen position of (-8, -31) and size (1936, 1216):
                    # The width should be physical_monitor_width + 16 (for 8px borders on each side)
                    new_width = physical_monitor_width + 16 
                    # The height should be physical_monitor_height + 16 (for 8px bottom border and 31px title bar, 1200 + 16 = 1216)
                    new_height = physical_monitor_height + 16 # This assumes 31px title bar + 8px bottom border.
                    
                    # Position it such that its left border is 8px left of the monitor's physical left edge
                    new_x = physical_monitor_left - 8
                    # Position it such that its top border is 31px above the monitor's physical top edge
                    new_y = physical_monitor_top - 31

                    win32gui.MoveWindow(hwnd, new_x, new_y, new_width, new_height, True)
                    if show_tooltip:
                        self.show_tooltip("Electricsheep Centered (Title bar hidden)!")
                    return True # Handled, exit function

            # Original centering logic if not Electricsheep or special centering is off
            # Get current window dimensions for standard centering
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            window_width = right - left
            window_height = bottom - top

            # Calculate new centered position relative to work area
            new_x = work_area_left + (work_area_width - window_width) // 2
            new_y = work_area_top + (work_area_height - window_height) // 2

            # Apply 'prevent_window_edges_off_screen' logic
            if self.settings['prevent_window_edges_off_screen']:
                new_x = max(work_area_left, new_x)
                new_y = max(work_area_top, new_y) # Ensure top is not off-screen
                # Also ensure it doesn't go off the right/bottom if window is larger than screen
                # Only apply if the window is smaller than the monitor in that dimension
                new_x = min(new_x, work_area_left + work_area_width - window_width) if window_width < work_area_width else new_x
                new_y = min(new_y, work_area_top + work_area_height - window_height) if window_height < work_area_height else new_y

            # Move the window
            win32gui.MoveWindow(hwnd, new_x, new_y, window_width, window_height, True)

            if show_tooltip:
                self.show_tooltip("Window Centered!")
            return True
        except Exception as e:
            if show_tooltip:
                self.show_message(f"Failed to center window: {e}", "red")
            self.show_message(f"Error centering window HWND {hwnd}: {e}", "red")
            return False

    def _get_hotkey_display_text(self, action, hotkey_str):
        """Generates the display text for a hotkey, including percentage if applicable."""
        if 'set_' in action and 'percent' in action:
            if 'xbutton2_shift' in action:
                return f"{hotkey_str} ({self.settings['transparency_levels']['preset_xbutton2_shift']}%)"
            elif 'xbutton2' in action:
                return f"{hotkey_str} ({self.settings['transparency_levels']['preset_xbutton2']}%)"
            elif 'xbutton1' in action:
                return f"{hotkey_str} ({self.settings['transparency_levels']['preset_xbutton1']}%)"
        elif 'set_' in action and 'brightness' in action: # Changed to elif to correctly handle brightness presets
            if 'set_80_percent_brightness' in action:
                return f"{hotkey_str} ({self.settings['brightness_levels']['preset_xbutton2']}%)"
            elif 'set_0_percent_brightness' in action:
                return f"{hotkey_str} ({self.settings['brightness_levels']['preset_xbutton1']}%)"

        parts = hotkey_str.split('+')
        display_parts = []
        for part in parts:
            part_lower = part.strip().lower()
            if part_lower == 'ctrl': display_parts.append('Ctrl')
            elif part_lower == 'shift': display_parts.append('Shift')
            elif part_lower == 'alt': display_parts.append('Alt')
            elif part_lower in ['win', 'windows']: display_parts.append('Win')
            else:
                found_display_name = self.internal_to_display_map.get(part, None)
                if found_display_name:
                    display_parts.append(found_display_name)
                else:
                    display_parts.append(part.upper() if len(part) == 1 else part)

        return '+'.join(display_parts) if display_parts else "Not Set"

    def apply_setting(self, entry_widget, category, key, is_top_level=False, value_type=int):
        """Applies a setting from an entry widget and validates input."""
        try:
            value_str = entry_widget.get()
            value = value_type(value_str)

            # --- Validation and Early Application Section ---
            # This block handles validation and some specific settings that return early.
            if category == 'transparency_levels':
                if key == 'fast_scroll_threshold_ms' and value < 0:
                    raise ValueError("Threshold cannot be negative.")
                if 'level' in key or 'preset' in key:
                    if not (1 <= value <= 100):
                        raise ValueError("Level must be between 1 and 100.")
                if 'increment' in key and value < 1:
                    raise ValueError("Increment must be at least 1.")
            elif category == 'brightness_levels': # Correct indentation for this 'elif'
                if key in ['initial', 'min', 'max', 'preset_xbutton2', 'preset_xbutton1']:
                    if not (0 <= value <= 100):
                        raise ValueError("Brightness level must be between 0 and 100.")
                elif key in ['scroll_increment_slow', 'scroll_increment_fast']:
                    if value < 1:
                        raise ValueError("Increment must be at least 1.")
                elif key == 'fast_scroll_threshold_ms':
                    if value < 0:
                        raise ValueError("Threshold cannot be negative.")
            # For top-level settings, handle specific cases here
            elif is_top_level:
                if category in ['new_window_transparency_level', 'active_window_transparency', 'inactive_window_transparency']:
                    if not (1 <= value <= 100):
                        raise ValueError("Level must be between 1 and 100.")
                elif category == 'tooltip_alpha':
                    value = max(0.0, min(1.0, value))
                    #self.settings[category] = value
                    #if self.tooltip_window.winfo_ismapped():
                    #    self.tooltip_window.attributes('-alpha', self.settings['tooltip_alpha'])
                    #self.show_message(f"Applied Tooltip Alpha: {value:.2f}", "green")
                    #self.save_settings() # Save immediately for tooltip_alpha
                    #return # Exit early as tooltip_alpha is fully handled here

                    self.settings[category] = value
                    # Apply new alpha using Windows API for smoother transition
                    if self.tooltip_window.winfo_exists(): # Check if window exists before trying to get HWND
                        hwnd_tooltip = self.tooltip_window.winfo_id()
                        new_alpha_percentage = self.settings['tooltip_alpha'] * 100
                        set_layered_window_colorkey_and_alpha(hwnd_tooltip, 0x00FF00, new_alpha_percentage)
                    self.show_message(f"Applied Tooltip Alpha: {value:.2f}", "green")
                    self.save_settings() # Save immediately for tooltip_alpha
                    return # Exit early as tooltip_alpha is fully handled here

            # --- General Application Section ---
            # This block applies the setting value and triggers related UI/logic updates.
            # It runs if the function hasn't returned early for tooltip_alpha.
            if is_top_level:
                self.settings[category] = value
                self.show_message(f"Applied {category.replace('_', ' ').title()}: {value}", "green")
                # These are now separate 'if' statements within the 'if is_top_level' block
                # to avoid the 'elif' chaining issue.
                if category in ['active_window_transparency', 'inactive_window_transparency'] and self.settings['dynamic_transparency_enabled']:
                    self._reapply_dynamic_transparency_on_all_windows()
                if category == 'new_window_transparency_level' and self.settings['apply_transparency_to_new_windows']:
                    pass # No specific action needed here beyond setting the value
            else: # For nested settings (e.g., transparency_levels sub-keys)
                self.settings[category][key] = value
                self.show_message(f"Applied {category.replace('_', ' ').title()}{' ' + key.replace('_', ' ').title() if key else ''}: {value}", "green")

                if category == 'transparency_levels':
                    if key == 'preset_xbutton2':
                        action = 'set_86_percent'
                        new_display_text = self._get_hotkey_display_text(action, self.settings['hotkeys'][action])
                        self.hotkey_labels[action].configure(text=new_display_text)
                    elif key == 'preset_xbutton2_shift':
                        action = 'set_100_percent'
                        new_display_text = self._get_hotkey_display_text(action, self.settings['hotkeys'][action])
                        self.hotkey_labels[action].configure(text=new_display_text)
                    elif key == 'preset_xbutton1':
                        action = 'set_30_percent'
                        new_display_text = self._get_hotkey_display_text(action, self.settings['hotkeys'][action])
                        self.hotkey_labels[action].configure(text=new_display_text)
                    # NEW: If any transparency setting changes, reset the transparency scrolling state
                    if key in ['initial', 'min', 'max', 'scroll_increment_slow', 'scroll_increment_fast', 'fast_scroll_threshold_ms', 'reset_on_scroll_start']:
                        self.is_transparency_scrolling = False
                        self.last_transparency_hotkey_press_time = 0
                elif category == 'brightness_levels': # Correct indentation for this 'elif'
                    # Update hotkey labels for brightness presets
                    if key == 'preset_xbutton2':
                        action = 'set_80_percent_brightness'
                        new_display_text = self._get_hotkey_display_text(action, self.settings['hotkeys'][action])
                        self.hotkey_labels[action].configure(text=new_display_text)
                    elif key == 'preset_xbutton1':
                        action = 'set_0_percent_brightness'
                        new_display_text = self._get_hotkey_display_text(action, self.settings['hotkeys'][action])
                        self.hotkey_labels[action].configure(text=new_display_text)
                    # If any brightness setting changes, reset the brightness scrolling state
                    if key in ['initial', 'min', 'max', 'scroll_increment_slow', 'scroll_increment_fast', 'fast_scroll_threshold_ms', 'reset_on_scroll_start']:
                        self.is_brightness_scrolling = False
                        self.last_brightness_hotkey_press_time = 0 # NEW: Reset hotkey press time
                        # Re-apply current brightness if it was changed
                        self._set_screen_brightness(self.current_brightness_level)

            self.save_settings() # Save settings after all modifications (if not returned early)

            # NEW: Conditional reset for inactivity tracking settings (placed correctly as a standalone 'if')
            # FIX: Removed 'minimize_inactive_ignore_count' from reset trigger to prevent mass popups when changing count
            if is_top_level and category in ['minimize_inactive_delay_ms']: 
                self._reset_inactivity_tracking_state()

        except ValueError as e:
            self.show_message(f"Invalid input for {category.replace('_', ' ').title()}{' ' + key.replace('_', ' ').title() if key else ''}: {e}", "red")
        except Exception as e:
            self.show_message(f"Error applying setting: {e}", "red")

    def show_message(self, message, color="white"):
        """Prints a message (can be expanded to a GUI message box/label)."""
        print(f"[{color}] {message}")

    def change_theme_color(self, new_theme):
        """Changes CustomTkinter theme color and shows restart warning."""
        self.settings['theme_color'] = new_theme
        customtkinter.set_default_color_theme(new_theme)
        self.save_settings()
        self.show_message(f"Theme color changed to {new_theme}. Restart recommended.", "blue")
        self.restart_warning_label.pack()

    def change_appearance_mode(self, new_mode):
        """Changes CustomTkinter appearance mode (Dark/Light/System)."""
        self.settings['appearance_mode'] = new_mode
        customtkinter.set_appearance_mode(new_mode)
        self.save_settings()
        self.show_message(f"Appearance mode changed to {new_mode}.", "blue")

    def toggle_ui_topmost(self):
        """Toggles the 'Always On Top' attribute for the main UI window."""
        new_state = self.ui_topmost_checkbox.get() == 1
        self.settings['ui_always_on_top'] = new_state
        self.save_settings()
        self.root.attributes('-topmost', new_state)
        self.show_message(f"UI 'Always On Top' set to: {new_state}", "blue")

    def toggle_show_mouse_position_ui(self):
        """Toggles the display of mouse position in the UI."""
        new_state = self.show_mouse_pos_checkbox.get() == 1
        self.settings['show_mouse_position_ui'] = new_state
        self.save_settings()
        if new_state:
            self.update_mouse_position_label()
            self.show_message("Showing mouse position in UI.", "blue")
        else:
            if self.mouse_pos_timer:
                self.root.after_cancel(self.mouse_pos_timer)
                self.mouse_pos_timer = None
            self.mouse_pos_label.configure(text="")
            self.show_message("Hiding mouse position in UI.", "blue")

    def _update_setting_from_checkbox(self, setting_key, checkbox_widget, category=None):
        """Generic helper to update a boolean setting from a checkbox widget."""
        new_state = checkbox_widget.get() == 1
        if category:
            self.settings[category][setting_key] = new_state
            self.show_message(f"'{category.replace('_', ' ').title()} - {setting_key.replace('_', ' ').title()}' set to: {new_state}", "blue")
            # Specific logic for transparency reset on scroll
            if category == 'transparency_levels' and setting_key == 'reset_on_scroll_start':
                self.is_transparency_scrolling = False
                self.last_transparency_hotkey_press_time = 0
        else:
            self.settings[setting_key] = new_state
            self.show_message(f"'{setting_key.replace('_', ' ').title()}' set to: {new_state}", "blue")
        self.save_settings()

    def _is_window_excluded(self, hwnd):
        """Checks if a window's executable name or class name is in the global exclusion list.
        Returns True if excluded, False otherwise."""
        if not win32gui.IsWindow(hwnd):
            # print(f"DEBUG: _is_window_excluded: HWND {hwnd} is not a valid window.")
            return False

        exe_name = get_window_exe_name(hwnd)
        window_class = get_window_class_name(hwnd)
        # window_text = win32gui.GetWindowText(hwnd) # Uncomment for more verbose debugging

        exclusion_list = [e.strip().lower() for e in self.settings['global_transparency_exclusions'].split(',') if e.strip()]
        
        # Uncomment for extensive debugging of exclusions
        # print(f"DEBUG: _is_window_excluded for HWND {hwnd} (EXE: '{exe_name}', Class: '{window_class}'). Exclusions: {exclusion_list}")

        if exe_name and exe_name in exclusion_list:
            # print(f"DEBUG: _is_window_excluded: Matched EXE '{exe_name}'. Returning True.")
            return True
        if window_class and window_class.lower() in exclusion_list:
            # print(f"DEBUG: _is_window_excluded: Matched Class '{window_class}'. Returning True.")
            return True
        
        return False

    def toggle_minimize_inactive_windows(self):
        """Toggles the 'minimize_inactive_windows' setting and applies changes."""
        new_state = self.minimize_inactive_windows_checkbox.get() == 1
        self.settings['minimize_inactive_windows'] = new_state
        self.save_settings()
        self.show_message(f"'Minimize inactive windows' set to: {new_state}", "blue")

        # Always reset inactivity tracking state when this setting is toggled.
        # This clears internal tracking. Windows previously minimized by the script
        # will NOT be automatically restored here. They will be restored when they gain focus.
        self._reset_inactivity_tracking_state()
        # The periodic _check_for_inactive_windows timer will handle subsequent minimization
        # based on the new setting state and reset timers.

    def _restore_minimized_windows_on_focus_change(self, new_fg_hwnd, old_fg_hwnd):
        """Restores windows that were minimized by the script if they gain focus,
        unless they are currently in the exclusion list."""
        if not self.settings['minimize_inactive_windows']:
            # If minimize inactive is off, ensure any windows previously minimized by us are restored if they become foreground.
            # This handles cases where the setting is toggled off, but a window was still minimized.
            if new_fg_hwnd in self.minimized_by_script_hwnds:
                if not self._is_window_excluded(new_fg_hwnd):
                    if win32gui.IsWindow(new_fg_hwnd):
                        win32gui.ShowWindow(new_fg_hwnd, win32con.SW_RESTORE)
                self.minimized_by_script_hwnds.discard(new_fg_hwnd)
            return

        # If the new foreground window was minimized by our script, attempt to restore it
        if new_fg_hwnd in self.minimized_by_script_hwnds:
            if not self._is_window_excluded(new_fg_hwnd): # Only restore if NOT excluded
                if win32gui.IsWindow(new_fg_hwnd):
                    win32gui.ShowWindow(new_fg_hwnd, win32con.SW_RESTORE)
                self.minimized_by_script_hwnds.discard(new_fg_hwnd)
            else:
                # If new_fg_hwnd is in minimized_by_script_hwnds but is now excluded,
                # it should not be restored by us, just remove from tracking.
                self.minimized_by_script_hwnds.discard(new_fg_hwnd)
        
        # If an old foreground window was minimized by us and is now excluded,
        # we should stop tracking it and NOT restore it.
        if old_fg_hwnd in self.minimized_by_script_hwnds and self._is_window_excluded(old_fg_hwnd):
            self.minimized_by_script_hwnds.discard(old_fg_hwnd)

    def update_mouse_position_label(self):
        """Updates the mouse position label in the UI."""
        if self.settings['show_mouse_position_ui']:
            x, y = win32api.GetCursorPos() # Using win32api.GetCursorPos() for consistency
            self.mouse_pos_label.configure(text=f"Mouse: X={x}, Y={y}")
            self.mouse_pos_timer = self.root.after(100, self.update_mouse_position_label)
        else:
            if self.mouse_pos_timer:
                self.root.after_cancel(self.mouse_pos_timer)
                self.mouse_pos_timer = None
            self.mouse_pos_label.configure(text="") # Clear text when disabled

    def restart_app(self):
        """Restarts the entire application."""
        self.show_message("Restarting application...", "yellow")
        self.save_settings()

        self.on_closing()

        os.execv(sys.executable, ['python'] + sys.argv)

    def setup_tooltip_window(self):
        """Initializes the custom semi-transparent tooltip window with rounded edges and transparent background."""
        self.tooltip_window = tk.Toplevel(self.root)
        self.tooltip_window.overrideredirect(True)
        self.tooltip_window.attributes('-topmost', True)

        self.tooltip_window.config(bg=self._CHROMA_KEY_COLOR_HEX)

        self.tooltip_frame = customtkinter.CTkFrame(self.tooltip_window,
                                                    fg_color="gray20",
                                                    corner_radius=10)
        self.tooltip_frame.pack(fill="both", expand=True)

        self.tooltip_label = customtkinter.CTkLabel(self.tooltip_frame, text="", font=("Arial", 14, "bold"))
        self.tooltip_label.pack(padx=15, pady=10)

        #self.tooltip_window.attributes('-transparentcolor', self._CHROMA_KEY_COLOR_HEX)
        #self.tooltip_window.attributes('-alpha', self.settings['tooltip_alpha'])
        # Apply Windows API layered window attributes for smoother transparency and chroma key
        hwnd_tooltip = self.tooltip_window.winfo_id()
        # Convert 0.0-1.0 alpha to 1-100 percentage for the helper function
        initial_alpha_percentage = self.settings['tooltip_alpha'] * 100
        set_layered_window_colorkey_and_alpha(hwnd_tooltip, 0x00FF00, initial_alpha_percentage)

        self.tooltip_window.withdraw()

    def show_tooltip(self, text, x_offset=None, y_offset=None):
        """Displays the tooltip near the mouse cursor with the given text and starts continuous following."""
        self.tooltip_label.configure(text=text)
        #self.tooltip_window.update_idletasks() # Force update of label content immediately
        #self.tooltip_window.deiconify()
        
        #self.tooltip_window.attributes('-alpha', self.settings['tooltip_alpha'])

        self.tooltip_window.update_idletasks() # Force update of label content immediately
        self.tooltip_window.deiconify()
        
        # Re-apply Windows API layered window attributes when showing tooltip
        hwnd_tooltip = self.tooltip_window.winfo_id()
        current_alpha_percentage = self.settings['tooltip_alpha'] * 100
        set_layered_window_colorkey_and_alpha(hwnd_tooltip, 0x00FF00, current_alpha_percentage)

        if self.tooltip_timer:
            self.root.after_cancel(self.tooltip_timer)
        self.tooltip_timer = self.root.after(self.settings['tooltip_display_time_ms'], self.hide_tooltip)
        
        # Store offsets for the loop
        self._current_tooltip_x_offset = x_offset if x_offset is not None else self.settings['tooltip_x_position']
        self._current_tooltip_y_offset = y_offset if y_offset is not None else self.settings['tooltip_y_position']

        self._start_tooltip_follow()

    def hide_tooltip(self):
        """Hides the tooltip window and stops continuous following."""
        self.tooltip_window.withdraw()
        self.tooltip_timer = None
        self._stop_tooltip_follow()

    def _start_tooltip_follow(self):
        """Starts the continuous loop to make the tooltip follow the mouse cursor."""
        if not self.tooltip_following:
            self.tooltip_following = True
            self._update_tooltip_position_loop()

    def _stop_tooltip_follow(self):
        """Stops the continuous tooltip following loop."""
        self.tooltip_following = False
        if self.tooltip_follow_timer:
            self.root.after_cancel(self.tooltip_follow_timer)
            self.tooltip_follow_timer = None

    def _update_tooltip_position_loop(self):
        """Continuously updates the tooltip's position to follow the mouse cursor, centered."""
        if self.tooltip_following and self.tooltip_window.winfo_exists():
            x, y = win32api.GetCursorPos()

            self.tooltip_window.update_idletasks()

            tooltip_width = self.tooltip_window.winfo_width()
            tooltip_height = self.tooltip_window.winfo_height()

            # Use stored offsets
            target_x = x + self._current_tooltip_x_offset - (tooltip_width // 2)
            target_y = y + self._current_tooltip_y_offset - (tooltip_height // 2)

            self.tooltip_window.wm_geometry(f"+{int(target_x)}+{int(target_y)}")
            self.tooltip_follow_timer = self.root.after(20, self._update_tooltip_position_loop)
        elif self.tooltip_follow_timer:
            self.root.after_cancel(self.tooltip_follow_timer)
            self.tooltip_follow_timer = None

    def update_status_label(self):
        """Updates the status label in the main GUI."""
        status_text = "Enabled" if self.script_enabled else "Disabled"
        self.status_label.configure(text=f"Status: {status_text}")
        self.toggle_button.configure(text=f"Toggle Script (Alt+W) - {'Disable' if self.script_enabled else 'Enable'}")

    def register_hotkeys(self):
        """
        Registers hotkeys based on current settings and script enabled state using AHK.
        Hotkeys for transparency, brightness, centering, and minimizing others are
        conditionally non-suppressing based on the 'enable_hotkey_passthrough' setting.
        """
        self.ahk.clear_hotkeys()

        # Hotkeys that are always non-suppressing by design (toggle script, focus mode)
        # Kill script hotkey should always be non-suppressing to ensure it works even if other hotkeys are suppressed.
        ahk_kill_hotkey = self._map_hotkey_to_ahk_syntax(self.settings['hotkeys']['kill_script_failsafe'], non_suppressing=True)
        if ahk_kill_hotkey:
            self.ahk.add_hotkey(ahk_kill_hotkey, self.kill_script)

        ahk_toggle_hotkey = self._map_hotkey_to_ahk_syntax(self.settings['hotkeys']['toggle_script'], non_suppressing=True)
        if ahk_toggle_hotkey:
            self.ahk.add_hotkey(ahk_toggle_hotkey, self.toggle_script_from_hotkey)

        ahk_toggle_focus_mode_hotkey = self._map_hotkey_to_ahk_syntax(self.settings['hotkeys']['toggle_focus_mode'], non_suppressing=True)
        if ahk_toggle_focus_mode_hotkey:
            self.ahk.add_hotkey(ahk_toggle_focus_mode_hotkey, self._ahk_toggle_focus_mode_callback)

        ahk_focus_mode_alt_tab_hotkey = self._map_hotkey_to_ahk_syntax(self.settings['hotkeys']['focus_mode_alt_tab'], non_suppressing=True)
        if ahk_focus_mode_alt_tab_hotkey:
            self.ahk.add_hotkey(ahk_focus_mode_alt_tab_hotkey, self._ahk_focus_mode_alt_tab_callback)

        if not self.script_enabled:
            self.show_message("Script is disabled. Only kill switch, toggle hotkey, and Focus Mode hotkeys are active.", "orange")
            return

        # Determine if problematic hotkeys should be non-suppressing
        use_non_suppressing_for_problematic = self.settings['enable_hotkey_passthrough']

        # Hotkeys that might interact with sensitive applications
        transparency_actions = [
            'increase_transparency', 'decrease_transparency',
            'set_86_percent', 'set_100_percent', 'set_30_percent'
        ]

        for action in transparency_actions:
            hotkey_str = self.settings['hotkeys'][action]
            ahk_hotkey = self._map_hotkey_to_ahk_syntax(hotkey_str, non_suppressing=use_non_suppressing_for_problematic)
            if ahk_hotkey:
                self.ahk.add_hotkey(ahk_hotkey, functools.partial(self._ahk_transparency_callback, action))
        
        center_hotkey_str = self.settings['hotkeys']['center_window']
        ahk_center_hotkey = self._map_hotkey_to_ahk_syntax(center_hotkey_str, non_suppressing=use_non_suppressing_for_problematic)
        if ahk_center_hotkey:
            self.ahk.add_hotkey(ahk_center_hotkey, self._ahk_center_window_callback)

        minimize_others_hotkey_str = self.settings['hotkeys']['minimize_others']
        ahk_minimize_others_hotkey = self._map_hotkey_to_ahk_syntax(minimize_others_hotkey_str, non_suppressing=use_non_suppressing_for_problematic)
        if ahk_minimize_others_hotkey:
            self.ahk.add_hotkey(ahk_minimize_others_hotkey, self._ahk_minimize_others_callback)

        brightness_actions = [
            'increase_brightness', 'decrease_brightness',
            'set_80_percent_brightness', 'set_0_percent_brightness'
        ]
        for action in brightness_actions:
            hotkey_str = self.settings['hotkeys'][action]
            ahk_hotkey = self._map_hotkey_to_ahk_syntax(hotkey_str, non_suppressing=use_non_suppressing_for_problematic)
            if ahk_hotkey:
                self.ahk.add_hotkey(ahk_hotkey, functools.partial(self._ahk_brightness_callback, action))

    def _map_hotkey_to_ahk_syntax(self, hotkey_str, non_suppressing=False):
        """Maps a hotkey string (e.g., 'ctrl+wheelup') to AHK syntax (e.g., '^WheelUp')."""
        if hotkey_str == 'none' or not hotkey_str:
            return None

        parts = hotkey_str.split('+')
        ahk_modifiers = []
        ahk_main_key = ''

        for part in parts:
            part_lower = part.strip().lower()
            if part_lower == 'ctrl':
                ahk_modifiers.append('^')
            elif part_lower == 'shift':
                ahk_modifiers.append('+')
            elif part_lower == 'alt':
                ahk_modifiers.append('!')
            elif part_lower in ['win', 'windows']:
                ahk_modifiers.append('#')
            else:
                found_ahk_key = next((k for k, v in self.internal_to_display_map.items()
                                      if k.lower() == part_lower or v.lower() == part_lower), None)
                if found_ahk_key:
                    ahk_main_key = found_ahk_key
                else:
                    if len(part) == 1 and (part.isalpha() or part.isdigit()):
                        ahk_main_key = part
                    else:
                        ahk_main_key = part.title()

        ahk_modifiers_str = "".join(sorted(ahk_modifiers, key=lambda x: ('^', '+', '!', '#').index(x)))

        prefix = '~' if non_suppressing else ''

        if ahk_main_key:
            return f"{prefix}{ahk_modifiers_str}{ahk_main_key}"
        elif ahk_modifiers_str and not non_suppressing:
            return None
        elif ahk_modifiers_str and non_suppressing:
            return None
        return None

    def _ahk_transparency_callback(self, action):
        """
        Generic callback for AHK hotkeys that modify transparency.
        This function runs in the AHK hotkey thread, so it schedules GUI updates.
        """
        if not self.script_enabled:
            return

        hotkey_config_str = self.settings['hotkeys'][action]

        if not self.check_modifiers_match(hotkey_config_str):
            if DEBUG_PRINT_MODIFIER_STATE_ON_MOUSE_EVENT:
                print(f"DEBUG: Modifiers mismatch for {action} with hotkey '{hotkey_config_str}'. Current state: Ctrl={self.ahk.key_state('Ctrl')}, Shift={self.ahk.key_state('Shift')}, Alt={self.ahk.key_state('Alt')}, Win={self.ahk.key_state('LWin') or self.ahk.key_state('RWin')}")
            return

        current_transparency_config = self.settings['transparency_levels'] # NEW: Get transparency config

        # NEW: Check if this is the start of a new scrolling sequence AND if reset is enabled
        current_hotkey_time = time.time() * 1000
        time_since_last_hotkey = current_hotkey_time - self.last_transparency_hotkey_press_time
        
        # A longer timeout to detect end of scroll sequence
        if time_since_last_hotkey > SCROLL_SEQUENCE_TIMEOUT_MS:
            self.is_transparency_scrolling = False

        if not self.is_transparency_scrolling and current_transparency_config['reset_on_scroll_start']:
            # Set internal transparency level to initial
            self.current_transparency_level = current_transparency_config['initial']
            # FIX: Ensure active_window_transparency (used by dynamic logic) is also reset here
            if self.settings['dynamic_transparency_enabled']:
                self.settings['active_window_transparency'] = current_transparency_config['initial']
                self.save_settings() # Persist this reset
            self.is_transparency_scrolling = True
        elif not self.is_transparency_scrolling and not current_transparency_config['reset_on_scroll_start']:
            self.is_transparency_scrolling = True
        
        self.last_transparency_hotkey_press_time = current_hotkey_time # Update last hotkey press time for next check


        if action == 'increase_transparency':
            self.root.after(0, lambda: self.update_transparency_gui(delta=1))
        elif action == 'decrease_transparency':
            self.root.after(0, lambda: self.update_transparency_gui(delta=-1))
        elif action == 'set_86_percent':
            self.root.after(0, lambda: self.update_transparency_gui(new_level=self.settings['transparency_levels']['preset_xbutton2']))
            self.is_transparency_scrolling = False # NEW: Presets are not part of a scroll sequence
        elif action == 'set_100_percent':
            self.root.after(0, lambda: self.update_transparency_gui(new_level=self.settings['transparency_levels']['preset_xbutton2_shift']))
            self.is_transparency_scrolling = False # NEW: Presets are not part of a scroll sequence
        elif action == 'set_30_percent':
            self.root.after(0, lambda: self.update_transparency_gui(new_level=self.settings['transparency_levels']['preset_xbutton1']))
            self.is_transparency_scrolling = False # NEW: Presets are not part of a scroll sequence
        else:
            self.show_message(f"Unhandled AHK hotkey action: {action}", "orange")

    def _ahk_center_window_callback(self):
        """
        Callback for AHK hotkey to center the active window.
        Schedules the centering logic on the main GUI thread.
        """
        if not self.script_enabled:
            return
        
        hotkey_config_str = self.settings['hotkeys']['center_window']
        if not self.check_modifiers_match(hotkey_config_str):
            if DEBUG_PRINT_MODIFIER_STATE_ON_MOUSE_EVENT:
                print(f"DEBUG: Modifiers mismatch for center_window with hotkey '{hotkey_config_str}'. Current state: Ctrl={self.ahk.key_state('Ctrl')}, Shift={self.ahk.key_state('Shift')}, Alt={self.ahk.key_state('Alt')}, Win={self.ahk.key_state('LWin') or self.ahk.key_state('RWin')}")
            return

        hwnd = win32gui.GetForegroundWindow()
        if not hwnd or hwnd == self.root.winfo_id() or hwnd == self.tooltip_window.winfo_id() or \
           (self.changer_window and hwnd == self.changer_window.winfo_id()):
            self.root.after(0, lambda: self.show_message("Cannot center UI window.", "red"))
            return
        
        # Corrected: Use _is_window_excluded - this is now handled inside _center_window.
        # if self._is_window_excluded(hwnd):
        #     self.root.after(0, lambda: self.show_message("Cannot center excluded window.", "red"))
        #     return

        self.root.after(0, lambda: self._center_window(hwnd, show_tooltip=True))

    def _ahk_minimize_others_callback(self):
        """
        Callback for AHK hotkey to minimize all windows except the clicked one.
        Schedules the minimization logic on the main GUI thread.
        """
        if not self.script_enabled:
            return

        hotkey_config_str = self.settings['hotkeys']['minimize_others']
        if not self.check_modifiers_match(hotkey_config_str):
            if DEBUG_PRINT_MODIFIER_STATE_ON_MOUSE_EVENT:
                print(f"DEBUG: Modifiers mismatch for minimize_others with hotkey '{hotkey_config_str}'. Current state: Ctrl={self.ahk.key_state('Ctrl')}, Shift={self.ahk.key_state('Shift')}, Alt={self.ahk.key_state('Alt')}, Win={self.ahk.key_state('LWin') or self.ahk.key_state('RWin')}")
            return

        mouse_x, mouse_y = win32api.GetCursorPos()
        clicked_hwnd_at_point = win32gui.WindowFromPoint((mouse_x, mouse_y)) # Get the window at the mouse point

        # Get the top-level parent window for the clicked HWND
        # This ensures we are working with a top-level window that EnumWindows will find.
        keep_hwnd = win32gui.GetAncestor(clicked_hwnd_at_point, win32con.GA_ROOT)

        if not keep_hwnd or not win32gui.IsWindowVisible(keep_hwnd) or not win32gui.GetWindowText(keep_hwnd):
            self.root.after(0, lambda: self.show_message("No valid top-level window found under cursor to keep open.", "red"))
            return

        if keep_hwnd == self.root.winfo_id() or keep_hwnd == self.tooltip_window.winfo_id() or \
           (self.changer_window and keep_hwnd == self.changer_window.winfo_id()):
            self.root.after(0, lambda: self.show_message("Cannot minimize others based on UI window.", "red"))
            return

        self.root.after(0, lambda: self._minimize_all_except_one(keep_hwnd, "Minimized others!"))

    def _minimize_all_except_one(self, keep_hwnd, tooltip_message, use_active_window=False):
        """
        Minimizes all visible windows except the specified keep_hwnd.
        If use_active_window is True, keep_hwnd is ignored and foreground window is used.
        """
        if use_active_window:
            keep_hwnd = win32gui.GetForegroundWindow()

        if not keep_hwnd or not win32gui.IsWindow(keep_hwnd) or not win32gui.IsWindowVisible(keep_hwnd) or not win32gui.GetWindowText(keep_hwnd):
            self.show_message("No valid window to keep open.", "red")
            return

        def callback(hwnd, extra):
            if hwnd == keep_hwnd:
                return True # Don't minimize the target window

            if not win32gui.IsWindowVisible(hwnd) or not win32gui.GetWindowText(hwnd):
                return True # Skip invisible or nameless windows

            if hwnd == self.root.winfo_id() or hwnd == self.tooltip_window.winfo_id() or \
               (self.changer_window and hwnd == self.changer_window.winfo_id()):
                return True # Skip script's own windows

            # NEW: Explicitly exclude Electricsheep from minimization if crash protection is enabled
            if self.settings['enable_hotkey_passthrough']:
                exe_name = get_window_exe_name(hwnd)
                window_class = get_window_class_name(hwnd)
                if (exe_name and exe_name.lower() == 'es') or \
                   (window_class and window_class.lower() == 'electricsheepwndclass'):
                    self.show_message(f"Skipping minimization for Electricsheep (HWND: {hwnd}) due to crash protection.", "yellow")
                    return True # Skip Electricsheep

            # Original exclusion check (for general exclusions)
            if self._is_window_excluded(hwnd):
                return True # Skip generally excluded windows

            # Check if already minimized
            if win32gui.IsIconic(hwnd):
                return True # Already minimized, skip

            try:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            except Exception as e:
                self.show_message(f"Failed to minimize HWND {hwnd}: {e}", "orange")
            return True

        win32gui.EnumWindows(callback, None)
        self.show_tooltip(tooltip_message, x_offset=self.settings['focus_tooltip_x_position'], y_offset=self.settings['focus_tooltip_y_position'])


    def toggle_focus_mode_ui(self):
        """Toggles the script's focus mode enabled state and updates the UI."""
        self.focus_mode_active = not self.focus_mode_active
        self.settings['focus_mode_active'] = self.focus_mode_active
        self.save_settings()
        
        if self.focus_mode_active:
            self.focus_mode_checkbox.select()
            tooltip_text = "Focus Mode ON"
        else:
            self.focus_mode_checkbox.deselect()
            tooltip_text = "Focus Mode OFF"
            # When turning off focus mode, ensure all previously minimized windows are restored if desired.
            # The AHK script doesn't restore, it just stops minimizing. We will follow that.

        self.show_tooltip(tooltip_text, x_offset=self.settings['focus_tooltip_x_position'], y_offset=self.settings['focus_tooltip_y_position'])
        self.show_message(f"Focus Mode state changed to: {self.focus_mode_active}", "blue")
        self.register_hotkeys() # Re-register hotkeys to reflect focus mode state for Alt+Tab

    def _ahk_toggle_focus_mode_callback(self):
        """Called when the toggle focus mode hotkey is pressed. Schedules UI update on main thread."""
        self.root.after(0, self.toggle_focus_mode_ui)

    def _ahk_focus_mode_alt_tab_callback(self):
        """
        Callback for AHK hotkey Alt+Tab when focus mode is active.
        Schedules the minimization logic on the main GUI thread after a delay.
        """
        hotkey_config_str = self.settings['hotkeys']['focus_mode_alt_tab']
        if not self.check_modifiers_match(hotkey_config_str):
            if DEBUG_PRINT_MODIFIER_STATE_ON_MOUSE_EVENT:
                print(f"DEBUG: Modifiers mismatch for focus_mode_alt_tab with hotkey '{hotkey_config_str}'. Current state: Ctrl={self.ahk.key_state('Ctrl')}, Shift={self.ahk.key_state('Shift')}, Alt={self.ahk.key_state('Alt')}, Win={self.ahk.key_state('LWin') or self.ahk.key_state('RWin')}")
            return

        if self.focus_mode_active:
            delay = self.settings['focus_mode_alt_tab_delay_ms']
            self.root.after(delay, lambda: self._minimize_all_except_one(None, "Focus Mode: Minimized others!", use_active_window=True))
        else:
            # If focus mode is off, the Alt+Tab hotkey still triggers but does nothing.
            # This matches the AHK script's behavior where the `if (is_focus_mode_active)` check is inside the hotkey.
            pass

    def toggle_script_from_hotkey(self):
        """Called when the toggle hotkey is pressed. Schedules UI update on main thread."""
        self.root.after(0, self.toggle_script_ui)
        self.show_message("Toggle script hotkey pressed.", "blue")

    def toggle_script_ui(self):
        """Toggles the script's enabled state and updates the UI."""
        self.script_enabled = not self.script_enabled
        self.settings['script_enabled'] = self.script_enabled
        self.save_settings()
        self.update_status_label()
        self.register_hotkeys()
        if not self.script_enabled:
            self.hide_tooltip()
        self.show_message(f"Script enabled state changed to: {self.script_enabled}", "green")

    def _initialize_hotkey_maps(self):
        """Initializes internal and display mappings for hotkeys, compatible with AHK and custom order."""
        self.internal_to_display_map = {
            'none': 'None',
            'WheelUp': 'Mouse Wheel Up',
            'WheelDown': 'Mouse Wheel Down',
            'XButton1': 'Mouse XButton1',
             'XButton2': 'Mouse XButton2',
            'LButton': 'Mouse Left Click',
            'RButton': 'Mouse Right Click', # Add RButton
            'MButton': 'Mouse Middle Click',
            'WheelUp': 'Mouse Wheel Up', # Ensure WheelUp/Down are here for Alt+Wheel
            'WheelDown': 'Mouse Wheel Down',
            'Ctrl': 'Ctrl', 'Shift': 'Shift', 'Alt': 'Alt', 'Win': 'Win',
        }

        ahk_key_mappings = {
            'Backspace': 'Backspace', 'Tab': 'Tab', 'Enter': 'Enter', 'Escape': 'Esc', 'Space': 'Space',
            'PgUp': 'Page Up', 'PgDn': 'Page Down', 'End': 'End', 'Home': 'Home',
            'Left': 'Left', 'Up': 'Up', 'Right': 'Right', 'Down': 'Down',
            'Insert': 'Ins', 'Delete': 'Del',
            'Numpad0': 'Numpad 0', 'Numpad1': 'Numpad 1', 'Numpad2': 'Numpad 2', 'Numpad3': 'Numpad 3', 'Numpad4': 'Numpad 4',
            'Numpad5': 'Numpad 5', 'Numpad6': 'Numpad 6', 'Numpad7': 'Numpad 7', 'Numpad8': 'Numpad 8', 'Numpad9': 'Numpad 9',
            'NumpadMult': 'Num *', 'NumpadAdd': 'Num +', 'NumpadSub': 'Num -', 'NumpadDot': 'Num .', 'NumpadDiv': 'Num /',
            'F1': 'F1', 'F2': 'F2', 'F3': 'F3', 'F4': 'F4', 'F5': 'F5', 'F6': 'F6', 'F7': 'F7', 'F8': 'F8', 'F9': 'F9', 'F10': 'F10', 'F11': 'F11', 'F12': 'F12',
            '~': '~', '-': '-', '=': '=', '[': '[', ']': ']', ';': ';', '"': "'", ',': ',', '.': '.', '\\': '\\', '/': '/',
        }

        for i in range(10):
            str_i = str(i)
            ahk_key_mappings[str_i] = str_i

        for char_code in range(ord('A'), ord('Z') + 1):
            char_upper = chr(char_code)
            ahk_key_mappings[char_upper] = char_upper

        for char_code in range(ord('a'), ord('z') + 1):
            char_lower = chr(char_code)
            ahk_key_mappings[char_lower] = char_lower.upper()

        for ahk_name, display_name in ahk_key_mappings.items():
            if ahk_name not in self.internal_to_display_map:
                self.internal_to_display_map[ahk_name] = display_name

        self.display_to_internal_map = {v: k for k, v in self.internal_to_display_map.items()
                                        if k not in ['Ctrl', 'Shift', 'Alt', 'Win']}

    def _get_hotkey_dropdown_values(self):
            """Returns a sorted list of display names for the hotkey dropdown using custom order,
            with single letters A-Z sorted alphabetically and inserted logically."""
            # Get all possible display names, excluding modifiers
            all_display_names = set(v for k, v in self.internal_to_display_map.items() if k not in ['Ctrl', 'Shift', 'Alt', 'Win'])

            dropdown_list_ordered = []
            letters_az = []
            remaining_unsorted = []

            # First, add items from the custom order list, maintaining their order
            for name in self._CUSTOM_KEY_DISPLAY_ORDER:
                if name in all_display_names:
                    # Check if it's a single uppercase letter (A-Z)
                    # The internal_to_display_map converts 'a' to 'A', so we check for uppercase here.
                    if len(name) == 1 and name.isalpha() and name.isupper():
                        letters_az.append(name)
                    else:
                        dropdown_list_ordered.append(name)
                    all_display_names.discard(name) # Remove from set to process remaining

            # Now, process any remaining names (those not in _CUSTOM_KEY_DISPLAY_ORDER)
            for name in all_display_names:
                if len(name) == 1 and name.isalpha() and name.isupper():
                    letters_az.append(name)
                else:
                    remaining_unsorted.append(name)

            letters_az.sort() # Sort A-Z letters alphabetically
            remaining_unsorted.sort() # Sort any other remaining keys alphabetically

            # Find the insertion point for letters (e.g., after 'Del' and before '0')
            insert_index_for_letters = -1
            try:
                # Find 'Del' and insert after it
                insert_index_for_letters = dropdown_list_ordered.index('Del') + 1
            except ValueError:
                # If 'Del' isn't there, try to find 'Down' as a fallback
                try:
                    insert_index_for_letters = dropdown_list_ordered.index('Down') + 1
                except ValueError:
                    # Default to inserting before numbers if no specific marker is found
                    # Find the first numeric key (e.g., '0')
                    for i, name in enumerate(dropdown_list_ordered):
                        if name.isdigit() or name.startswith('Numpad'):
                            insert_index_for_letters = i
                            break
                    if insert_index_for_letters == -1: # If no numbers found, append to end of custom ordered
                        insert_index_for_letters = len(dropdown_list_ordered)

            # Insert sorted letters into the custom ordered list at the determined index
            dropdown_list_ordered[insert_index_for_letters:insert_index_for_letters] = letters_az

            # Append any other sorted keys that were not in _CUSTOM_KEY_DISPLAY_ORDER and not A-Z
            dropdown_list_ordered.extend(remaining_unsorted)

            # Handle 'None' at the very beginning
            if 'None' in dropdown_list_ordered:
                dropdown_list_ordered.remove('None')
            dropdown_list_ordered.insert(0, 'None')
            
            return dropdown_list_ordered

    def open_manual_hotkey_changer(self, action):
        """Opens a modal window to configure a hotkey using modifier checkboxes and a key/action dropdown (manual selection)."""
        if self.changer_window:
            return

        self.root.attributes('-topmost', False)
        self.changer_window = customtkinter.CTkToplevel(self.root)
        self.changer_window.title(f"Configure {action.replace('_', ' ').title()} Hotkey")
        self.changer_window.geometry("350x300")
        self.changer_window.attributes('-topmost', True)
        self.changer_window.grab_set()
        self.changer_window.resizable(False, False)
        self.changer_window.protocol("WM_DELETE_WINDOW", self.cancel_hotkey_capture)

        current_hotkey_str = self.settings['hotkeys'][action]
        current_modifiers, current_main_key = self._parse_hotkey_string_for_changer(current_hotkey_str)

        modifier_frame = customtkinter.CTkFrame(self.changer_window, fg_color="transparent")
        modifier_frame.pack(pady=10, padx=20, fill="x")
        customtkinter.CTkLabel(modifier_frame, text="Modifiers:", font=("Arial", 12, "bold")).pack(anchor="w")

        self.modifier_vars = {
            'win': customtkinter.BooleanVar(value='win' in current_modifiers),
            'ctrl': customtkinter.BooleanVar(value='ctrl' in current_modifiers),
            'shift': customtkinter.BooleanVar(value='shift' in current_modifiers),
            'alt': customtkinter.BooleanVar(value='alt' in current_modifiers),
        }
        customtkinter.CTkCheckBox(modifier_frame, text="Win", variable=self.modifier_vars['win']).pack(anchor="w", padx=10, pady=2)
        customtkinter.CTkCheckBox(modifier_frame, text="Ctrl", variable=self.modifier_vars['ctrl']).pack(anchor="w", padx=10, pady=2)
        customtkinter.CTkCheckBox(modifier_frame, text="Shift", variable=self.modifier_vars['shift']).pack(anchor="w", padx=10, pady=2)
        customtkinter.CTkCheckBox(modifier_frame, text="Alt", variable=self.modifier_vars['alt']).pack(anchor="w", padx=10, pady=2)

        key_frame = customtkinter.CTkFrame(self.changer_window, fg_color="transparent")
        key_frame.pack(pady=10, padx=20, fill="x")
        customtkinter.CTkLabel(key_frame, text="Main Key/Action:", font=("Arial", 12, "bold")).pack(anchor="w")

        dropdown_values = self._get_hotkey_dropdown_values()

        initial_dropdown_value = self.internal_to_display_map.get(current_main_key, 'None')
        if initial_dropdown_value not in dropdown_values:
             initial_dropdown_value = 'None'
        self.main_key_var = customtkinter.StringVar(value=initial_dropdown_value)

        self.main_key_dropdown = customtkinter.CTkOptionMenu(key_frame, values=dropdown_values, variable=self.main_key_var)
        self.main_key_dropdown.pack(fill="x", padx=10, pady=5)

        apply_button = customtkinter.CTkButton(self.changer_window, text="Apply Hotkey",
                                               command=lambda: self._apply_unified_hotkey(action))
        apply_button.pack(pady=10)

        self.hotkey_capture_active = True
        self.current_hotkey_action = action

    def _parse_hotkey_string_for_changer(self, hotkey_str):
        """
        Parses a hotkey string into its modifier components and the main key/action.
        Returns (list_of_modifiers, main_key_or_action_string).
        """
        parts = hotkey_str.split('+')
        modifiers = []
        main_key = 'none'

        for part in parts:
            part_lower = part.strip().lower()
            if part_lower in ['ctrl', 'shift', 'alt', 'win', 'windows']:
                modifiers.append(part_lower.replace('windows', 'win'))
            else:
                found_internal_key = next((k for k in self.internal_to_display_map.keys() if k.lower() == part_lower), None)
                if not found_internal_key:
                    found_internal_key = next((k for k, v in self.internal_to_display_map.items() if v.lower() == part_lower), None)

                if found_internal_key:
                    main_key = found_internal_key
                else:
                    main_key = part
        return sorted(modifiers), main_key

    def _apply_unified_hotkey(self, action):
        """Applies the selected modifiers and key/action for a hotkey (manual selection)."""
        selected_modifiers = []
        if self.modifier_vars['win'].get():
            selected_modifiers.append('win')
        if self.modifier_vars['ctrl'].get():
            selected_modifiers.append('ctrl')
        if self.modifier_vars['shift'].get():
            selected_modifiers.append('shift')
        if self.modifier_vars['alt'].get():
            selected_modifiers.append('alt')

        selected_main_key_display = self.main_key_var.get()
        selected_main_key_internal = self.display_to_internal_map.get(selected_main_key_display, 'none')

        if selected_main_key_internal == 'none' and not selected_modifiers:
            new_hotkey_str = 'none'
        else:
            modifier_part = '+'.join(sorted(selected_modifiers))
            if selected_main_key_internal == 'none':
                new_hotkey_str = modifier_part
            else:
                new_hotkey_str = f"{modifier_part}+{selected_main_key_internal}" if modifier_part else selected_main_key_internal

        for other_action, existing_hotkey in self.settings['hotkeys'].items():
            if other_action != action and existing_hotkey == new_hotkey_str and new_hotkey_str != 'none':
                self.show_message(f"Conflict: '{self._get_hotkey_display_text(other_action, new_hotkey_str)}' is already assigned to {other_action.replace('_', ' ').title()}.", "red")
                return

        self.settings['hotkeys'][action] = new_hotkey_str
        new_display_text = self._get_hotkey_display_text(action, new_hotkey_str)
        self.hotkey_labels[action].configure(text=new_display_text)
        self.save_settings()
        self.show_message(f"Hotkey for {action.replace('_', ' ').title()} set to '{new_display_text}'.", "green")
        self.finalize_hotkey_capture()

    def cancel_hotkey_capture(self):
        """Cancels hotkey configuration and restores normal operation."""
        self.show_message("Hotkey configuration cancelled.", "orange")
        self.finalize_hotkey_capture()

    def finalize_hotkey_capture(self):
        """Cleans up hotkey changer state and re-registers hotkeys."""
        if not self.hotkey_capture_active:
            return

        self.hotkey_capture_active = False

        if self.changer_window:
            self.changer_window.destroy()
            self.changer_window.grab_release()
            self.changer_window = None

        self.root.attributes('-topmost', self.settings['ui_always_on_top'])
        self.register_hotkeys()
        self.current_hotkey_action = None

    def check_modifiers_match(self, hotkey_config_str):
        """
        Checks if the currently pressed modifiers (Win, Ctrl, Shift, Alt)
        exactly match the modifiers specified in the hotkey_config_str using ahk.key_state().
        """
        modifier_parts = []
        for part in hotkey_config_str.split('+'):
            part_lower = part.strip().lower()
            if part_lower in ['ctrl', 'shift', 'alt', 'win', 'windows']:
                modifier_parts.append(part_lower.replace('windows', 'win'))

        requires_win = 'win' in modifier_parts
        requires_ctrl = 'ctrl' in modifier_parts
        requires_shift = 'shift' in modifier_parts
        requires_alt = 'alt' in modifier_parts

        current_win_pressed = self.ahk.key_state('LWin') or self.ahk.key_state('RWin')
        current_ctrl_pressed = self.ahk.key_state('LCtrl') or self.ahk.key_state('RCtrl')
        current_shift_pressed = self.ahk.key_state('LShift') or self.ahk.key_state('RShift')
        current_alt_pressed = self.ahk.key_state('LAlt') or self.ahk.key_state('RAlt')

        if (current_win_pressed != requires_win or
            current_ctrl_pressed != requires_ctrl or
            current_shift_pressed != requires_shift or
            current_alt_pressed != requires_alt):
            return False

        return True

    def update_transparency_gui(self, new_level=None, delta=0):
        """
        Updates the transparency of the foreground window and shows a tooltip.
        This function is scheduled to run on the main GUI thread.
        """
        if not self.script_enabled:
            return

        hwnd = win32gui.GetForegroundWindow()
        if not hwnd or hwnd == self.root.winfo_id() or hwnd == self.tooltip_window.winfo_id() or \
           (self.changer_window and hwnd == self.changer_window.winfo_id()):
            return

        # Corrected: Use _is_window_excluded
        if self._is_window_excluded(hwnd):
            # If the foreground window is excluded, do not apply transparency changes via hotkey.
            self.show_tooltip(f"'{get_window_exe_name(hwnd) or get_window_class_name(hwnd)}' is excluded from transparency changes.")
            self.show_message(f"Attempted to change transparency for excluded window '{get_window_exe_name(hwnd) or get_window_class_name(hwnd)}'. Ignored.", "yellow")
            return

        # If dynamic transparency is enabled, hotkeys should modify the 'active' level
        if self.settings['dynamic_transparency_enabled']:
            # Hotkeys should always be able to change transparency of the foreground window
            # if dynamic transparency is enabled, regardless of 'manage_all' or 'manual update' settings.
            # The foreground window is explicitly targeted by the user.

            current_active_level = self.settings['active_window_transparency']
            calculated_new_active_level = current_active_level # Initialize with current for cases where new_level is None

            if new_level is not None:
                calculated_new_active_level = new_level
            elif delta != 0:
                current_time = time.time() * 1000
                time_diff = current_time - self.last_scroll_time
                self.last_scroll_time = current_time

                if time_diff < self.settings['transparency_levels']['fast_scroll_threshold_ms'] and time_diff > 0:
                    increment = self.settings['transparency_levels']['scroll_increment_fast']
                else:
                    increment = self.settings['transparency_levels']['scroll_increment_slow']
                calculated_new_active_level = current_active_level + (increment * delta)

            calculated_new_active_level = max(self.settings['transparency_levels']['min'],
                                   min(self.settings['transparency_levels']['max'],
                                       calculated_new_active_level))
            self.settings['active_window_transparency'] = calculated_new_active_level # THIS LINE IS KEY FOR THE NUANCE
            self.current_transparency_level = calculated_new_active_level # Keep for tooltip display consistency
            
            # Crucially, add the window to managed_by_script_hwnds if hotkey was successful
            self.managed_by_script_hwnds.add(hwnd) # Ensure it's now dynamically managed
            self.save_settings() # Save the updated active level
            
            # Reapply dynamic transparency to ensure all windows are updated, especially the foreground one
            # and potentially other inactive ones.
            self._reapply_dynamic_transparency_on_all_windows(force_all=self.settings['manage_all_windows_dynamically'])

        else: # Dynamic transparency is NOT enabled, use the old logic for direct transparency
            # Hotkey changes should directly apply to the foreground window if dynamic is OFF.
            if new_level is not None:
                self.current_transparency_level = new_level
            elif delta != 0:
                current_time = time.time() * 1000
                time_diff = current_time - self.last_scroll_time
                self.last_scroll_time = current_time

                if time_diff < self.settings['transparency_levels']['fast_scroll_threshold_ms'] and time_diff > 0:
                    increment = self.settings['transparency_levels']['scroll_increment_fast']
                else:
                    increment = self.settings['transparency_levels']['scroll_increment_slow']

                self.current_transparency_level += (increment * delta)

            self.current_transparency_level = max(self.settings['transparency_levels']['min'],
                                                  min(self.settings['transparency_levels']['max'],
                                                      self.current_transparency_level))
            # Apply transparency to the current foreground window
            success = set_transparency_for_hwnd(hwnd, self.current_transparency_level)
            if success:
                # If not dynamically managed, hotkey changes add it to managed for potential future restoration
                # or if dynamic mode is later enabled.
                self.managed_by_script_hwnds.add(hwnd)
            else:
                if self.last_processed_hwnd != hwnd:
                    self.show_tooltip(f"Failed to set transparency for window.", "red")
                    self.show_message(f"Could not set transparency for HWND {hwnd}. It might not support layering or require elevated privileges.", "red")
                self.last_processed_hwnd = hwnd


        self.show_tooltip(f"Transparency: {self.current_transparency_level}%")
        self.last_processed_hwnd = hwnd # Update last processed HWND regardless of success for message suppression

    def _on_entry_scroll(self, event, entry_widget, category, key, is_top_level, value_type, increment):
        """Handles mouse wheel scrolling on an entry widget to adjust its value."""
        entry_widget.focus_set()
        delta_sign = 1 if event.delta > 0 else -1
        self._adjust_entry_value(entry_widget, category, key, is_top_level, delta_sign, value_type, increment)
        return "break"

    def _on_entry_arrow_key(self, event, entry_widget, category, key, is_top_level, delta_sign, value_type, increment):
        """Handles Up/Down arrow key presses on an entry widget to adjust its value."""
        self._adjust_entry_value(entry_widget, category, key, is_top_level, delta_sign, value_type, increment)
        return "break"

    def _adjust_entry_value(self, entry_widget, category, key, is_top_level, delta_sign, value_type, increment):
        """Adjusts the value in an entry widget based on scroll/arrow key input."""
        try:
            current_value = value_type(entry_widget.get())
            new_value = current_value + (delta_sign * increment)

            # Apply bounds and specific logic based on category and key
            if category == 'transparency_levels':
                if key in ['initial', 'min', 'max', 'preset_xbutton2', 'preset_xbutton2_shift', 'preset_xbutton1']:
                    new_value = max(1, min(100, new_value))
                elif key in ['scroll_increment_slow', 'scroll_increment_fast']:
                    new_value = max(1, new_value)
                elif key == 'fast_scroll_threshold_ms':
                    new_value = max(0, new_value)
            elif category == 'brightness_levels':
                if key in ['initial', 'min', 'max', 'preset_xbutton2', 'preset_xbutton1']:
                    new_value = max(0, min(100, new_value))
                elif key in ['scroll_increment_slow', 'scroll_increment_fast']: # Removed scroll_stop_delay_ms
                    new_value = max(1, new_value)
                elif key == 'fast_scroll_threshold_ms':
                    new_value = max(0, new_value)
            elif is_top_level:
                if category == 'tooltip_alpha':
                    new_value = max(0.0, min(1.0, new_value))
                    new_value = round(new_value, 2)
                elif category in ['new_window_transparency_level', 'active_window_transparency', 'inactive_window_transparency']:
                    new_value = max(1, min(100, new_value))

            entry_widget.delete(0, customtkinter.END)
            if value_type == float:
                entry_widget.insert(0, f"{new_value:.2f}")
            else:
                entry_widget.insert(0, str(new_value))

            self.apply_setting(entry_widget, category, key, is_top_level, value_type)

        except ValueError:
            self.show_message("Invalid number format in entry.", "red")
        except Exception as e:
            self.show_message(f"Error adjusting value: {e}", "red")

    def reset_to_defaults(self):
            """Restores all settings to their default values and refreshes the UI."""
            self.settings = DEFAULT_SETTINGS.copy()
            self.save_settings()

            self.theme_menu_var.set(self.settings['theme_color'])
            self.appearance_menu_var.set(self.settings['appearance_mode'])
            self.apply_theme_settings()

            if self.settings['ui_always_on_top']:
                self.ui_topmost_checkbox.select()
            else:
                self.ui_topmost_checkbox.deselect()
            self.root.attributes('-topmost', self.settings['ui_always_on_top'])

            if self.settings['show_mouse_position_ui']:
                self.show_mouse_pos_checkbox.select()
                self.update_mouse_position_label()
            else:
                if self.mouse_pos_timer:
                    self.root.after_cancel(self.mouse_pos_timer)
                    self.mouse_pos_timer = None
                self.mouse_pos_label.configure(text="")

            # Update new window transparency checkbox
            if self.settings['apply_transparency_to_new_windows']:
                self.new_window_transparency_checkbox.select()
            else:
                self.new_window_transparency_checkbox.deselect()

            # Update dynamic transparency checkbox
            if self.settings['dynamic_transparency_enabled']:
                self.dynamic_transparency_checkbox.select()
            else:
                self.dynamic_transparency_checkbox.deselect()

            # Update manage all windows switch
            if self.settings['manage_all_windows_dynamically']:
                self.manage_all_windows_dynamically_switch.select()
            else:
                self.manage_all_windows_dynamically_switch.deselect()

            # New: Controls for granular behavior
            if self.settings['inactive_window_auto_update']:
                self.inactive_window_auto_update_checkbox.select()
            else:
                self.inactive_window_auto_update_checkbox.deselect()

            # New: Center on First Launch checkbox
            if self.settings['center_on_first_launch']:
                self.center_on_first_launch_checkbox.select()
            else:
                self.center_on_first_launch_checkbox.deselect()

            # New: Prevent Window Edges Off Screen checkbox
            if self.settings['prevent_window_edges_off_screen']:
                self.prevent_edges_off_screen_checkbox.select()
            else:
                self.prevent_edges_off_screen_checkbox.deselect()

            # NEW: Electricsheep Special Centering Checkbox
            if self.settings['center_electricsheep_special']:
                self.center_electricsheep_special_checkbox.select()
            else:
                self.center_electricsheep_special_checkbox.deselect()

            # NEW: Electricsheep Crash Protection Checkbox
            if self.settings['enable_hotkey_passthrough']:
                self.enable_hotkey_passthrough_checkbox.select()
            else:
                self.enable_hotkey_passthrough_checkbox.deselect()

            # New: Minimize Inactive Windows checkbox
            if self.settings['minimize_inactive_windows']:
                self.minimize_inactive_windows_checkbox.select()
            else:
                self.minimize_inactive_windows_checkbox.deselect()

            # New: Apply on Script Start checkbox
            if self.settings['apply_on_script_start']:
                self.apply_on_script_start_checkbox.select()
            else:
                self.apply_on_script_start_checkbox.deselect()

            # New: Brightness reset on scroll checkbox
            if self.settings['brightness_levels']['reset_on_scroll_start']:
                self.brightness_reset_on_scroll_checkbox.select()
            else:
                self.brightness_reset_on_scroll_checkbox.deselect()

            # NEW: Transparency reset on scroll checkbox
            if self.settings['transparency_levels']['reset_on_scroll_start']:
                self.transparency_reset_on_scroll_checkbox.select()
            else:
                self.transparency_reset_on_scroll_checkbox.deselect()

            # New: Focus Mode checkbox
            if self.settings['focus_mode_active']:
                self.focus_mode_checkbox.select()
            else:
                self.focus_mode_checkbox.deselect()

            for action, hotkey in self.settings['hotkeys'].items():
                if action in self.hotkey_labels:
                    new_display_text = self._get_hotkey_display_text(action, hotkey)
                    self.hotkey_labels[action].configure(text=new_display_text)

            # Update all setting entry widgets
            for key_tuple, (entry_widget, value_type, increment) in self.setting_entries.items():
                entry_widget.delete(0, customtkinter.END)
                category = key_tuple[0]
                if len(key_tuple) == 1: # Top-level setting
                    value = self.settings[category]
                else: # Nested setting
                    key = key_tuple[1]
                    value = self.settings[category][key]

                if value_type == float:
                    entry_widget.insert(0, f"{value:.2f}")
                else:
                    entry_widget.insert(0, str(value))

            # Update all exclusion list entry widgets (now only one)
            for setting_key, entry_widget in self.exclusion_list_entries.items():
                entry_widget.delete(0, customtkinter.END)
                entry_widget.insert(0, self.settings[setting_key])

            self.current_transparency_level = self.settings['transparency_levels']['initial']
            self.script_enabled = self.settings['script_enabled']
            self.focus_mode_active = self.settings['focus_mode_active'] # Reset focus mode state
            self.update_status_label()

            self.register_hotkeys()

            self.show_message("Settings reset to defaults.", "green")
            self.restart_warning_label.pack_forget()

            # Manually re-apply tooltip alpha if it's visible after reset
            if self.tooltip_window.winfo_ismapped():
                hwnd_tooltip = self.tooltip_window.winfo_id()
                current_alpha_percentage = self.settings['tooltip_alpha'] * 100
                set_layered_window_colorkey_and_alpha(hwnd_tooltip, 0x00FF00, current_alpha_percentage)

            # Restore all windows to 100% opacity and clear managed lists
            self._restore_managed_transparency_to_full_opacity()
            self.managed_by_script_hwnds.clear()
            
            # Reset brightness state
            self.current_brightness_level = self.settings['brightness_levels']['initial']
            self.is_brightness_scrolling = False
            self.last_brightness_scroll_time = 0
            self.last_brightness_hotkey_press_time = 0
            self._set_screen_brightness(self.current_brightness_level) # Apply default brightness

            # Reset transparency state
            self.is_transparency_scrolling = False # NEW
            self.last_transparency_hotkey_press_time = 0 # NEW

            self.processed_new_windows.clear()
            self.initial_script_start_hwnds.clear() # Re-populate on next script start
            # NEW: Use _reset_inactivity_tracking_state for a comprehensive reset of minimization and tracking
            self._reset_inactivity_tracking_state() 
        
            # Re-initialize the initial window list
            self._populate_initial_script_hwnds()
            current_time_ms = time.time() * 1000
            for hwnd in self.initial_script_start_hwnds:
                self.window_last_active_time[hwnd] = current_time_ms
            fg_hwnd = win32gui.GetForegroundWindow()
            if fg_hwnd:
                self.window_last_active_time[fg_hwnd] = current_time_ms

            # Reapply dynamic transparency if enabled by defaults
            if self.settings['dynamic_transparency_enabled']:
                if self.settings['manage_all_windows_dynamically']:
                    # If managing all, add all initial non-excluded windows to the managed set
                    for hwnd in self.initial_script_start_hwnds:
                        if not self._is_window_excluded(hwnd):
                            self.managed_by_script_hwnds.add(hwnd)
                    self._reapply_dynamic_transparency_on_all_windows(force_all=True)
                else:
                    # If dynamic is enabled but not managing all, ensure managed_by_script_hwnds is empty
                    self.managed_by_script_hwnds.clear()
            else:
                self.managed_by_script_hwnds.clear()

    def kill_script(self):
        """Failsafe hotkey to initiate a clean shutdown."""
        self.show_message("Kill hotkey pressed. Exiting script.", "red")
        self.on_closing()

    def on_closing(self):
        """Handles graceful shutdown when the main window is closed."""
        self.ahk.stop_hotkeys()

        if self.mouse_pos_timer:
            self.root.after_cancel(self.mouse_pos_timer)
            self.mouse_pos_timer = None

        if self.window_monitor_fg_timer:
            self.root.after_cancel(self.window_monitor_fg_timer)
            self.window_monitor_fg_timer = None

        if self.window_monitor_new_timer:
            self.root.after_cancel(self.window_monitor_new_timer)
            self.window_monitor_new_timer = None
        
        if self.window_monitor_inactivity_timer: # NEW: Cancel inactivity timer
            self.root.after_cancel(self.window_monitor_inactivity_timer)
            self.window_monitor_inactivity_timer = None

        self._stop_tooltip_follow()

        # Restore any dynamically transparent windows to full opacity before closing
        self._restore_managed_transparency_to_full_opacity()

        # NEW: Restore any windows minimized by the script to full size before closing
        for hwnd in list(self.minimized_by_script_hwnds):
            if win32gui.IsWindow(hwnd) and not self._is_window_excluded(hwnd): # Only restore if NOT excluded
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        self.minimized_by_script_hwnds.clear() # Clear the set after restoring/ignoring

        self.save_settings()
        self.root.destroy()

# --- Helper Functions (outside class for reusability) ---

def get_window_exe_name(hwnd):
    """
    Retrieves the executable name (e.g., 'notepad.exe') for a given window handle.
    Returns None if unable to retrieve.
    """
    pid = ctypes.c_ulong()
    GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    if pid.value == 0:
        return None

    process_handle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
    if not process_handle:
        process_handle = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid.value)
        if not process_handle:
            return None

    image_name_buffer = ctypes.create_unicode_buffer(260)
    buffer_size = ctypes.c_ulong(260)

    if QueryFullProcessImageNameW(process_handle, 0, image_name_buffer, ctypes.byref(buffer_size)):
        full_path = image_name_buffer.value
        base_name = os.path.splitext(os.path.basename(full_path))[0].lower()
        CloseHandle(process_handle)
        return base_name
    CloseHandle(process_handle)
    return None

def get_window_class_name(hwnd):
    """
    Retrieves the class name for a given window handle.
    Returns None if unable to retrieve.
    """
    try:
        return win32gui.GetClassName(hwnd)
    except win32gui.error:
        return None

def set_transparency_for_hwnd(hwnd, transparency_percentage):
    """
    Sets the transparency of a specific window using Windows API calls.
    Returns True on success, False on failure.
    """
    transparency_percentage = max(1, min(100, transparency_percentage))

    # Convert 1-100 percentage to 0-255 alpha value
    alpha = int(transparency_percentage * 2.55)

    # Enforce a minimum effective alpha value to prevent artifacting at very low transparencies.
    # A value of 15 corresponds roughly to 15 / 2.55 = ~5.88% transparency.
    # This helps avoid visual glitches that can occur when Windows tries to render
    # windows with near-zero alpha values.
    MIN_EFFECTIVE_ALPHA_VALUE = 15 
    alpha = max(MIN_EFFECTIVE_ALPHA_VALUE, min(255, alpha)) # Ensure alpha is within [MIN_EFFECTIVE_ALPHA_VALUE, 255]

    try:
        current_ex_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE)

        if not (current_ex_style & WS_EX_LAYERED):
            new_ex_style = current_ex_style | WS_EX_LAYERED
            SetWindowLongPtrW(hwnd, GWL_EXSTYLE, new_ex_style)

        success = SetLayeredWindowAttributes(hwnd, 0, alpha, LWA_ALPHA)
        return success
    except Exception as e:
        return False

def set_layered_window_colorkey_and_alpha(hwnd, colorkey_rgb, alpha_percentage):
    """
    Sets the transparency and colorkey for a layered window using Windows API.
    colorkey_rgb: BGR format (e.g., 0x00FF00 for green). This color will be made transparent.
    alpha_percentage: 1-100. This alpha will be applied to the *non-colorkey* parts.
    Returns True on success, False on failure.
    """
    alpha = int(alpha_percentage * 2.55)
    alpha = max(0, min(255, alpha))

    try:
        current_ex_style = GetWindowLongPtrW(hwnd, GWL_EXSTYLE)

        if not (current_ex_style & WS_EX_LAYERED):
            new_ex_style = current_ex_style | WS_EX_LAYERED
            SetWindowLongPtrW(hwnd, GWL_EXSTYLE, new_ex_style)

        success = SetLayeredWindowAttributes(hwnd, colorkey_rgb, alpha, LWA_COLORKEY | LWA_ALPHA)
        return success
    except Exception as e:
        print(f"Error setting layered window attributes for HWND {hwnd}: {e}")
        return False


if __name__ == "__main__":
    settings = DEFAULT_SETTINGS.copy()
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'rb') as f:
                loaded_settings = pickle.load(f)

            # ... (merge and cleanup logic) ...

            settings = loaded_settings
        except (EOFError, pickle.UnpickingError):
            # If file is corrupt or empty, use default settings
            settings = DEFAULT_SETTINGS.copy()
            pass 

    # Always save settings after loading/merging to ensure file is up-to-date with current structure
    # and deprecated keys are removed for next launch.
    with open(SETTINGS_FILE, 'wb') as f:
        pickle.dump(settings, f)

    root = customtkinter.CTk()
    app = TransparencyControllerApp(root)
    root.mainloop()