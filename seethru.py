import subprocess
import time
import win32gui
import win32con
import keyboard
import sys
import threading
import pystray
from PIL import Image
import os
import json
from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QFileDialog, QMessageBox, QComboBox
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon

class TransparentApp:
    def __init__(self):
        self.hwnd = None  # Browser window handle
        self.is_transparent = False
        self.is_click_through = True
        self.current_opacity = 150
        self.is_fully_transparent = False
        self.saved_opacity = self.current_opacity
        self.running = False
        self.thread = None

    def find_window(self, title_contains):
        """Find the window by partial title match."""
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title_contains.lower() in title.lower():
                    windows.append(hwnd)
        windows = []
        win32gui.EnumWindows(callback, windows)
        return windows[0] if windows else None

    def set_window_transparent(self, hwnd, opacity=150, click_through=False):
        """Set the window to be transparent and optionally click-through."""
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        ex_style |= win32con.WS_EX_LAYERED
        if click_through:
            ex_style |= win32con.WS_EX_TRANSPARENT
        else:
            ex_style &= ~win32con.WS_EX_TRANSPARENT
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style)
        win32gui.SetLayeredWindowAttributes(hwnd, 0, opacity, win32con.LWA_ALPHA)
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    def disable_transparency(self, hwnd):
        """Disable transparency and click-through."""
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        ex_style &= ~win32con.WS_EX_LAYERED & ~win32con.WS_EX_TRANSPARENT
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style)
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    def restore_window(self, hwnd):
        """Restore the window to its normal state."""
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        ex_style &= ~win32con.WS_EX_LAYERED & ~win32con.WS_EX_TRANSPARENT
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex_style)
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

    def run_transparency(self, browser_path, browser_title):
        """Run the transparency logic for the browser window."""
        # Launch the browser
        try:
            subprocess.run([browser_path, ""], check=True)
        except Exception as e:
            print(f"Failed to launch browser: {e}")
            return

        time.sleep(5)
        # Find the browser window
        self.hwnd = self.find_window(browser_title)
        if not self.hwnd:
            print(f"Could not find browser window with title containing '{browser_title}'")
            return

        self.is_transparent = True
        self.is_click_through = True
        self.current_opacity = 150
        self.saved_opacity = self.current_opacity
        self.running = True

        self.set_window_transparent(self.hwnd, opacity=self.current_opacity, click_through=self.is_click_through)

        print("Browser window is transparent and click-through. Use Alt+Tab to switch to the browser window.")
        print("Hotkeys (active when browser window is focused):")
        print("- Win+Z: Toggle interactive/click-through")
        print("- Win+X: Toggle transparency on/off")
        print("- Win+A: Decrease transparency (increase opacity)")
        print("- Win+D: Increase transparency (decrease opacity)")
        print("- Win+W: Toggle full transparency")
        print("- Win+Q: Stop and exit")

        def toggle_click_through():
            if not self.running:
                return
            self.is_click_through = not self.is_click_through
            self.set_window_transparent(self.hwnd, opacity=self.current_opacity, click_through=self.is_click_through)
            print(f"Browser window is now {'click-through' if self.is_click_through else 'interactive'}")

        def toggle_transparency():
            if not self.running:
                return
            self.is_transparent = not self.is_transparent
            if self.is_transparent:
                self.current_opacity = self.saved_opacity if self.is_fully_transparent else self.current_opacity
                self.is_fully_transparent = False
                self.set_window_transparent(self.hwnd, opacity=self.current_opacity, click_through=self.is_click_through)
                print(f"Transparency enabled (opacity: {self.current_opacity})")
            else:
                self.disable_transparency(self.hwnd)
                print("Transparency disabled")

        def increase_opacity():
            if not self.running:
                return
            if self.is_fully_transparent:
                self.is_fully_transparent = False
                self.current_opacity = self.saved_opacity
            self.current_opacity = min(255, self.current_opacity + 10)
            self.saved_opacity = self.current_opacity
            if self.is_transparent:
                self.set_window_transparent(self.hwnd, opacity=self.current_opacity, click_through=self.is_click_through)
                print(f"Opacity increased to {self.current_opacity}")

        def decrease_opacity():
            if not self.running:
                return
            if self.is_fully_transparent:
                self.is_fully_transparent = False
                self.current_opacity = self.saved_opacity
            self.current_opacity = max(0, self.current_opacity - 10)
            self.saved_opacity = self.current_opacity
            if self.is_transparent:
                self.set_window_transparent(self.hwnd, opacity=self.current_opacity, click_through=self.is_click_through)
                print(f"Opacity decreased to {self.current_opacity}")

        def toggle_full_transparency():
            if not self.running:
                return
            self.is_fully_transparent = not self.is_fully_transparent
            if self.is_fully_transparent:
                self.saved_opacity = self.current_opacity
                self.current_opacity = 0
            else:
                self.current_opacity = self.saved_opacity
            if self.is_transparent:
                self.set_window_transparent(self.hwnd, opacity=self.current_opacity, click_through=self.is_click_through)
                print(f"{'Fully transparent' if self.is_fully_transparent else f'Restored opacity to {self.current_opacity}'}")

        # Register hotkeys
        keyboard.add_hotkey("win+z", toggle_click_through)
        keyboard.add_hotkey("win+x", toggle_transparency)
        keyboard.add_hotkey("win+a", increase_opacity)
        keyboard.add_hotkey("win+d", decrease_opacity)
        keyboard.add_hotkey("win+w", toggle_full_transparency)
        keyboard.add_hotkey("win+q", lambda: None)  # Exit handled by keyboard.wait

        keyboard.wait("win+q")
        self.running = False
        if self.hwnd:
            self.restore_window(self.hwnd)
            print("Transparency and click-through disabled.")
            self.hwnd = None

    def start_transparency(self, browser_path, browser_title):
        """Start the transparency logic in a separate thread."""
        if self.running:
            print("Transparency is already running.")
            return
        self.thread = threading.Thread(target=self.run_transparency, args=(browser_path, browser_title))
        self.thread.daemon = True
        self.thread.start()

    def stop_transparency(self):
        """Stop the transparency effect and restore the window."""
        if self.running and self.hwnd:
            self.running = False
            self.restore_window(self.hwnd)
            print("Transparency and click-through disabled.")
            self.hwnd = None

class AppGUI(QWidget):
    def __init__(self, transparent_app, tray_icon):
        super().__init__()
        self.transparent_app = transparent_app
        self.tray_icon = tray_icon
        self.setWindowTitle("Transparent Browser Overlay")
        self.setFixedSize(400, 250)

        # Load saved settings
        self.settings_file = "settings.json"
        self.saved_settings = self.load_settings()

        # Layout
        layout = QVBoxLayout()

        # Browser Executable
        browser_path_label = QLabel("Browser Executable (e.g., Chrome):")
        layout.addWidget(browser_path_label)
        self.browser_path_edit = QLineEdit()
        layout.addWidget(self.browser_path_edit)
        browser_browse_button = QPushButton("Browse")
        browser_browse_button.clicked.connect(self.browse_browser)
        layout.addWidget(browser_browse_button)

        # Browser Window Title Substring
        browser_title_label = QLabel("Browser Window Title Substring (e.g., Google Chrome):")
        layout.addWidget(browser_title_label)
        self.browser_title_edit = QLineEdit()
        layout.addWidget(self.browser_title_edit)

        # Saved Settings Dropdown
        saved_label = QLabel("Saved Settings:")
        layout.addWidget(saved_label)
        self.combo_box = QComboBox()
        self.combo_box.currentTextChanged.connect(self.on_select_setting)
        layout.addWidget(self.combo_box)
        self.update_combo_box()

        # Start Button
        start_button = QPushButton("Start Transparency")
        start_button.clicked.connect(self.start)
        layout.addWidget(start_button)

        self.setLayout(layout)

        # Center the window
        self.center()

    def center(self):
        """Center the window on the screen."""
        screen = QApplication.primaryScreen().geometry()
        self.move((screen.width() - self.width()) // 2, (screen.height() - self.height()) // 2)

    def load_settings(self):
        """Load saved settings from settings.json, handling old format."""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                # Filter and convert settings
                valid_settings = []
                for s in settings:
                    if 'browser_path' in s and 'browser_title' in s:
                        valid_settings.append(s)
                    elif 'path' in s and 'title' in s:
                        # Convert old format to new format
                        valid_settings.append({
                            'browser_path': s['path'],
                            'browser_title': s['title']
                        })
                return valid_settings
        except Exception as e:
            print(f"Error loading settings: {e}")
        return []

    def save_settings(self, browser_path, browser_title):
        """Save the current settings to settings.json."""
        if not any(s['browser_path'] == browser_path and s['browser_title'] == browser_title for s in self.saved_settings):
            self.saved_settings.append({
                "browser_path": browser_path,
                "browser_title": browser_title
            })
            with open(self.settings_file, 'w') as f:
                json.dump(self.saved_settings, f, indent=4)
            self.update_combo_box()

    def update_combo_box(self):
        """Update the combo box with saved settings."""
        self.combo_box.clear()
        options = [f"{os.path.basename(s['browser_path'])} ({s['browser_title']})" for s in self.saved_settings]
        if not options:
            options = ["No saved settings"]
        self.combo_box.addItems(options)
        if self.saved_settings and len(self.saved_settings) > 0:
            self.browser_path_edit.setText(self.saved_settings[0]['browser_path'])
            self.browser_title_edit.setText(self.saved_settings[0]['browser_title'])
            self.combo_box.setCurrentText(f"{os.path.basename(self.saved_settings[0]['browser_path'])} ({self.saved_settings[0]['browser_title']})")

    def on_select_setting(self, value):
        """Populate entry fields when a saved setting is selected."""
        if value and value != "No saved settings":
            for setting in self.saved_settings:
                if f"{os.path.basename(setting['browser_path'])} ({setting['browser_title']})" == value:
                    self.browser_path_edit.setText(setting['browser_path'])
                    self.browser_title_edit.setText(setting['browser_title'])
                    break
        else:
            self.browser_path_edit.clear()
            self.browser_title_edit.clear()

    def browse_browser(self):
        """Open file dialog to select the browser executable."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Browser Executable", "", "Executable files (*.exe)")
        if file_path:
            self.browser_path_edit.setText(file_path)

    def start(self):
        """Start transparency and save settings."""
        browser_path = self.browser_path_edit.text().strip()
        browser_title = self.browser_title_edit.text().strip()
        if not browser_path or not browser_title:
            QMessageBox.critical(self, "Error", "Please specify both the browser executable path and window title substring.")
            return
        if not os.path.isfile(browser_path):
            QMessageBox.critical(self, "Error", "The specified browser executable path does not exist.")
            return
        self.save_settings(browser_path, browser_title)
        self.transparent_app.start_transparency(browser_path, browser_title)
        self.hide()

def create_system_tray():
    try:
        icon_image = Image.open("icon.ico")
    except:
        icon_image = Image.new("RGB", (64, 64), color="blue")

    app = QApplication(sys.argv)
    transparent_app = TransparentApp()
    gui = AppGUI(transparent_app, None)

    def show_gui(icon, item):
        gui.show()

    def stop_transparency(icon, item):
        transparent_app.stop_transparency()

    def exit_app(icon, item):
        transparent_app.stop_transparency()
        icon.stop()
        app.quit()

    menu = pystray.Menu(
        pystray.MenuItem("Open GUI", show_gui),
        pystray.MenuItem("Stop Transparency", stop_transparency),
        pystray.MenuItem("Exit", exit_app)
    )
    icon = pystray.Icon("Transparent App", icon_image, "Transparent App", menu)
    gui.tray_icon = icon

    tray_thread = threading.Thread(target=icon.run, daemon=True)
    tray_thread.start()

    gui.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    create_system_tray()