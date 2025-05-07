import tkinter as tk
from tkinter import scrolledtext, messagebox
import psutil
import threading
import time
import subprocess
import ctypes
import sys
import os
from ping3 import ping
from pystray import Icon, MenuItem, Menu
from PIL import Image
from datetime import datetime
import win10toast  # For Windows notifications

class NetworkMonitor:
    def __init__(self, app):
        self.monitoring = False
        self.logs = []
        self.active_adapter = None
        self.app = app  # Reference to the App instance
        self.last_ping_status = None  # Initialize to track last ping status

    def log(self, message):
        # Format log message with timestamp
        timestamp = datetime.now().strftime("%d/%m/%y - %H:%M")
        log_message = f"[{timestamp}] - {message}"
        self.logs.append(log_message)
        print(log_message)  # Output to console for debugging
        
        if self.app:  # Check if self.app is valid
            self.app.update_log_display(log_message)  # Update GUI log display
        else:
            print("App reference is not set.")  # Debugging line

    def get_wifi_adapter(self):
        """Fetches a list of active network adapters."""
        adapters = []
        for adapter in psutil.net_if_addrs().items():
            status = psutil.net_if_stats()[adapter[0]]  # Get status for each adapter
            if status.isup:  # Only include adapters that are up
                adapters.append(adapter[0])
        return adapters

    def ping_target(self, target='1.1.1.1'):
        """Ping a target and return True if reachable."""
        try:
            response = ping(target)
            current_ping_status = response is not None
            
            # Log the current ping status
            if current_ping_status and self.last_ping_status is False:
                self.log("Network is online.")
            elif not current_ping_status and self.last_ping_status is True:
                self.log("Network is offline.")

            self.last_ping_status = current_ping_status  # Update last ping status
            return current_ping_status
        except Exception as e:
            self.log(f"An error occurred while pinging: {e}")
            return False

    def toggle_adapter(self, action, adapter):
        if adapter is None:
            self.log("Adapter is None. Cannot perform action.")
            return

        # Use PowerShell commands to enable or disable the adapter
        command = f"PowerShell -Command {action}-NetAdapter -Name '{adapter}' -Confirm:$false"
        self.log(f"Executing command: {command}")
        
        # Execute the PowerShell command
        result = subprocess.call(command, shell=True)
        if result == 0:  # Check if command executed successfully
            self.log(f"{action.capitalize()}d the network adapter: {adapter}")
        else:
            self.log(f"Failed to {action} adapter: {adapter}")

    def check_network(self):
        offline_start_time = None  # Time tracking for offline state
        
        while self.monitoring:
            if not self.ping_target():  # If the network is offline
                if offline_start_time is None:
                    # Start the timer when pings first fail
                    offline_start_time = time.time()
                    self.log("Network is offline. Starting timer...")
                
                # Calculate total time offline
                time_offline = time.time() - offline_start_time
                
                if time_offline >= 15:  # If offline for 15 seconds
                    self.log("Network is offline for 15 seconds. Restarting adapter...")
                    self.toggle_adapter('Disable', self.active_adapter)
                    time.sleep(5)  # Wait for a few seconds
                    self.toggle_adapter('Enable', self.active_adapter)

                    # Reset the timer after trying to restart
                    offline_start_time = time.time()  # Start timing for reconnection attempts
                    self.log("Attempting to reestablish connectivity...")
                    
                    # Attempt to reestablish connectivity
                    timeout_count = 0
                    while timeout_count < 3:  # Allow up to 15 seconds for connection restore
                        time.sleep(5)  # Wait 5 seconds before next ping
                        if self.ping_target():  # Check again for connectivity
                            self.log("Connection restored.")
                            break
                        timeout_count += 1
                    else:
                        self.log("Unable to restore the connection after adapter restart.")
                        self.show_notification("Network Issue", "Unable to restore connection after adapter restart.")
            else:
                offline_start_time = None  # Reset timer if network is back online

            time.sleep(1)  # Delay before the next check

    def start_monitoring(self):
        self.monitoring = True
        self.log("Starting network monitoring...")
        threading.Thread(target=self.check_network, daemon=True).start()

    def stop_monitoring(self):
        self.monitoring = False
        self.log("Monitoring stopped.")

    def show_notification(self, title, message):
        """Show a Windows notification."""
        toaster = win10toast.ToastNotifier()
        toaster.show_toast(title, message, duration=10)

    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except Exception as e:
            self.log(f"Error checking admin status: {e}")
            return False

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Network Monitor")

        self.network_monitor = NetworkMonitor(self)  # Pass reference of App to NetworkMonitor

        # Adapter selection dropdown
        self.adapter_label = tk.Label(root, text="Select Adapter:")
        self.adapter_label.pack(pady=10)

        self.adapter_var = tk.StringVar(root)
        self.adapter_dropdown = tk.OptionMenu(root, self.adapter_var, *self.fetch_adapters())
        self.adapter_dropdown.pack(pady=10)

        # Start and Stop buttons
        self.start_button = tk.Button(root, text="Start Monitoring", command=self.start_monitoring)
        self.start_button.pack(pady=10)

        self.stop_button = tk.Button(root, text="Stop Monitoring", command=self.stop_monitoring)
        self.stop_button.pack(pady=10)

        self.log_area = scrolledtext.ScrolledText(root, width=50, height=15)
        self.log_area.pack(pady=10)

        # Handle close event in the GUI
        self.root.protocol("WM_DELETE_WINDOW", self.hide)

        # Initialize tray icon
        self.tray_icon = Icon("NetworkMonitor")
        self.tray_icon.title = "Network Monitor"
        self.tray_icon.icon = self.create_image()

        # Create tray menu
        self.tray_icon.menu = Menu(
            MenuItem("Show", self.show),  # Restore GUI
            MenuItem("Start", self.start_monitoring),
            MenuItem("Stop", self.stop_monitoring),
            MenuItem("Exit", self.exit_app)
        )

        # Start the icon in the tray (run in a separate thread)
        threading.Thread(target=self.tray_icon.run, args=(self.setup_tray,), daemon=True).start()

    def fetch_adapters(self):
        adapters = self.network_monitor.get_wifi_adapter()
        if adapters:
            self.network_monitor.active_adapter = adapters[0]  # Set the first adapter as active
            return adapters
        else:
            messagebox.showerror("Error", "No active network adapters found.")
            return []

    def setup_tray(self, icon):
        icon.visible = True  # Ensure tray icon is visible

    def create_image(self):
        """Create and return the icon image for the system tray."""
        if getattr(sys, 'frozen', False):  # Check if running as a compiled executable
            icon_path = os.path.join(sys._MEIPASS, 'wifi.ico')  # _MEIPASS is the path where bundled files are extracted
        else:
            icon_path = "wifi.ico"  # Default path for the script run
        return Image.open(icon_path)  # Load and return the image

    def start_monitoring(self):
        self.network_monitor.active_adapter = self.adapter_var.get()
        if not self.network_monitor.active_adapter:
            messagebox.showerror("Error", "Please select a network adapter.")
            return
        self.network_monitor.start_monitoring()
        self.update_log_display("Monitoring started on adapter: " + self.network_monitor.active_adapter)

    def stop_monitoring(self):
        self.network_monitor.stop_monitoring()
        self.update_log_display("Monitoring stopped.")

    def hide(self):
        self.root.withdraw()  # Hide the window

    def show(self):
        self.root.deiconify()  # Restore the window from the tray

    def exit_app(self):
        self.network_monitor.stop_monitoring()  # Stop monitoring
        self.tray_icon.stop()  # Stop the tray icon
        self.root.quit()  # Quit the application

    def update_log_display(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)

if __name__ == "__main__":
    # Request admin permissions
    network_monitor = NetworkMonitor(None)  # Pass None for admin checking
    if not network_monitor.is_admin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    else:
        root = tk.Tk()
        app = App(root)  # Pass root window to the App instance
        root.deiconify()  # Ensure the GUI is shown on startup
        root.mainloop()