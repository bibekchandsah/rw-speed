import os
import time
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox
import win32api
import win32com.client
from collections import defaultdict

class DiskSpeedTesterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Disk Speed Tester")
        
        # Center the window on the screen
        window_width = 600
        window_height = 400
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width // 2) - (window_width // 2)
        y = (screen_height // 2) - (window_height // 2)
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # Initialize UI components
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        
        self.label_select = ttk.Label(self.main_frame, text="Select a storage device to test:")
        self.label_select.grid(row=0, column=0, columnspan=2, pady=10)
        
        self.device_combobox = ttk.Combobox(self.main_frame, state="readonly", width=50)
        self.device_combobox.grid(row=1, column=0, columnspan=2, pady=5)
        
        self.button_test = ttk.Button(self.main_frame, text="Test Speed", command=self.test_speed)
        self.button_test.grid(row=2, column=0, pady=10)
        
        self.button_test_all = ttk.Button(self.main_frame, text="Test All Drives", command=self.test_all_drives)
        self.button_test_all.grid(row=2, column=1, pady=10)
        
        self.result_frame = ttk.Frame(self.root, padding="20")
        self.result_frame.grid(row=1, column=0, sticky="nsew")
        
        self.label_result = ttk.Label(self.result_frame, text="Test Results:", font=("Arial", 14, "bold"))
        self.label_result.grid(row=0, column=0, columnspan=4, pady=10)
        
        self.result_labels = []
        self.result_values = []
        
        # Labels for displaying results dynamically
        labels = ["Disk", "Read Speed", "Write Speed", "RPM"]
        for i, label in enumerate(labels):
            ttk.Label(self.result_frame, text=label).grid(row=1, column=i, padx=10, pady=5)
            self.result_labels.append(label)
            self.result_values.append(tk.StringVar())
            ttk.Label(self.result_frame, textvariable=self.result_values[i]).grid(row=2, column=i, padx=10, pady=5)
        
        # Populate device list
        self.populate_device_list()

    def populate_device_list(self):
        drive_partitions = self.list_physical_drives()
        if not drive_partitions:
            messagebox.showerror("Error", "No storage devices found.")
            self.root.quit()
            return
        
        self.drive_mapping = []
        for drive_id, partitions in drive_partitions.items():
            drive_name = self.get_drive_name(partitions[0])
            for partition in partitions:
                self.drive_mapping.append((drive_id, partition, drive_name))
                self.device_combobox["values"] = [f"{partition} ({drive_name})" for drive_id, partition, drive_name in self.drive_mapping]
    
    def list_physical_drives(self):
        """List all physical drives and their corresponding partitions."""
        wmi = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        service = wmi.ConnectServer(".", "root\\cimv2")
        drives = service.ExecQuery("SELECT * FROM Win32_DiskDrive")
        
        drive_partitions = defaultdict(list)
        for drive in drives:
            for partition in drive.Associators_("Win32_DiskDriveToDiskPartition"):
                for logical_disk in partition.Associators_("Win32_LogicalDiskToPartition"):
                    drive_partitions[drive.DeviceID].append(logical_disk.DeviceID)
        
        return drive_partitions

    def get_drive_name(self, device):
        """Get the name of the physical drive."""
        try:
            volume_name = win32api.GetVolumeInformation(device)[0]
            return volume_name if volume_name else "Unnamed Volume"
        except:
            return "Unknown"

    def test_speed(self):
        selected_device_index = self.device_combobox.current()
        if selected_device_index == -1:
            messagebox.showerror("Error", "Please select a device to test.")
            return
        
        selected_drive, selected_partition, selected_drive_name = self.drive_mapping[selected_device_index]
        
        # Display "Calculating..." message
        self.result_values[0].set(selected_partition)
        for i in range(1, 4):
            self.result_values[i].set("Calculating...")
        
        # Perform speed tests
        file_path = os.path.join(selected_partition, 'temp', 'test_file.bin')
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        file_size = 1024 * 1024 * 100  # 100 MB
        
        try:
            write_speed = self.write_speed_test(file_path, file_size)
            self.result_values[1].set(self.format_speed(self.read_speed_test(file_path)))
            self.result_values[2].set(self.format_speed(write_speed))
            self.result_values[3].set(self.get_drive_rpm(selected_drive))
        except PermissionError:
            messagebox.showerror("Error", "Permission denied. Unable to write to the selected partition.")
            for i in range(1, 4):
                self.result_values[i].set("N/A")
            return
        
        os.remove(file_path)  # Clean up
    
    def write_speed_test(self, file_path, file_size):
        """Test the write speed."""
        data = os.urandom(file_size)  # Generate random data of the specified size
        start_time = time.time()
        with open(file_path, 'wb') as f:
            f.write(data)
        end_time = time.time()
        write_speed = file_size / (end_time - start_time) / (1024 * 1024)  # Convert to MB/s
        return write_speed

    def read_speed_test(self, file_path):
        """Test the read speed."""
        start_time = time.time()
        with open(file_path, 'rb') as f:
            data = f.read()
        end_time = time.time()
        read_speed = len(data) / (end_time - start_time) / (1024 * 1024)  # Convert to MB/s
        return read_speed
    
    def format_speed(self, speed):
        """Format the speed for display."""
        if speed < 1024:
            return f"{speed:.2f} MB/s"
        else:
            return f"{speed / 1024:.2f} GB/s"
    
    def get_drive_rpm(self, drive_id):
        """Get the RPM of the physical drive."""
        try:
            wmi = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            service = wmi.ConnectServer(".", "root\\cimv2")
            query = f"SELECT * FROM Win32_DiskDrive WHERE DeviceID = '{drive_id}'"
            drive = service.ExecQuery(query)[0]
            rpm = drive.SpindleSpeed
            return rpm if rpm else "N/A"
        except:
            return "Unknown"

    def test_all_drives(self):
        """Test all drives and display results."""
        for drive_id, partition, drive_name in self.drive_mapping:
            self.result_values[0].set(partition)
            for i in range(1, 4):
                self.result_values[i].set("Calculating...")
            self.root.update()  # Update the main window to show "Calculating..." message
            
            file_path = os.path.join(partition, 'temp', 'test_file.bin')
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            file_size = 1024 * 1024 * 100  # 100 MB
            
            try:
                write_speed = self.write_speed_test(file_path, file_size)
                self.result_values[1].set(self.format_speed(self.read_speed_test(file_path)))
                self.result_values[2].set(self.format_speed(write_speed))
                self.result_values[3].set(self.get_drive_rpm(drive_id))
            except PermissionError:
                for i in range(1, 4):
                    self.result_values[i].set("N/A")
            finally:
                os.remove(file_path)  # Clean up
                time.sleep(1)  # Add a delay for visibility
                
    def exit(self):
        self.root.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = DiskSpeedTesterApp(root)
    root.mainloop()
