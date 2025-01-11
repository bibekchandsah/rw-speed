import os
import time
import psutil
import win32api
import win32com.client
import subprocess
from collections import defaultdict

def list_physical_drives():
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

def get_drive_name(device):
    """Get the name of the physical drive."""
    try:
        volume_name = win32api.GetVolumeInformation(device)[0]
        return volume_name if volume_name else "Unnamed Volume"
    except:
        return "Unknown"

def write_speed_test(file_path, file_size):
    """Test the write speed."""
    data = os.urandom(file_size)  # Generate random data of the specified size
    start_time = time.time()
    with open(file_path, 'wb') as f:
        f.write(data)
    end_time = time.time()
    write_speed = file_size / (end_time - start_time) / (1024 * 1024)  # Convert to MB/s
    return write_speed

def read_speed_test(file_path):
    """Test the read speed."""
    start_time = time.time()
    with open(file_path, 'rb') as f:
        data = f.read()
    end_time = time.time()
    read_speed = len(data) / (end_time - start_time) / (1024 * 1024)  # Convert to MB/s
    return read_speed

def format_speed(speed):
    """Format the speed for display."""
    if speed < 1024:
        return f"{speed:.2f} MB/s"
    else:
        return f"{speed / 1024:.2f} GB/s"

def get_drive_rpm(drive_id):
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

def get_smart_data(drive):
    """Retrieve S.M.A.R.T. data for the given drive."""
    try:
        result = subprocess.run(["smartctl", "-a", drive], capture_output=True, text=True, check=True)
        return result.stdout
    except FileNotFoundError:
        print(f"smartctl not found. Make sure smartmontools is installed and in your PATH.")
        return None
    except subprocess.CalledProcessError as e:
        print(f"Failed to retrieve S.M.A.R.T. data for {drive}: {e}")
        return None

def parse_smart_data(smart_data):
    """Parse the S.M.A.R.T. data to extract usage hours and connection count."""
    usage_hours = None
    connection_count = None
    
    for line in smart_data.splitlines():
        if "Power_On_Hours" in line:
            usage_hours = int(line.split()[-1])
        elif "Power_Cycle_Count" in line:
            connection_count = int(line.split()[-1])
    
    return usage_hours, connection_count

def test_drive_speed(drive, partition):
    file_path = os.path.join(partition, 'temp', 'test_file.bin')
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    file_size = 1024 * 1024 * 100  # 100 MB

    print(f"\nTesting write speed for {partition} ({drive})...")
    try:
        write_speed = write_speed_test(file_path, file_size)
        print(f"Write speed: {format_speed(write_speed)}")
    except PermissionError:
        print("Permission denied. Unable to write to the selected partition.")
        return

    print("Testing read speed...")
    read_speed = read_speed_test(file_path)
    print(f"Read speed: {format_speed(read_speed)}")

    os.remove(file_path)  # Clean up

    rpm = get_drive_rpm(drive)
    print(f"Drive RPM: {rpm}")

    smart_data = get_smart_data(f"//./{partition}")
    if smart_data:
        usage_hours, connection_count = parse_smart_data(smart_data)
        print(f"Usage Hours: {usage_hours} hours")
        print(f"Connection Count: {connection_count} times")

def main():
    while True:
        drive_partitions = list_physical_drives()
        if not drive_partitions:
            print("No storage devices found.")
            return

        print("Available storage devices:")
        drive_mapping = []
        for i, (drive_id, partitions) in enumerate(drive_partitions.items()):
            drive_name = get_drive_name(partitions[0])
            for partition in partitions:
                print(f"{len(drive_mapping) + 1}: {partition} ({drive_name})")
                drive_mapping.append((drive_id, partition))
        print(f"{len(drive_mapping) + 1}: Test all drives")
        print(f"{len(drive_mapping) + 2}: Exit")

        choice = input("Select a storage device to test (by number): ").strip()
        if not choice.isdigit() or int(choice) - 1 not in range(len(drive_mapping) + 2):
            print("Invalid choice. Please select a valid option.")
            continue

        choice = int(choice) - 1
        if choice == len(drive_mapping):
            print("Testing all drives...")
            for drive, partition in drive_mapping:
                test_drive_speed(drive, partition)
        elif choice == len(drive_mapping) + 1:
            print("Exiting the program.")
            break
        else:
            selected_drive, selected_partition = drive_mapping[choice]
            print(f"Selected device: {selected_partition} ({selected_drive})")
            test_drive_speed(selected_drive, selected_partition)

        repeat = input("Do you want to test another storage device? (yes/no): ").strip().lower()
        if repeat != 'yes':
            break

if __name__ == "__main__":
    main()

