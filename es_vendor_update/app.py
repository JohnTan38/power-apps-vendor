import os
from datetime import datetime

def write_to_txt(directory):
    try:
        with open(directory, 'w') as file:
            file.write('task schedule at ' + datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    directory = r"C:/Users/john.tan/Documents/power_apps_esker_vendor/esker_vendor_update/task_schedule.txt"
    write_to_txt(directory)
    print(f"Task schedule written to {directory}")
