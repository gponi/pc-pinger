import os
import datetime
import json
import concurrent.futures
from collections import defaultdict
import pandas as pd

# 1. Function to read computers from a text file
def read_computer_list(file_path='computers.txt'):
    try:
        with open(file_path, 'r') as f:
            computers = [line.strip() for line in f if line.strip()]
        return computers
    except FileNotFoundError:
        return []

# 2. Function to ping a single computer (used in multithreading)
def ping_single_computer(computer):
    try:
        response = os.system(f"ping -n 1 {computer}")
        current_day = datetime.datetime.now().strftime("%A")  # Get the current day of the week
        if response == 0:
            return {'computer': computer, 'day': current_day, 'status': 'online'}
        else:
            return {'computer': computer, 'day': current_day, 'status': 'offline'}
    except Exception as e:
        return {'computer': computer, 'day': 'unknown', 'status': f'error: {str(e)}'}

# 3. Function to ping computers concurrently using multithreading
def ping_computers_multithreaded(computer_list, output_file='ping_results.json'):
    results = defaultdict(list)
    current_day = datetime.datetime.now().strftime("%A")  # Store current day once
    
    # Use ThreadPoolExecutor to ping multiple computers at the same time
    with concurrent.futures.ThreadPoolExecutor() as executor:
        ping_results = executor.map(ping_single_computer, computer_list)
    
    # Collect results from the executor
    for result in ping_results:
        results[result['computer']].append({'day': result['day'], 'status': result['status']})
    
    # Handle JSON loading or file initialization
    try:
        with open(output_file, 'r') as f:
            previous_data = json.load(f) if f.read().strip() else {}
    except (FileNotFoundError, json.JSONDecodeError):
        previous_data = {}  # Initialize empty if the file doesn't exist or is invalid
    
    # Update the previous data with the new results (without duplicates for the same day)
    for computer, new_logs in results.items():
        if computer not in previous_data:
            previous_data[computer] = new_logs
        else:
            # Check if there's already an entry for the current day
            existing_days = {entry['day'] for entry in previous_data[computer]}
            for log in new_logs:
                if log['day'] not in existing_days:
                    previous_data[computer].append(log)
    
    # Save the updated results
    with open(output_file, 'w') as f:
        json.dump(previous_data, f, indent=4)
    
    print(f"Ping results saved to {output_file}.")

# 4. Function to convert the ping results into an Excel file
def generate_excel_report(input_file='ping_results.json', output_file='network_days_report.xlsx'):
    try:
        with open(input_file, 'r') as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return
    
    pc_days_map = {}
    day_abbreviations = {
        "Monday": "M", "Tuesday": "T", "Wednesday": "W", 
        "Thursday": "Th", "Friday": "F"
    }
    
    for computer, logs in data.items():
        days_connected = sorted({day_abbreviations[log['day']] for log in logs if log['status'] == 'online'})
        pc_days_map[computer] = ','.join(days_connected)
    
    # Convert the data into a DataFrame and write to an Excel file
    df = pd.DataFrame(list(pc_days_map.items()), columns=['PC Name', 'Days Connected'])
    df.to_excel(output_file, index=False)
    
    print(f"Excel report generated: {output_file}")

# Example usage:
if __name__ == "__main__":
    # Step 1: Read computer list from a text file
    computers = read_computer_list('computers.txt')
    
    if computers:
        # Step 2: Ping computers concurrently (faster)
        ping_computers_multithreaded(computers)
    
        # Step 3: Generate Excel report with the days users have connected
        generate_excel_report()
