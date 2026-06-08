from flask import Flask, render_template, request, jsonify, send_file
from flask_socketio import SocketIO, emit, join_room
import os
import sys
import threading
import time
import logging
from werkzeug.utils import secure_filename
import pandas as pd
import paramiko
import re
import concurrent.futures
import random
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
import io
import uuid
from pathlib import Path

# Add reportgen to path
sys.path.append(os.path.join(os.path.dirname(__file__), 'reportgen'))
import generate_report
import enterprise_report


app = Flask(__name__)
app.secret_key = 'gpon_assistant_secret_key_change_me'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading')

sessions = {}  # {session_id: {'driver': None, 'wait': None, 'available_categories': [], 'paused': False, 'stopped': False}}

logging.basicConfig(level=logging.INFO, filename='app.log', filemode='a',
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Category options
los_category_options = [
    "B.C. Enclosure Maintenance-GPON-Service Affecting-LOS Fiber Cut-Cable Maintenance Incident",
    "Private Developer-GPON-Service Affecting-LOS Fiber Cut-Cable Maintenance Incident",
    "Monkeys-GPON-Service Affecting-LOS Fiber Cut-Cable Maintenance Incident",
    "Damage of OH cable by Lorry-GPON-Service Affecting-LOS Fiber Cut-Cable Maintenance Incident",
    "B.C. Rodents-GPON-Service Affecting-LOS Fiber Cut-Cable Maintenance Incident"
]
lowrx_category_options = [
    "Low Rx Optimisation-GPON-Service Affecting-LOS Fiber Cut-Cable Maintenance Incident",
]

# Category options for Low Rx and Extreme Low Rx transitions
lowrx_transition_category = "Low Rx-GPON-Service Affecting-Low RX-Cable Maintenance Incident"
extreme_lowrx_transition_category = "Extreme Low Rx-GPON-Service Affecting-Extreme Low RX-Cable Maintenance Incident"

def get_non_repeating_category(available, options):
    if not available:
        available[:] = options[:]
        random.shuffle(available)
    return available.pop(0)

def print_message(message, session_id):
    logging.info(f"[{session_id}] {message}")
    socketio.emit('log_update', {'message': message}, room=session_id)

def update_progress(value, max_value, message, session_id):
    percentage = (value / max_value * 100) if max_value else 0
    socketio.emit('progress_update', {'percentage': percentage, 'message': message}, room=session_id)

def read_olt_data(session_id):
    file_path = "OLT_DATA1.xlsx"
    if not os.path.exists(file_path):
        print_message("Error: OLT_DATA1.xlsx not found in root directory!", session_id)
        return {}
    try:
        df = pd.read_excel(file_path)
        ne_data = {}
        for _, row in df.iterrows():
            ne = str(row["NE"]).strip().lower()
            ne_type = row["Type"]
            host = row["Host"]
            ne_data[ne] = (ne_type, host)
        print_message("OLT data loaded successfully.", session_id)
        return ne_data
    except Exception as e:
        print_message(f"Error reading OLT data: {str(e)}", session_id)
        return {}

def login_and_process(username, password, session_id):
    try:
        chrome_options = Options()
        #chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--ignore-certificate-errors")
        chrome_options.add_argument("--disable-gpu")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        #driver.set_window_size(1440, 900)  # Not needed when maximized
        wait = WebDriverWait(driver, 5)
        sessions[session_id]['driver'] = driver
        sessions[session_id]['wait'] = wait
        sessions[session_id]['available_categories'] = los_category_options[:]  # Default for LOS/lowrx
        sessions[session_id]['paused'] = False
        sessions[session_id]['stopped'] = False

        url = "https://intranet.jtl.co.ke/login"
        driver.get(url)
        username_input = wait.until(EC.presence_of_element_located((By.ID, "login_email")))
        password_input = wait.until(EC.presence_of_element_located((By.ID, "login_password")))
        username_input.send_keys(username)
        password_input.send_keys(password)
        login_button = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button.btn.btn-sm.btn-default.btn-block.btn-login.btn-ldap-login")))
        login_button.click()
        wait.until(lambda d: "login" not in d.current_url.lower())

        # element = wait.until(EC.presence_of_element_located((By.XPATH, "//a[@href='/app/support']")))
        # element.click()
        # wait.until(lambda d: "/app/support" in d.current_url)
        # element2 = wait.until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Issue')]")))
        # element2.click()
        # wait.until(lambda d: "issue" in d.current_url.lower())

        # try:
        #     element3 = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "use[href='#es-small-close']")))
        #     element3.click()
        # except:
        #     pass

        print_message("CRM login successful. Waiting 5 seconds before processing...", session_id)
        time.sleep(5)  # Wait 5 seconds after login before starting processing
        print_message("Ready to process tickets.", session_id)
        return True
    except Exception as e:
        print_message(f"CRM login failed: {str(e)}", session_id)
        if 'driver' in sessions[session_id]:
            sessions[session_id]['driver'].quit()
            del sessions[session_id]['driver']
        return False

def extract_ne_and_location(text):
    pattern = r"(.+?)\s+(\d+/\d{1,2}:\d{1,2})"
    match = re.search(pattern, text)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return None, None

def ssh_task(row_index, ticket, service_id, ne, location, ne_type, host, ssh_email, ssh_password, subject, session_id, mode='los'):
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        client.connect(host, port=22, username=ssh_email, password=ssh_password)
        shell = client.invoke_shell()
        time.sleep(2)
        shell.send("enable\n")
        time.sleep(1)
        shell.send("zxr10\n")
        time.sleep(1)

        if ne_type in ["C620", "C650"]:
            shell.send(f"show pon power attenuation gpon_onu-1/{location}\n")
            time.sleep(2)
            shell.send(f"show gpon onu detail-info gpon_onu-1/{location}\n\n")
            shell.send(" ")
            time.sleep(2)
        else:
            shell.send(f"show pon power attenuation gpon-onu_1/{location}\n")
            time.sleep(2)
            shell.send(f"show gpon onu detail-info gpon-onu_1/{location}\n\n")
            shell.send(" ")
            time.sleep(2)

        output = ""
        time.sleep(2)
        while shell.recv_ready():
            output += shell.recv(4096).decode()
            time.sleep(0.1)

        pattern = r"\s+up\s+Rx\s+:(.*?)\s+Tx:(.*?)\s+(.*?)\s+down\s+Tx\s+:(.*?)\s+Rx:(.*?)\s+(.*?)\s"
        matches = re.search(pattern, output, re.DOTALL)
        if matches:
            olt_rx = matches.group(1).strip()
            onu_rx = matches.group(5).strip()
            power_levels = f"Ticket {ticket}: ONU Rx: {onu_rx} OLT Rx: {olt_rx}"
            print_message(power_levels, session_id)

            onu_prx = onu_rx.split("(")[0] if "(" in onu_rx else onu_rx
            olt_prx = olt_rx.split("(")[0] if "(" in olt_rx else olt_rx
            if onu_rx != "N/A" and olt_rx != "no signal":
                try:
                    if float(onu_prx) <= -28 or float(olt_prx) <= -28:
                        power_status = "Extreme Low Rx"
                    elif -28 < float(onu_prx) <= -24.5:
                        power_status = "Low Rx"
                    else:
                        power_status = "Power levels are optimal"
                except ValueError:
                    power_status = "Invalid power values"
            else:
                power_status = "Link currently down"

            if power_status == "Link currently down":
                comment = "LOS"
                ticket_status = "LOS"
            elif power_status in ["Low Rx", "Extreme Low Rx"]:
                comment = "Not Optimized"
                ticket_status = "Pending"
            else:
                comment = "Optimized, TT Closed"
                ticket_status = "Closed"
            return comment, power_status, ticket_status, power_levels
        return "No valid output", "Error", "Error", ""
    except Exception as e:
        print_message(f"SSH error for ticket {ticket}: {str(e)}", session_id)
        return "Error", "Error", "Error", ""
    finally:
        try:
            client.close()
        except:
            pass

def check_if_already_processed(df, session_id):
    """
    Check if the DataFrame already has SSH processing results.
    Returns True if already processed, False otherwise.
    """
    required_ssh_columns = ["Power Status", "Action", "Power Levels"]

    # Check if all required columns exist
    if not all(col in df.columns for col in required_ssh_columns):
        return False

    # Check if the columns have data (not all empty/NaN)
    has_data = (
        df["Power Status"].notna().any() and
        df["Action"].notna().any() and
        df["Power Levels"].notna().any()
    )

    if has_data:
        print_message("✅ File already contains SSH processing results. Skipping SSH step.", session_id)
        return True

    return False

def los_load_and_filter(file_path, session_id):
    try:
        df = pd.read_excel(file_path)
        col_customer = "Customer"
        col_category1 = "Category 1"
        col_category3 = "Category 3"
        col_subject = "Subject"
        if col_subject not in df.columns:
            print_message("Error: Subject column missing", session_id)
            return None

        original_rows = len(df)
        print_message(f"Loaded {original_rows} rows", session_id)

        if col_customer in df.columns:
            df = df[~df[col_customer].isin(["JAMII TELECOMMUNICATIONS LIMITED", "JAMII LIMITED"])]
        if col_category1 in df.columns:
            df = df[df[col_category1].str.lower() == "los fiber cut-cable maintenance incident"]
        exact_category3 = "GPON-Service Affecting-LOS Fiber Cut-Cable Maintenance Incident"
        if col_category3 in df.columns:
            df[col_category3] = df[col_category3].fillna('').astype(str).str.strip()
            df = df[df[col_category3].str.lower() == exact_category3.lower()]
        df[col_subject] = df[col_subject].fillna('').astype(str).str.strip()
        df = df[df[col_subject].str.contains(r'\bLOS\b', case=True, regex=True)]

        required_cols = ["ID", "Subscription", "Subject", "Creation", "Category 1", "Category 3", "Customer", "_Assign"]
        # Also keep SSH processing columns if they exist
        ssh_cols = ["Power Status", "Action", "Power Levels"]
        all_cols = required_cols + ssh_cols
        existing_cols = [c for c in all_cols if c in df.columns]
        df = df[existing_cols]
        print_message(f"Filtered to {len(df)} rows ready for processing", session_id)
        return df
    except Exception as e:
        print_message(f"Filter error: {str(e)}", session_id)
        return None

def los_process_ssh(df, username, password, session_id):
    olt_data = read_olt_data(session_id)
    if not olt_data:
        return None, None, []
    total_steps = len(df)
    completed = 0
    results = {}
    available_categories = los_category_options[:]

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = []
        for idx, row in df.iterrows():
            ticket = row["ID"]
            service_id = row["Subscription"]
            subject = row["Subject"]
            ne, location = extract_ne_and_location(subject)
            if ne and location:
                cleaned_ne = ne.replace("RB - ", "").strip().lower()
                if cleaned_ne in olt_data:
                    ne_type, host = olt_data[cleaned_ne]
                    ssh_email = "support" if ne in ["Karen OLT 1", "Kericho OLT"] else username
                    ssh_password = "Support@2024" if ne in ["Karen OLT 1", "Kericho OLT"] else password
                    future = executor.submit(ssh_task, idx, ticket, service_id, cleaned_ne, location, ne_type, host, ssh_email, ssh_password, subject, session_id, 'los')
                    futures.append((future, idx))
                else:
                    results[idx] = {"comment": "NE not found in OLT data", "status": "Error", "ticket_status": "Error", "power": ""}
            else:
                results[idx] = {"comment": "Could not parse NE/Location", "status": "Error", "ticket_status": "Error", "power": ""}

        for future, idx in futures:
            try:
                comment, status, ticket_status, power = future.result()
                results[idx] = {"comment": comment, "status": status, "ticket_status": ticket_status, "power": power}
            except Exception as e:
                results[idx] = {"comment": str(e), "status": "Error", "ticket_status": "Error", "power": ""}
            completed += 1
            update_progress(completed, total_steps, f"Processing SSH ({completed}/{total_steps})", session_id)

    for idx, res in results.items():
        if idx in df.index:
            df.at[idx, "Ticket Comment"] = res["comment"]
            df.at[idx, "Status"] = res["status"]
            df.at[idx, "Ticket Status"] = res["ticket_status"]
            df.at[idx, "Power Levels"] = res["power"]
            df.at[idx, "Power Status"] = res["status"]
            # Assign action based on power status
            if res["status"] == "Power levels are optimal":
                df.at[idx, "Action"] = "Process"
            elif res["status"] in ["Low Rx", "Extreme Low Rx"]:
                df.at[idx, "Action"] = "Transition"
            else:
                df.at[idx, "Action"] = "Skip"

    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"los_processed_{uuid.uuid4()}.xlsx")
    df.to_excel(output_path, index=False)
    print_message(f"SSH processing complete. Updated file: {output_path}", session_id)

    to_process = [r for r in results.values() if r["status"] == "Power levels are optimal"]
    return df, output_path, to_process

def lowrx_load_and_filter(file_path, session_id):
    # Similar to LOS, but adjust filters
    try:
        df = pd.read_excel(file_path)
        col_customer = "Customer"
        col_category1 = "Category 1"
        col_category3 = "Category 3"
        col_subject = "Subject"
        if col_subject not in df.columns:
            print_message("Error: Subject column missing", session_id)
            return None

        original_rows = len(df)
        print_message(f"Loaded {original_rows} rows", session_id)

        if col_customer in df.columns:
            df = df[~df[col_customer].isin(["JAMII TELECOMMUNICATIONS LIMITED", "JAMII LIMITED"])]
        if col_category1 in df.columns:
            df = df[df[col_category1].str.lower() == "low rx-cable maintenance incident"]  # Adjust for Low Rx
        exact_category3 = "GPON-Service Affecting-Low RX-Cable Maintenance Incident"  # Adjust
        if col_category3 in df.columns:
            df[col_category3] = df[col_category3].fillna('').astype(str).str.strip()
            df = df[df[col_category3].str.lower() == exact_category3.lower()]
        df[col_subject] = df[col_subject].fillna('').astype(str).str.strip()
        df = df[df[col_subject].str.contains(r'\bLow Rx\b', case=True, regex=True)]  # Adjust subject filter

        required_cols = ["ID", "Subscription", "Subject", "Creation", "Category 1", "Category 3", "Customer", "_Assign"]
        # Also keep SSH processing columns if they exist
        ssh_cols = ["Power Status", "Action", "Power Levels"]
        all_cols = required_cols + ssh_cols
        existing_cols = [c for c in all_cols if c in df.columns]
        df = df[existing_cols]
        print_message(f"Filtered to {len(df)} rows ready for processing", session_id)
        return df
    except Exception as e:
        print_message(f"Filter error: {str(e)}", session_id)
        return None

def lowrx_process_ssh(df, username, password, session_id):
    # Identical to LOS SSH, but use lowrx categories in CRM
    olt_data = read_olt_data(session_id)
    if not olt_data:
        return None, None, []
    total_steps = len(df)
    completed = 0
    results = {}
    available_categories = lowrx_category_options[:]

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = []
        for idx, row in df.iterrows():
            ticket = row["ID"]
            service_id = row["Subscription"]
            subject = row["Subject"]
            ne, location = extract_ne_and_location(subject)
            if ne and location:
                cleaned_ne = ne.replace("RB - ", "").strip().lower()
                if cleaned_ne in olt_data:
                    ne_type, host = olt_data[cleaned_ne]
                    ssh_email = "support" if ne in ["Karen OLT 1", "Kericho OLT"] else username
                    ssh_password = "Support@2024" if ne in ["Karen OLT 1", "Kericho OLT"] else password
                    future = executor.submit(ssh_task, idx, ticket, service_id, cleaned_ne, location, ne_type, host, ssh_email, ssh_password, subject, session_id, 'lowrx')
                    futures.append((future, idx))
                else:
                    results[idx] = {"comment": "NE not found in OLT data", "status": "Error", "ticket_status": "Error", "power": ""}
            else:
                results[idx] = {"comment": "Could not parse NE/Location", "status": "Error", "ticket_status": "Error", "power": ""}

        for future, idx in futures:
            try:
                comment, status, ticket_status, power = future.result()
                results[idx] = {"comment": comment, "status": status, "ticket_status": ticket_status, "power": power}
            except Exception as e:
                results[idx] = {"comment": str(e), "status": "Error", "ticket_status": "Error", "power": ""}
            completed += 1
            update_progress(completed, total_steps, f"Processing SSH ({completed}/{total_steps})", session_id)

    for idx, res in results.items():
        if idx in df.index:
            df.at[idx, "Ticket Comment"] = res["comment"]
            df.at[idx, "Status"] = res["status"]
            df.at[idx, "Ticket Status"] = res["ticket_status"]
            df.at[idx, "Power Levels"] = res["power"]
            df.at[idx, "Power Status"] = res["status"]
            # Assign action based on power status
            if res["status"] == "Power levels are optimal":
                df.at[idx, "Action"] = "Process"
            elif res["status"] in ["Low Rx", "Extreme Low Rx"]:
                df.at[idx, "Action"] = "Transition"
            else:
                df.at[idx, "Action"] = "Skip"

    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"lowrx_processed_{uuid.uuid4()}.xlsx")
    df.to_excel(output_path, index=False)
    print_message(f"SSH processing complete. Updated file: {output_path}", session_id)

    to_process = [r for r in results.values() if r["status"] == "Power levels are optimal"]
    return df, output_path, to_process

# Double Tickets (no SSH)
ALLOWED_CATEGORY1 = {
    "LOS Fiber Cut-Cable Maintenance Incident",
    "Extreme Low RX-Cable Maintenance Incident",
    "Low RX-Cable Maintenance Incident",
    "GPON LOSi-Cable Maintenance Incident",
    "GPON LOFi-Cable Maintenance Incident",
    "GPON Sfi-Cable Maintenance Incident",
}
CAT1_LOW = "Low RX-Cable Maintenance Incident"
CAT1_EXTREME = "Extreme Low RX-Cable Maintenance Incident"
LOW_EXTREME_SET = {CAT1_LOW, CAT1_EXTREME}

def safe_to_datetime(series, colname="Creation"):
    s = pd.to_datetime(series, errors="coerce")
    if s.isna().all():
        return None
    return s

def double_load_and_filter(file_path, apply_rules, session_id):
    """
    Load and filter Excel file to identify duplicate tickets.
    Uses the same logic as transition_lowrx_tickets:
    - Same subscription
    - Similar Category 1 (e.g., Low RX and Extreme Low RX are similar)
    - Keep oldest, close newer tickets
    """
    try:
        df = pd.read_excel(file_path)
        parsed_creation = safe_to_datetime(df.get("Creation"), "Creation")
        if parsed_creation is None:
            print_message("Could not parse Creation dates", session_id)
            return None, None
        df["Creation"] = parsed_creation
        print_message(f"Loaded {len(df)} rows", session_id)

        if apply_rules:
            print_message("Applying filters & rules", session_id)
            # Filters
            if "Customer" in df.columns:
                before = len(df)
                mask = ~df["Customer"].astype(str).str.contains("JAMIIL LIMITED", case=False, na=False)
                df = df[mask]
                print_message(f"Filtered Customer: removed {before - len(df)} rows", session_id)
            if "Category 3" in df.columns:
                before = len(df)
                mask = df["Category 3"].astype(str).str.strip().str.upper().str.startswith("GPON")
                df = df[mask]
                print_message(f"Filtered Category 3 GPON: kept {len(df)}", session_id)
            if "Category 1" in df.columns:
                before = len(df)
                df = df[df["Category 1"].astype(str).isin(ALLOWED_CATEGORY1)]
                print_message(f"Filtered Category 1: kept {len(df)}", session_id)

        # Use the same duplicate detection logic as transition_lowrx_tickets
        print_message("Checking for duplicate tickets (same subscription + similar category)...", session_id)

        processed_tickets = set()
        duplicate_groups = {}

        for idx, row in df.iterrows():
            ticket_id = row['ID']

            # Skip if already processed as part of a duplicate group
            if ticket_id in processed_tickets:
                continue

            subscription = row['Subscription']
            category1 = row.get('Category 1', '')

            # Check for duplicates using the same function as transition
            oldest_ticket, newer_tickets = check_duplicate_tickets_for_subscription(df, subscription, ticket_id, category1)

            if oldest_ticket is not None and len(newer_tickets) > 0:
                # Found duplicates!
                oldest_id = oldest_ticket['ID']

                # Mark all tickets in this group as processed
                processed_tickets.add(oldest_id)
                for _, newer_row in newer_tickets.iterrows():
                    processed_tickets.add(newer_row['ID'])

                # Store duplicate group
                key = f"{subscription} || {category1}"
                duplicate_groups[key] = {
                    'oldest': oldest_ticket,
                    'newer': newer_tickets,
                    'subscription': str(subscription),
                    'group_label': str(category1)
                }

        if len(duplicate_groups) > 0:
            total_newer = sum(len(v['newer']) for v in duplicate_groups.values())
            print_message(f"✅ Found {len(duplicate_groups)} duplicate groups with {total_newer} newer tickets to close", session_id)
        else:
            print_message("No duplicate tickets found", session_id)

        return df, duplicate_groups
    except Exception as e:
        print_message(f"Double filter error: {str(e)}", session_id)
        return None, None

def update_crm_ticket(driver, wait, ticket, status_value, category_value, resolution_text, session_id, is_closing=True, power_status=None, original_subject=None, power_levels=None):
    """
    Update CRM ticket with status, categories, resolution, and subject (when transitioning).

    Args:
        is_closing: If True (default), updates Status and Category 4 only (for closing tickets).
                   If False, updates Categories 1, 2, 3 and Subject (for transitioning tickets).
        power_status: Required when is_closing=False. Either "Low Rx" or "Extreme Low Rx".
        original_subject: Original ticket subject (required when is_closing=False for subject update).
        power_levels: Power level details (required when is_closing=False for subject update).
    """
    try:
        # Navigate directly to the ticket URL (much faster than searching)
        print_message(f"Opening ticket {ticket}...", session_id)
        ticket_url = f"https://intranet.jtl.co.ke/app/issue/{ticket}"
        driver.get(ticket_url)

        # Wait for page to load
        time.sleep(0.5)

        # Verify we're on the correct ticket page
        if ticket not in driver.current_url:
            raise Exception(f"Failed to navigate to ticket {ticket}")

        # Wait for the form to be visible
        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "form[data-doctype='Issue']")))
            time.sleep(0.1)  # Additional wait for form to fully render
        except:
            pass  # Continue even if form selector not found

        print_message(f"✅ Opened ticket {ticket}", session_id)

        print_message(f"Updating ticket {ticket}...", session_id)

        # Update status with retry using correct CSS selector
        # Only update status if closing, leave as-is if transitioning
        if is_closing:
            # ========== STEP 1: UPDATE STATUS TO "CLOSED" ==========
            max_status_attempts = 3
            for status_attempt in range(max_status_attempts):
                try:
                    # Wait for status element to be present
                    print_message(f"Looking for status field...", session_id)
                    status_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "select[data-fieldname='status']")))
                    print_message(f"✅ Found status field", session_id)

                    # Get current status value
                    current_status = status_element.get_attribute("value")
                    print_message(f"Current status: '{current_status}'", session_id)

                    # Scroll into view
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", status_element)
                    time.sleep(0.1)

                    # Use JavaScript to set the value
                    print_message(f"Using JavaScript to set status to '{status_value}'...", session_id)
                    driver.execute_script(f"""
                        var select = arguments[0];
                        select.value = '{status_value}';

                        // Trigger change event to notify the CRM
                        var event = new Event('change', {{ bubbles: true }});
                        select.dispatchEvent(event);

                        // Also trigger input event
                        var inputEvent = new Event('input', {{ bubbles: true }});
                        select.dispatchEvent(inputEvent);
                    """, status_element)
                    time.sleep(0.1)

                    # Verify the change
                    new_status = status_element.get_attribute("value")
                    print_message(f"Status after update: '{new_status}'", session_id)

                    if new_status == status_value:
                        print_message(f"✅ Status successfully changed to '{status_value}'", session_id)
                    else:
                        print_message(f"⚠️ Status did NOT change! Still: '{new_status}'", session_id)

                    break
                except Exception as e:
                    if status_attempt < max_status_attempts - 1:
                        print_message(f"⚠️ Retry status update (attempt {status_attempt + 1}/{max_status_attempts})", session_id)
                        time.sleep(0.5)
                    else:
                        raise Exception(f"Failed to update status: {str(e)}")
        else:
            print_message(f"⏭️ Transitioning mode - leaving status as-is for {ticket}", session_id)

        # Parse category value into parts
        # Input Format: "Part1-Part2-Part3-Part4"
        # Example Input: "B.C. Enclosure Maintenance-GPON-Service Affecting-LOS Fiber Cut"
        # For closing: Only update Category 4 = Part1 (first part)
        category_parts = category_value.split('-')

        max_category_attempts = 3

        if is_closing:
            # ========== STEP 2: UPDATE CATEGORY 4 ==========
            # CLOSING MODE: Only update Category 4 (faster - skip Categories 1, 2, 3)
            # Category 4 = Part 1 (first part of the input string)
            if len(category_parts) >= 1:
                category4_value = category_parts[0].strip()
                for category_attempt in range(max_category_attempts):
                    try:
                        print_message(f"Looking for Category 4 field...", session_id)
                        category4_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[data-fieldname='custom_category_4']")))
                        print_message(f"✅ Found Category 4 field", session_id)

                        # Scroll into view
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", category4_input)
                        time.sleep(0.3)

                        # Clear and type
                        category4_input.clear()
                        time.sleep(0.2)
                        category4_input.send_keys(category4_value)
                        time.sleep(1)  # Wait for autocomplete dropdown to appear

                        # Wait for dropdown to appear and click first item
                        try:
                            dropdown_item = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".awesomplete ul li")))
                            dropdown_item.click()
                            time.sleep(0.5)
                        except:
                            # Fallback to RETURN if dropdown doesn't appear
                            category4_input.send_keys(Keys.RETURN)
                            time.sleep(0.5)

                        # Verify
                        final_cat4 = category4_input.get_attribute("value")
                        if final_cat4 == category4_value:
                            print_message(f"✅ Set Category 4: {category4_value}", session_id)
                        else:
                            print_message(f"⚠️ Category 4 mismatch! Expected: {category4_value}, Got: {final_cat4}", session_id)
                        break
                    except Exception as e:
                        if category_attempt < max_category_attempts - 1:
                            print_message(f"⚠️ Retry Category 4 (attempt {category_attempt + 1}/{max_category_attempts})", session_id)
                            time.sleep(0.5)
                        else:
                            raise Exception(f"Failed to update Category 4: {str(e)}")
        else:
            # TRANSITIONING MODE: Update Description, Subject, then Categories 1, 2, 3 (DO NOT update Category 4)
            # Order: Description -> Subject -> Categories -> Resolution

            # ========== STEP 0: UPDATE DESCRIPTION ==========
            if power_levels:
                try:
                    print_message("Updating description with power levels...", session_id)
                    description_container = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-fieldname='description']")))
                    description_editor = description_container.find_element(By.CSS_SELECTOR, "div.ql-editor")

                    # Get current description text
                    current_description = description_editor.text

                    # Remove old power levels line if it exists
                    cleaned_description = re.sub(r'Power Levels:.*', '', current_description, flags=re.IGNORECASE).strip()

                    # Determine status label and replace 'LOS'
                    status_label = "Extreme Low Rx" if power_status and "Extreme" in power_status else "Low Rx"
                    modified_description = re.sub(r'\bLOS\b', status_label, cleaned_description, flags=re.IGNORECASE)

                    # Format the new power levels string from the ssh output
                    power_details_match = re.search(r'ONU Rx:.*', power_levels)
                    if power_details_match:
                        power_details = power_details_match.group(0)
                        new_power_string = f"Power Levels: {power_details}"
                        
                        # Combine the modified description with the new power levels
                        final_description = f"{modified_description}\n{new_power_string}"
                    else:
                        final_description = modified_description


                    # Use JavaScript to update the Quill editor, handling newlines
                    driver.execute_script(
                        """
                        var editor = arguments[0];
                        var newContent = arguments[1];
                        // Replace newlines with paragraph tags for Quill
                        editor.innerHTML = '<p>' + newContent.replace(/\\n/g, '</p><p>') + '</p>';
                        // Trigger input event to notify CRM of the change
                        var event = new Event('input', { bubbles: true });
                        editor.dispatchEvent(event);
                        """,
                        description_editor,
                        final_description
                    )
                    time.sleep(1)
                    print_message("✅ Description updated with new power levels.", session_id)
                except Exception as e:
                    print_message(f"⚠️ Could not update description: {str(e)}", session_id)
                    # Continue with other updates even if description fails

            # ========== STEP 1: UPDATE SUBJECT ==========
            if original_subject and power_levels:
                # Determine status label from power_status
                status_label = "Extreme Low Rx" if power_status and "Extreme" in power_status else "Low Rx"

                # Replace "LOS" with power status and add power level details
                new_subject = original_subject.replace("LOS", status_label)

                # Parse power levels to identify which power is low
                power_detail = ""
                if "ONU Rx:" in power_levels and "OLT Rx:" in power_levels:
                    # Extract ONU and OLT values
                    onu_match = re.search(r'ONU Rx:\s*([-\d.]+)', power_levels)
                    olt_match = re.search(r'OLT Rx:\s*([-\d.]+)', power_levels)

                    if onu_match and olt_match:
                        onu_val = float(onu_match.group(1))
                        olt_val = float(olt_match.group(1))

                        # Determine which power is low
                        if power_status == "Extreme Low Rx":
                            if onu_val <= -28 and olt_val <= -28:
                                power_detail = f" (ONU: {onu_val} dBm, OLT: {olt_val} dBm)"
                            elif onu_val <= -28:
                                power_detail = f" (ONU: {onu_val} dBm)"
                            elif olt_val <= -28:
                                power_detail = f" (OLT: {olt_val} dBm)"
                        else:  # Low Rx
                            if onu_val <= -24.5 and onu_val > -28:
                                power_detail = f" (ONU: {onu_val} dBm)"
                            elif olt_val <= -24.5 and olt_val > -28:
                                power_detail = f" (OLT: {olt_val} dBm)"
                            elif onu_val <= -24.5 and olt_val <= -24.5:
                                power_detail = f" (ONU: {onu_val} dBm, OLT: {olt_val} dBm)"

                # Append power detail to subject
                new_subject = new_subject + power_detail

                # Truncate to 140 characters
                if len(new_subject) > 140:
                    new_subject = new_subject[:140]

                # Update subject field using Selenium (click, clear, type)
                max_subject_attempts = 3
                for subject_attempt in range(max_subject_attempts):
                    try:
                        # Wait for subject field to be clickable
                        subject_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-fieldname='subject']")))

                        # Get original value
                        original_value = subject_input.get_attribute("value")
                        print_message(f"📊 Original subject: {original_value[:50]}...", session_id)

                        # Scroll into view
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", subject_input)
                        time.sleep(0.5)

                        # Click to focus
                        subject_input.click()
                        time.sleep(0.3)

                        # Select all and delete
                        subject_input.send_keys(Keys.CONTROL + "a")
                        time.sleep(0.2)
                        subject_input.send_keys(Keys.DELETE)
                        time.sleep(0.3)

                        # Type new subject
                        subject_input.send_keys(new_subject)
                        time.sleep(0.5)

                        # Press TAB to trigger blur/change events
                        subject_input.send_keys(Keys.TAB)
                        time.sleep(0.5)

                        # Verify the change
                        final_value = subject_input.get_attribute("value")
                        if final_value == new_subject:
                            print_message(f"✅ Updated Subject: {new_subject[:50]}... (verified)", session_id)
                        else:
                            print_message(f"⚠️ Subject mismatch! Expected: {new_subject[:30]}... Got: {final_value[:30]}...", session_id)
                        break
                    except Exception as e:
                        if subject_attempt < max_subject_attempts - 1:
                            print_message(f"⚠️ Retry Subject update (attempt {subject_attempt + 1}/{max_subject_attempts})", session_id)
                            time.sleep(0.5)
                        else:
                            raise Exception(f"Failed to update Subject: {str(e)}")

            # ========== STEP 2: UPDATE CATEGORIES ==========
            # Category 1 = "Extreme Low RX" or "Low RX" (based on power_status)
            # Category 2 = "QoS Affecting"
            # Category 3 = "GPON"

            # Determine Category 1 value based on power status
            if power_status and "Extreme" in power_status:
                cat1_value = "Extreme Low RX"
            else:
                cat1_value = "Low RX"

            # Update Category 1
            for category_attempt in range(max_category_attempts):
                try:
                    category1_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[data-fieldname='custom_category_1']")))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", category1_input)
                    time.sleep(0.3)
                    category1_input.clear()
                    time.sleep(0.2)
                    category1_input.send_keys(cat1_value)
                    time.sleep(1)  # Wait for autocomplete dropdown to appear

                    # Wait for dropdown to appear and click first item
                    try:
                        dropdown_item = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".awesomplete ul li")))
                        dropdown_item.click()
                        time.sleep(0.5)
                    except:
                        # Fallback to RETURN if dropdown doesn't appear
                        category1_input.send_keys(Keys.RETURN)
                        time.sleep(0.5)

                    print_message(f"✅ Set Category 1: {cat1_value}", session_id)
                    break
                except Exception as e:
                    if category_attempt < max_category_attempts - 1:
                        print_message(f"⚠️ Retry Category 1 (attempt {category_attempt + 1}/{max_category_attempts})", session_id)
                        time.sleep(0.5)
                    else:
                        raise Exception(f"Failed to update Category 1: {str(e)}")

            # Update Category 2 = "QoS Affecting"
            for category_attempt in range(max_category_attempts):
                try:
                    category2_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[data-fieldname='custom_category_2']")))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", category2_input)
                    time.sleep(0.3)
                    category2_input.clear()
                    time.sleep(0.2)
                    category2_input.send_keys("QoS Affecting")
                    time.sleep(1)  # Wait for autocomplete dropdown to appear

                    # Wait for dropdown to appear and click first item
                    try:
                        dropdown_item = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".awesomplete ul li")))
                        dropdown_item.click()
                        time.sleep(0.5)
                    except:
                        # Fallback to RETURN if dropdown doesn't appear
                        category2_input.send_keys(Keys.RETURN)
                        time.sleep(0.5)

                    print_message(f"✅ Set Category 2: QoS Affecting", session_id)
                    break
                except Exception as e:
                    if category_attempt < max_category_attempts - 1:
                        print_message(f"⚠️ Retry Category 2 (attempt {category_attempt + 1}/{max_category_attempts})", session_id)
                        time.sleep(0.5)
                    else:
                        raise Exception(f"Failed to update Category 2: {str(e)}")

            # Update Category 3 = "GPON"
            for category_attempt in range(max_category_attempts):
                try:
                    category3_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[data-fieldname='custom_category_3']")))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", category3_input)
                    time.sleep(0.3)
                    category3_input.clear()
                    time.sleep(0.2)
                    category3_input.send_keys("GPON")
                    time.sleep(1)  # Wait for autocomplete dropdown to appear

                    # Wait for dropdown to appear and click first item
                    try:
                        dropdown_item = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".awesomplete ul li")))
                        dropdown_item.click()
                        time.sleep(0.5)
                    except:
                        # Fallback to RETURN if dropdown doesn't appear
                        category3_input.send_keys(Keys.RETURN)
                        time.sleep(0.5)

                    print_message(f"✅ Set Category 3: GPON", session_id)
                    break
                except Exception as e:
                    if category_attempt < max_category_attempts - 1:
                        print_message(f"⚠️ Retry Category 3 (attempt {category_attempt + 1}/{max_category_attempts})", session_id)
                        time.sleep(0.5)
                    else:
                        raise Exception(f"Failed to update Category 3: {str(e)}")

            # DO NOT update Category 4 when transitioning
            print_message(f"⏭️ Transitioning mode - Category 4 left as-is", session_id)

        # ========== STEP 3: UPDATE RESOLUTION ==========
        max_resolution_attempts = 3
        for resolution_attempt in range(max_resolution_attempts):
            try:
                print_message(f"Looking for resolution field...", session_id)
                resolution_container = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-fieldname='resolution_details']")))
                print_message(f"✅ Found resolution container", session_id)

                # Scroll into view
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", resolution_container)
                time.sleep(0.3)

                # Find the ql-editor within the container
                resolution_box = resolution_container.find_element(By.CSS_SELECTOR, "div.ql-editor")

                # Clear and set new value using JavaScript
                driver.execute_script("""
                    var editor = arguments[0];
                    var newText = arguments[1];

                    // Clear existing content
                    editor.innerHTML = '';

                    // Set new content
                    editor.innerHTML = '<p>' + newText + '</p>';

                    // Trigger input event
                    var event = new Event('input', { bubbles: true });
                    editor.dispatchEvent(event);
                """, resolution_box, resolution_text)
                time.sleep(0.5)

                print_message(f"✅ Set resolution", session_id)
                break
            except Exception as e:
                if resolution_attempt < max_resolution_attempts - 1:
                    print_message(f"⚠️ Retry resolution update (attempt {resolution_attempt + 1}/{max_resolution_attempts})", session_id)
                    time.sleep(0.5)
                else:
                    raise Exception(f"Failed to update resolution: {str(e)}")

        # Save with retry
        max_save_attempts = 3
        for save_attempt in range(max_save_attempts):
            try:
                save_button = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@data-label='Save']")))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", save_button)
                time.sleep(0.3)
                driver.execute_script("arguments[0].click();", save_button)
                time.sleep(3)  # Wait longer for save to complete

                # Check for save success indicators
                try:
                    # Look for success message or check if form is no longer dirty
                    success_msg = driver.find_elements(By.CSS_SELECTOR, ".indicator.green, .msgprint")
                    if success_msg:
                        print_message(f"✅ Saved ticket {ticket} (success indicator found)", session_id)
                    else:
                        print_message(f"✅ Saved ticket {ticket}", session_id)
                except:
                    print_message(f"✅ Saved ticket {ticket}", session_id)
                break
            except Exception as e:
                if save_attempt < max_save_attempts - 1:
                    print_message(f"Retry save (attempt {save_attempt + 1}/{max_save_attempts})", session_id)
                    time.sleep(0.5)
                else:
                    raise Exception(f"Failed to save ticket: {str(e)}")

       
    except Exception as e:
        print_message(f"CRM update error for {ticket}: {str(e)}", session_id)
        raise

def check_duplicate_tickets_for_subscription(df, subscription, current_ticket_id, current_category1):
    """
    Check if there are other tickets in the file with:
    1. Same subscription
    2. Similar Category 1 (same base category)
    3. Return the oldest ticket and all newer tickets
    """
    # Extract base category from current ticket's Category 1
    # E.g., "Low RX-Cable Maintenance Incident" or "Extreme Low RX-Cable Maintenance Incident"
    current_base_category = None
    if pd.notna(current_category1):
        current_cat_lower = str(current_category1).lower()
        if "low rx" in current_cat_lower or "low rx" in current_cat_lower:
            current_base_category = "low_rx"  # Matches both Low RX and Extreme Low RX
        elif "los" in current_cat_lower:
            current_base_category = "los"
        elif "lofi" in current_cat_lower:
            current_base_category = "lofi"
        elif "losi" in current_cat_lower:
            current_base_category = "losi"
        elif "sfi" in current_cat_lower:
            current_base_category = "sfi"

    if current_base_category is None:
        # No recognizable category, no duplicates
        return None, []

    # Find all tickets with same subscription
    same_subscription = df[df['Subscription'] == subscription].copy()

    # Filter by similar Category 1
    def matches_category(cat1):
        if pd.isna(cat1):
            return False
        cat_lower = str(cat1).lower()
        if current_base_category == "low_rx":
            return "low rx" in cat_lower or "low rx" in cat_lower
        else:
            return current_base_category in cat_lower

    same_subscription['matches_category'] = same_subscription['Category 1'].apply(matches_category)
    duplicates = same_subscription[same_subscription['matches_category']]

    if len(duplicates) <= 1:
        # Only current ticket or no tickets found
        return None, []

    # Sort by Creation date to find oldest
    if 'Creation' in duplicates.columns:
        duplicates = duplicates.sort_values('Creation')

    # Get oldest ticket
    oldest_ticket = duplicates.iloc[0]

    # Get all newer tickets (excluding oldest)
    newer_tickets = duplicates.iloc[1:]

    return oldest_ticket, newer_tickets

def transition_lowrx_tickets(df, session_id):
    """Transition Low Rx and Extreme Low Rx tickets, checking for duplicates"""
    driver = sessions[session_id]['driver']
    wait = sessions[session_id]['wait']
    processed = 0
    skipped_as_newer = set()  # Track tickets we skip because they're newer duplicates

    # Filter tickets that need transition (Low Rx or Extreme Low Rx)
    tickets_to_transition = df[df['Power Status'].isin(['Low Rx', 'Extreme Low Rx'])]
    total = len(tickets_to_transition)

    if total == 0:
        print_message("No Low Rx or Extreme Low Rx tickets to transition.", session_id)
        return

    print_message(f"Found {total} tickets with Low Rx or Extreme Low Rx to transition...", session_id)
    print_message("Checking for duplicate tickets (same subscription + similar category)...", session_id)

    for idx, row in tickets_to_transition.iterrows():
        # Check for pause
        if sessions[session_id].get('paused', False):
            print_message("⏸️ Processing PAUSED. Click 'Resume' to continue.", session_id)
            while sessions[session_id].get('paused', False):
                time.sleep(1)  # Wait until resumed
            print_message("▶️ Processing RESUMED.", session_id)

        # Check for stop
        if sessions[session_id].get('stopped', False):
            print_message(f"⏹️ Processing STOPPED by user. Processed {processed}/{total} tickets.", session_id)
            return

        ticket = row["ID"]

        # Skip if this ticket was already identified as a newer duplicate
        if ticket in skipped_as_newer:
            print_message(f"Ticket {ticket}: Skipping (already closed as duplicate)", session_id)
            processed += 1
            update_progress(processed, total, f"Transitioning Low/Extreme Rx ({processed}/{total})", session_id)
            continue

        subscription = row["Subscription"]
        power_status = row["Power Status"]
        power_levels = row["Power Levels"]
        original_subject = row.get("Subject", "")
        current_category1 = row.get("Category 1", "")

        try:
            # Determine category based on power status
            if power_status == "Extreme Low Rx":
                category = extreme_lowrx_transition_category
                status_label = "Extreme Low Rx"
            else:  # Low Rx
                category = lowrx_transition_category
                status_label = "Low Rx"

            # Check for duplicate tickets with same subscription and similar Category 1
            oldest_ticket, newer_tickets = check_duplicate_tickets_for_subscription(df, subscription, ticket, current_category1)

            if oldest_ticket is not None and len(newer_tickets) > 0:
                # Found duplicates - determine if current ticket is oldest or newer
                oldest_ticket_id = oldest_ticket['ID']

                if ticket == oldest_ticket_id:
                    # Current ticket is the OLDEST - keep it, close all newer ones
                    print_message(f"Ticket {ticket}: OLDEST ticket for subscription {subscription}. Transitioning this one.", session_id)

                    # Transition the oldest ticket
                    resolution = f"{status_label} detected. Power Levels: {power_levels}. Ticket transitioned for further investigation."
                    update_crm_ticket(driver, wait, ticket, "In Process", category, resolution, session_id, is_closing=False, power_status=power_status, original_subject=original_subject, power_levels=power_levels)

                    # Close all newer tickets with reference to this oldest one
                    for _, newer_row in newer_tickets.iterrows():
                        newer_ticket_id = newer_row['ID']
                        newer_power_levels = newer_row.get('Power Levels', '')
                        newer_subject = newer_row.get('Subject', '')

                        resolution = f"{status_label} detected. Power Levels: {newer_power_levels}. Duplicate ticket found for same subscription. Closing and referencing oldest ticket: {oldest_ticket_id}."
                        print_message(f"Ticket {newer_ticket_id}: NEWER duplicate. Closing with reference to {oldest_ticket_id}.", session_id)
                        update_crm_ticket(driver, wait, newer_ticket_id, "Closed", category, resolution, session_id, is_closing=True, power_status=power_status, original_subject=newer_subject, power_levels=newer_power_levels)

                        # Mark as processed so we skip it when we encounter it in the loop
                        skipped_as_newer.add(newer_ticket_id)
                else:
                    # Current ticket is a NEWER duplicate - close it with reference to oldest
                    print_message(f"Ticket {ticket}: NEWER duplicate. Closing with reference to oldest ticket {oldest_ticket_id}.", session_id)
                    resolution = f"{status_label} detected. Power Levels: {power_levels}. Duplicate ticket found for same subscription. Closing and referencing oldest ticket: {oldest_ticket_id}."
                    update_crm_ticket(driver, wait, ticket, "Closed", category, resolution, session_id, is_closing=True, power_status=power_status, original_subject=original_subject, power_levels=power_levels)
                    skipped_as_newer.add(ticket)
            else:
                # No duplicates - transition to appropriate category
                resolution = f"{status_label} detected. Power Levels: {power_levels}. Ticket transitioned for further investigation."
                print_message(f"Ticket {ticket}: No duplicates found. Transitioning to {status_label} category.", session_id)
                update_crm_ticket(driver, wait, ticket, "In Process", category, resolution, session_id, is_closing=False, power_status=power_status, original_subject=original_subject, power_levels=power_levels)

            processed += 1
            update_progress(processed, total, f"Transitioning Low/Extreme Rx ({processed}/{total})", session_id)

        except Exception as e:
            print_message(f"Failed to transition ticket {ticket}: {str(e)}", session_id)

    print_message(f"Low Rx/Extreme Low Rx transition complete. Processed {processed}/{total} tickets.", session_id)

def los_lowrx_update_crm(df, session_id, mode):
    driver = sessions[session_id]['driver']
    wait = sessions[session_id]['wait']
    available = sessions[session_id]['available_categories']
    options = los_category_options if mode == 'los' else lowrx_category_options
    processed = 0
    total = len(df[df['Action'] == 'Process'])
    for idx, row in df[df['Action'] == 'Process'].iterrows():
        # Check for pause
        if sessions[session_id].get('paused', False):
            print_message("⏸️ Processing PAUSED. Click 'Resume' to continue.", session_id)
            while sessions[session_id].get('paused', False):
                time.sleep(1)  # Wait until resumed
            print_message("▶️ Processing RESUMED.", session_id)

        # Check for stop
        if sessions[session_id].get('stopped', False):
            print_message(f"⏹️ Processing STOPPED by user. Processed {processed}/{total} tickets.", session_id)
            return

        ticket = row["ID"]
        power_levels = row["Power Levels"]
        try:
            category = get_non_repeating_category(available, options)
            resolution = power_levels.replace(f"Ticket {ticket}: ", "") if power_levels else "Optimized"
            update_crm_ticket(driver, wait, ticket, "Closed", category, resolution, session_id, is_closing=True)
            processed += 1
            update_progress(processed, total, f"CRM Update ({processed}/{total})", session_id)
        except Exception as e:
            print_message(f"Failed to update {ticket}: {str(e)}", session_id)
    print_message("CRM updates complete.", session_id)

def double_update_crm(df, service_tickets, session_id):
    driver = sessions[session_id]['driver']
    wait = sessions[session_id]['wait']
    processed = 0
    total = sum(len(v['newer']) for v in service_tickets.values())
    category_value = "Double Ticket"
    for key, bundle in service_tickets.items():
        old_ticket = bundle['oldest']['ID']
        for _, newer_row in bundle['newer'].iterrows():
            # Check for pause
            if sessions[session_id].get('paused', False):
                print_message("⏸️ Processing PAUSED. Click 'Resume' to continue.", session_id)
                while sessions[session_id].get('paused', False):
                    time.sleep(1)  # Wait until resumed
                print_message("▶️ Processing RESUMED.", session_id)

            # Check for stop
            if sessions[session_id].get('stopped', False):
                print_message(f"⏹️ Processing STOPPED by user. Processed {processed}/{total} tickets.", session_id)
                return

            ticket = newer_row['ID']
            try:
                resolution = f"Existing TT: {old_ticket}"
                update_crm_ticket(driver, wait, ticket, "Closed", category_value, resolution, session_id, is_closing=True)
                if 'Comments' not in df.columns:
                    df['Comments'] = ''
                df.loc[df['ID'] == ticket, 'Comments'] = f"Duplicate referenced to {old_ticket}"
                processed += 1
                update_progress(processed, total, f"CRM Update ({processed}/{total})", session_id)
            except Exception as e:
                print_message(f"Failed to update {ticket}: {str(e)}", session_id)
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"double_updated_{uuid.uuid4()}.xlsx")
    df.to_excel(output_path, index=False)
    print_message(f"CRM complete. Updated file: {output_path}", session_id)
    return output_path

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file'})
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'})
    session_id = str(uuid.uuid4())
    sessions[session_id] = {}
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    mode = request.form['mode']
    username = request.form['username']
    password = request.form['password']
    apply_rules = request.form.get('apply_rules', 'on') == 'on'
    dry_run = request.form.get('dry_run', 'off') == 'on'
    thread = threading.Thread(target=process_mode, args=(mode, file_path, username, password, session_id, apply_rules, dry_run))
    thread.daemon = True
    thread.start()
    return jsonify({'session_id': session_id, 'status': 'started'})

def process_mode(mode, file_path, username, password, session_id, apply_rules, dry_run):
    print_message(f"Starting {mode.upper()} mode processing...", session_id)
    update_progress(0, 100, "Initializing...", session_id)
    if mode in ['los', 'lowrx']:
        if mode == 'los':
            df_filtered = los_load_and_filter(file_path, session_id)
            if df_filtered is None:
                return

            # Check if file already has SSH processing results
            if check_if_already_processed(df_filtered, session_id):
                df = df_filtered
                ssh_path = file_path
                # Extract tickets that need processing based on existing Action column
                to_process = df[df['Action'] == 'Process'].to_dict('records')
            else:
                df, ssh_path, to_process = los_process_ssh(df_filtered, username, password, session_id)
        else:
            df_filtered = lowrx_load_and_filter(file_path, session_id)
            if df_filtered is None:
                return

            # Check if file already has SSH processing results
            if check_if_already_processed(df_filtered, session_id):
                df = df_filtered
                ssh_path = file_path
                # Extract tickets that need processing based on existing Action column
                to_process = df[df['Action'] == 'Process'].to_dict('records')
            else:
                df, ssh_path, to_process = lowrx_process_ssh(df_filtered, username, password, session_id)

        # Count tickets by status
        num_to_process = len(to_process)
        num_lowrx = len(df[df['Power Status'] == 'Low Rx'])
        num_extreme_lowrx = len(df[df['Power Status'] == 'Extreme Low Rx'])

        print_message(f"Processing summary: {num_to_process} optimal, {num_lowrx} Low Rx, {num_extreme_lowrx} Extreme Low Rx", session_id)

        # Login to CRM if there are any tickets to process
        if num_to_process > 0 or num_lowrx > 0 or num_extreme_lowrx > 0:
            if not login_and_process(username, password, session_id):
                print_message("CRM login failed. Aborting.", session_id)
                return

            # Process optimal tickets (close them)
            if num_to_process > 0:
                print_message(f"{num_to_process} optimal tickets ready for CRM closure...", session_id)
                los_lowrx_update_crm(df, session_id, mode)

            # Transition Low Rx and Extreme Low Rx tickets
            if num_lowrx > 0 or num_extreme_lowrx > 0:
                print_message(f"Transitioning {num_lowrx + num_extreme_lowrx} Low/Extreme Rx tickets...", session_id)
                transition_lowrx_tickets(df, session_id)
        else:
            print_message("No tickets to process. Processing complete.", session_id)
            update_progress(100, 100, "Complete (no CRM updates needed)", session_id)
            return
    else:  # double
        df, service_tickets = double_load_and_filter(file_path, apply_rules, session_id)
        if df is None or not service_tickets:
            print_message("No duplicates found or filter error.", session_id)
            return
        num_to_process = sum(len(v['newer']) for v in service_tickets.values())
        if dry_run:
            print_message(f"Dry run: Would process {num_to_process} tickets. No CRM changes.", session_id)
            update_progress(100, 100, "Dry run complete", session_id)
            return
        print_message(f"Processing {num_to_process} duplicate tickets in CRM...", session_id)
        if not login_and_process(username, password, session_id):
            print_message("CRM login failed. Aborting.", session_id)
            return
        double_update_crm(df, service_tickets, session_id)
    update_progress(100, 100, "Complete", session_id)
    # Cleanup
    if 'driver' in sessions[session_id]:
        sessions[session_id]['driver'].quit()
        del sessions[session_id]['driver']
    del sessions[session_id]

@app.route('/report', methods=['POST'])
def report():
    if 'file' not in request.files:
        return jsonify({'error': 'No file'})
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'})
    session_id = str(uuid.uuid4())
    sessions[session_id] = {}
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    report_type = request.form['report_type']
    thread = threading.Thread(target=process_report_mode, args=(report_type, file_path, session_id))
    thread.daemon = True
    thread.start()
    return jsonify({'session_id': session_id, 'status': 'started'})

def process_report_mode(report_type, file_path, session_id):
    print_message(f"Starting {report_type.upper()} report generation...", session_id)
    update_progress(0, 100, "Initializing...", session_id)
    try:
        in_path = Path(file_path)
        out_filename = f"{in_path.stem}_{report_type}_report_{uuid.uuid4().hex[:8]}.xlsx"
        out_path = Path(app.config['UPLOAD_FOLDER']) / out_filename

        update_progress(10, 100, "Processing workbook...", session_id)
        if report_type == 'gpon':
            stats = generate_report.process_workbook(in_path, out_path)
        elif report_type == 'enterprise':
            stats = enterprise_report.process_workbook(in_path, out_path)
        else:
            print_message(f"Error: Unknown report type '{report_type}'", session_id)
            update_progress(100, 100, "Error", session_id)
            return

        update_progress(80, 100, "Finalizing report...", session_id)
        for sheet, st in stats.items():
            log_msg = (
                f"[Sheet: {sheet}] rows: {st.get('original_rows')} -> {st.get('final_rows')} | "
                f"closed rows removed: {st.get('closed_rows_removed')} | "
            )
            if 'gpon_rows' in st:
                log_msg += f"GPON rows: {st.get('gpon_rows', 0)}"

            print_message(log_msg, session_id)
        
        print_message(f"Wrote cleaned report to: {out_path.as_posix()}", session_id)
        update_progress(100, 100, "Complete", session_id)

    except Exception as e:
        logging.error(f"Report generation failed for session {session_id}: {e}", exc_info=True)
        print_message(f"Error during report generation: {e}", session_id)
        update_progress(100, 100, "Error", session_id)


@socketio.on('connect')
def handle_connect():
    print('Client connected')

@socketio.on('join_room')
def on_join(data):
    join_room(data['room'])

@app.route('/download/<path:filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

@app.route('/pause', methods=['POST'])
def pause_processing():
    """Pause processing after current ticket completes"""
    data = request.json
    session_id = data.get('session_id')

    if session_id and session_id in sessions:
        sessions[session_id]['paused'] = True
        return jsonify({'status': 'success', 'message': 'Processing will pause after current ticket'})
    return jsonify({'status': 'error', 'message': 'Invalid session'})

@app.route('/resume', methods=['POST'])
def resume_processing():
    """Resume paused processing"""
    data = request.json
    session_id = data.get('session_id')

    if session_id and session_id in sessions:
        sessions[session_id]['paused'] = False
        return jsonify({'status': 'success', 'message': 'Processing resumed'})
    return jsonify({'status': 'error', 'message': 'Invalid session'})

@app.route('/stop', methods=['POST'])
def stop_processing():
    """Stop processing after current ticket completes"""
    data = request.json
    session_id = data.get('session_id')

    if session_id and session_id in sessions:
        sessions[session_id]['stopped'] = True
        return jsonify({'status': 'success', 'message': 'Processing will stop after current ticket'})
    return jsonify({'status': 'error', 'message': 'Invalid session'})

if __name__ == '__main__':
    socketio.run(app, host='0.0.0.0', port=5000, debug=False)
    