from flask import Flask, render_template, request, jsonify, send_file
from flask_socketio import SocketIO, emit, join_room
import os
import sys
import threading
import time
import logging
from werkzeug.utils import secure_filename
import pandas as pd
from netmiko import ConnectHandler, NetmikoTimeoutException
import re
import concurrent.futures
import random
import platform
import subprocess
import shutil as shutil_module
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
import shutil
from datetime import datetime

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

# {session_id: {'driver': None, 'wait': None, 'available_categories': [], 'paused': False, 'stopped': False}}
sessions = {}

logging.basicConfig(
    level=logging.INFO,
    filename='app.log',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ---------------------------------------------------------------------------
# GPON OLT DATA CACHE
# ---------------------------------------------------------------------------
OLT_DATA_CACHE = None
OLT_DATA_MTIME = None
OLT_DATA_LOCK = threading.Lock()





# ---------------------------------------------------------------------------
# Category options
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# UTILITIES
# ---------------------------------------------------------------------------

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
    socketio.emit(
        'progress_update',
        {'percentage': percentage, 'message': message},
        room=session_id
    )

def read_olt_data(session_id):
    """Read OLT_DATA1.xlsx once and cache it."""
    global OLT_DATA_CACHE, OLT_DATA_MTIME

    file_path = "OLT_DATA1.xlsx"
    if not os.path.exists(file_path):
        print_message("Error: OLT_DATA1.xlsx not found in root directory!", session_id)
        return {}

    try:
        mtime = os.path.getmtime(file_path)
    except OSError as e:
        print_message(f"Error reading OLT data mtime: {e}", session_id)
        return {}

    with OLT_DATA_LOCK:
        if OLT_DATA_CACHE is not None and OLT_DATA_MTIME == mtime:
            print_message("OLT data loaded from cache.", session_id)
            return OLT_DATA_CACHE

        try:
            df = pd.read_excel(file_path)
            required_cols = {"NE", "Type", "Host"}
            missing = required_cols - set(df.columns)
            if missing:
                print_message(f"Error: OLT data missing columns: {', '.join(missing)}", session_id)
                return {}

            ne_data = {}
            for _, row in df.iterrows():
                ne_raw = row.get("NE", "")
                if pd.isna(ne_raw):
                    continue
                ne = str(ne_raw).strip().lower()
                ne_type = row.get("Type", "")
                host = row.get("Host", "")
                if not host:
                    continue
                ne_data[ne] = (ne_type, host)

            OLT_DATA_CACHE = ne_data
            OLT_DATA_MTIME = mtime
            print_message(f"OLT data loaded successfully ({len(ne_data)} NE entries).", session_id)
            return ne_data

        except Exception as e:
            print_message(f"Error reading OLT data: {str(e)}", session_id)
            return {}

# ---------------------------------------------------------------------------
# CHROMEDRIVER INITIALIZATION - WINDOWS SAFE
# ---------------------------------------------------------------------------
def initialize_chrome_driver(session_id):
    """
    Safely initialize ChromeDriver with Windows compatibility.
    Handles cache clearing and architecture detection.
    """
    print_message("Initializing Chrome WebDriver...", session_id)
    
    try:
        # System info logging
        system_info = platform.platform()
        arch = platform.architecture()
        print_message(f"System: {system_info}, Architecture: {arch[0]}", session_id)
        
        # Chrome options
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--ignore-certificate-errors")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-web-resources")
        chrome_options.add_argument("--headless") # Optional: Keep commented for visibility
        
        # Try to initialize ChromeDriver
        try:
            print_message("Attempting to get ChromeDriver path...", session_id)
            chromedriver_path = ChromeDriverManager().install()
            
            # Fix: Handle case where path points to wrong file (e.g., THIRD_PARTY_NOTICES)
            # The actual executable should be chromedriver.exe (Windows) or chromedriver (Linux/Mac)
            if "THIRD_PARTY_NOTICES" in chromedriver_path or not chromedriver_path.endswith(("chromedriver.exe", "chromedriver")):
                print_message(f"ChromeDriver path invalid: {chromedriver_path}. Searching for correct executable...", session_id)
                # Get the directory and find the actual executable
                base_dir = os.path.dirname(chromedriver_path)
                
                # Try to find chromedriver.exe or chromedriver in the directory
                if platform.system() == "Windows":
                    potential_paths = [
                        os.path.join(base_dir, "chromedriver.exe"),
                        os.path.join(os.path.dirname(base_dir), "chromedriver.exe"),
                    ]
                else:
                    potential_paths = [
                        os.path.join(base_dir, "chromedriver"),
                        os.path.join(os.path.dirname(base_dir), "chromedriver"),
                    ]
                
                chromedriver_path = None
                for path in potential_paths:
                    if os.path.exists(path) and os.path.isfile(path):
                        chromedriver_path = path
                        print_message(f"Found valid ChromeDriver at: {chromedriver_path}", session_id)
                        break
                
                if not chromedriver_path:
                    raise FileNotFoundError(f"Could not find valid ChromeDriver executable in {base_dir}")
            else:
                print_message(f"ChromeDriver path: {chromedriver_path}", session_id)
            
            # Make sure the file exists and is executable
            if not os.path.exists(chromedriver_path):
                raise FileNotFoundError(f"ChromeDriver not found at {chromedriver_path}")
            
            # On Windows, make it executable
            if platform.system() == "Windows":
                os.chmod(chromedriver_path, 0o755)
            
            driver = webdriver.Chrome(
                service=Service(chromedriver_path),
                options=chrome_options
            )
            
        except FileNotFoundError as e:
            print_message(f"ChromeDriver file not found: {str(e)}", session_id)
            print_message("Clearing ChromeDriver cache and retrying...", session_id)
            
            # Clear webdriver-manager cache
            try:
                cache_path = os.path.expanduser("~/.wdm")
                if os.path.exists(cache_path):
                    shutil_module.rmtree(cache_path)
                    print_message("Cleared webdriver-manager cache", session_id)
            except Exception as cache_error:
                print_message(f"Could not clear cache: {cache_error}", session_id)
            
            # Retry with fresh download
            print_message("Downloading fresh ChromeDriver...", session_id)
            chromedriver_path = ChromeDriverManager().install()
            
            # Apply same fix to retry
            if "THIRD_PARTY_NOTICES" in chromedriver_path or not chromedriver_path.endswith(("chromedriver.exe", "chromedriver")):
                base_dir = os.path.dirname(chromedriver_path)
                if platform.system() == "Windows":
                    chromedriver_path = os.path.join(base_dir, "chromedriver.exe")
                else:
                    chromedriver_path = os.path.join(base_dir, "chromedriver")
            
            if platform.system() == "Windows":
                os.chmod(chromedriver_path, 0o755)
            
            driver = webdriver.Chrome(
                service=Service(chromedriver_path),
                options=chrome_options
            )
        
        # Set window size
        driver.set_window_size(1440, 900)
        
        # Initialize WebDriverWait
        wait = WebDriverWait(driver, 10)
        
        print_message("Chrome WebDriver initialized successfully!", session_id)
        return driver, wait
        
    except Exception as e:
        error_msg = f"Failed to initialize ChromeDriver: {str(e)}"
        print_message(error_msg, session_id)
        logging.error(f"[{session_id}] {error_msg}", exc_info=True)
        raise RuntimeError(error_msg)

# ---------------------------------------------------------------------------
# LOGIN LOGIC (UPDATED TO MATCH OLD VERSION)
# ---------------------------------------------------------------------------
def login_and_process(username, password, session_id):
    print_message("Starting browser initialization for CRM login...", session_id)
    try:
        # Use the new ChromeDriver initialization function
        driver, wait = initialize_chrome_driver(session_id)
        
        # Store driver and wait in session
        sessions[session_id]['driver'] = driver
        sessions[session_id]['wait'] = wait
        sessions[session_id]['available_categories'] = los_category_options[:]
        sessions[session_id]['paused'] = False
        sessions[session_id]['stopped'] = False

        url = "https://intranet.jtl.co.ke/login"
        print_message(f"Navigating to login URL: {url}", session_id)
        
        driver.get(url)

        username_input = wait.until(EC.presence_of_element_located((By.ID, "login_email")))
        password_input = wait.until(EC.presence_of_element_located((By.ID, "login_password")))
        username_input.send_keys(username)
        password_input.send_keys(password)

        login_button = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "button.btn.btn-sm.btn-default.btn-block.btn-login.btn-ldap-login")
        ))
        login_button.click()

        wait.until(lambda d: "login" not in d.current_url.lower())
        print_message("Login successful.", session_id)

        print_message("Ready to process tickets.", session_id)
        return True

    except Exception as e:
        print_message(f"CRM login failed: {str(e)}", session_id)
        if session_id in sessions and 'driver' in sessions[session_id]:
            try:
                sessions[session_id]['driver'].quit()
            except Exception:
                pass
            sessions[session_id].pop('driver', None)
        return False

# ---------------------------------------------------------------------------
# SSH & FILE PROCESSING
# ---------------------------------------------------------------------------
def extract_ne_and_location(text):
    pattern = r"(.+?)\s+(\d+/\d{1,2}:\d{1,2})"
    match = re.search(pattern, text)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return None, None


# default timeouts (seconds) used for SSH connections.  Can be overridden by
# setting the SSH_CONN_TIMEOUT environment variable, e.g.:
#
#     export SSH_CONN_TIMEOUT=60
#
#  The Netmiko library itself will suggest increasing 'conn_timeout' when a
#  ``NetmikoTimeoutException`` is raised with a message like "No existing
#  session".  _run_ssh_commands now catches that specific case and retries once
#  with a larger timeout.
SSH_CONN_TIMEOUT = int(os.getenv('SSH_CONN_TIMEOUT', '30'))
SSH_MAX_CONNECTION_ATTEMPTS = 2


def _run_ssh_commands(host, ssh_email, ssh_password, commands):
    """Open a fresh SSH connection via Netmiko, send commands, return output.

    Returns the combined output string or ``None`` on failure.  On timeout errors
    that include the hint about ``conn_timeout`` the function will automatically
    retry with a doubled timeout (up to ``SSH_MAX_CONNECTION_ATTEMPTS``).
    """

    def _inner(conn_timeout):
        conn = None
        try:
            device = {
                'device_type': 'zte_zxros',
                'host': host,
                'username': ssh_email,
                'password': ssh_password,
                'secret': 'zxr10',
                'port': 22,
                'timeout': 30,
                'banner_timeout': 30,
                'auth_timeout': 30,
                'conn_timeout': conn_timeout,
                'session_log': None,
                'disabled_algorithms': {'pubkeys': []},
            }
            conn = ConnectHandler(**device)

            # Wait for OLT to be ready after connection
            time.sleep(7)

            # Enter enable mode: send 'enable' then 'zxr10' as password
            conn.send_command_timing('enable', read_timeout=10)
            time.sleep(1)
            conn.send_command_timing('zxr10', read_timeout=10)
            time.sleep(1)

            # Send show commands and collect output
            output = ''
            for cmd in commands:
                result = conn.send_command_timing(cmd, read_timeout=60, delay_factor=2)
                output += result + '\n'
                time.sleep(1)

            return output if output.strip() else None
        finally:
            try:
                if conn:
                    conn.disconnect()
            except Exception:
                pass

    attempt = 1
    current_timeout = SSH_CONN_TIMEOUT
    while attempt <= SSH_MAX_CONNECTION_ATTEMPTS:
        try:
            return _inner(current_timeout)
        except NetmikoTimeoutException as e:
            msg = str(e)
            # if netmiko/paramiko suggests increasing the connection timeout then
            # retry once with a larger value
            if 'No existing session' in msg and attempt < SSH_MAX_CONNECTION_ATTEMPTS:
                logging.warning(
                    "SSH to %s timed out (conn_timeout=%s). retrying with larger timeout",
                    host,
                    current_timeout,
                )
                attempt += 1
                current_timeout *= 2
                continue
            # otherwise give up and propagate error string
            return f"ERROR:{type(e).__name__}: {e}"
        except Exception as e:
            return f"ERROR:{type(e).__name__}: {e}"
    # should not reach here
    return None


def ssh_task(row_index, ticket, service_id, ne, location, ne_type, host, ssh_email, ssh_password, subject, session_id, mode='los', semaphore=None):
    def parse_db(value_raw):
        if value_raw is None:
            raise ValueError("Empty power value")
        v = str(value_raw)
        if "(" in v:
            v = v.split("(", 1)[0]
        v = v.replace("dBm", "").strip()
        return float(v)

    try:
        # Prepare show commands (enable/zxr10 handled by Netmiko)
        commands = []
        if ne_type in ["C600", "C620", "C650"]:
            commands.append(f"show pon power attenuation gpon_onu-1/{location}")
            commands.append(f"show gpon onu detail-info gpon_onu-1/{location}")
        else:
            commands.append(f"show pon power attenuation gpon-onu_1/{location}")
            commands.append(f"show gpon onu detail-info gpon-onu_1/{location}")

        output = _run_ssh_commands(host, ssh_email, ssh_password, commands)
        if not output:
            print_message(f"Ticket {ticket}: SSH returned no output (host {host})", session_id)
            output = ""
        elif output.startswith("ERROR:"):
            print_message(f"Ticket {ticket}: SSH failed on {host} - {output[6:]}", session_id)
            output = ""

        pattern = r"\s+up\s+Rx\s+:(.*?)\s+Tx:(.*?)\s+(.*?)\s+down\s+Tx\s+:(.*?)\s+Rx:(.*?)\s+(.*?)\s"
        matches = re.search(pattern, output, re.DOTALL)

        power_status = "Link currently down"
        ticket_status = "Error"
        comment = "No valid output"
        power_levels = ""
        powers=""

        if matches:
            olt_rx = matches.group(1).strip()
            onu_rx = matches.group(5).strip()
            powers = f"ONU Rx: {onu_rx} OLT Rx: {olt_rx}"
            power_levels = f"Ticket {ticket}: ONU Rx: {onu_rx} OLT Rx: {olt_rx}"
            print_message(power_levels, session_id)

            try:
                if onu_rx != "N/A" and olt_rx != "no signal":
                    onu_val = parse_db(onu_rx)
                    olt_val = parse_db(olt_rx)

                    if onu_val <= -28 or olt_val <= -27.5:
                        power_status = "Extreme Low Rx"
                    elif -28 < onu_val <= -24.5:
                        power_status = "Low Rx"
                    else:
                        power_status = "Power levels are optimal"
                else:
                    power_status = "Link currently down"
            except Exception as parse_err:
                print_message(f"Power parse error for ticket {ticket}: {parse_err}", session_id)
                power_status = "Invalid power values"

            if power_status == "Link currently down":
                comment = "LOS"
                ticket_status = "LOS"
            elif power_status in ["Low Rx", "Extreme Low Rx"]:
                comment = "Not Optimized"
                ticket_status = "Pending"
            elif power_status == "Power levels are optimal":
                comment = "Optimized, TT Closed"
                ticket_status = "Closed"
            else:
                comment = power_status
                ticket_status = "Error"

            return comment, power_status, ticket_status, power_levels

        print_message(f"No valid SSH output for ticket {ticket}", session_id)
        return "No valid output", "Error", "Error", ""

    except Exception as e:
        print_message(f"SSH error for ticket {ticket}: {str(e)}", session_id)
        logging.error(f"SSH task failed for ticket {ticket}", exc_info=True)
        return "Error", "Error", "Error", ""

    finally:
        if semaphore:
            try:
                semaphore.release()
            except Exception:
                pass

def check_if_already_processed(df, session_id):
    """
    Check if SSH processing results already exist for all rows.
    Returns True only if ALL rows have SSH results populated.
    """
    required_ssh_columns = ["Power Status", "Action", "Power Levels", "Powers"]
    missing_ssh_columns = [col for col in required_ssh_columns if col not in df.columns]
    
    if missing_ssh_columns:
        print_message(f"ℹ️ SSH columns not found ({', '.join(missing_ssh_columns)}). SSH processing will be performed.", session_id)
        return False
    
    # Check if ALL rows have SSH results (not just any row)
    # A row is considered "already processed" if it has values in all SSH columns
    rows_with_all_ssh_data = (
        df["Power Status"].notna() & 
        df["Action"].notna() & 
        df["Powers"].notna() & 
        df["Power Levels"].notna()
    )
    
    total_rows = len(df)
    processed_rows = rows_with_all_ssh_data.sum()
    
    if processed_rows == total_rows:
        print_message(f"✅ All {total_rows} rows already contain SSH processing results. Skipping SSH step.", session_id)
        return True
    elif processed_rows > 0:
        print_message(f"⚠️ Partial SSH results detected: {processed_rows}/{total_rows} rows processed. Running SSH for remaining {total_rows - processed_rows} rows.", session_id)
        return False
    else:
        print_message(f"ℹ️ SSH columns exist but are empty. SSH processing will be performed for all {total_rows} rows.", session_id)
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

        # Validate all required columns are present
        required_cols = ["ID", "Subscription", "Subject", "Creation", "Category 1", "Category 3", "Customer", "_Assign"]
        missing_required = [col for col in required_cols if col not in df.columns]
        if missing_required:
            print_message(f"Error: Required columns missing: {', '.join(missing_required)}", session_id)
            return None
        
        ssh_cols = ["Power Status", "Action", "Power Levels", "Powers"]
        all_cols = required_cols + ssh_cols
        existing_cols = [c for c in all_cols if c in df.columns]
        df = df[existing_cols]
        print_message(f"Filtered to {len(df)} rows ready for processing", session_id)
        return df
    except Exception as e:
        print_message(f"Filter error: {str(e)}", session_id)
        return None

def los_process_ssh(df, username, password, session_id, mode='los'):
    if df is None or df.empty:
        print_message("Error: No data to process for SSH", session_id)
        return None, None, []
    
    # Validate required columns exist
    required_for_ssh = ["ID", "Subscription", "Subject"]
    missing = [col for col in required_for_ssh if col not in df.columns]
    if missing:
        print_message(f"Error: SSH processing requires missing columns: {', '.join(missing)}", session_id)
        return None, None, []
    
    olt_data = read_olt_data(session_id)
    if not olt_data:
        return None, None, []
    total_steps = len(df)
    completed = 0
    results = {}
    
    unique_hosts = {data[1] for data in olt_data.values()}
    per_host_limit = 5
    host_semaphores = {host: threading.Semaphore(per_host_limit) for host in unique_hosts}
    max_workers = 30

    print_message(f"SSH scheduler: hosts={len(unique_hosts)} per_host_limit={per_host_limit} max_workers={max_workers}", session_id)

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        sessions[session_id]['executor'] = executor
        future_to_info = {}
        for idx, row in df.iterrows():
            try:
                ticket = row["ID"]
                service_id = row["Subscription"]
                subject = row["Subject"]
                
                # Skip rows that already have SSH results (optimization)
                if (pd.notna(row.get("Power Status")) and 
                    pd.notna(row.get("Action")) and 
                    pd.notna(row.get("Powers")) and 
                    pd.notna(row.get("Power Levels"))):
                    print_message(f"Ticket {ticket}: Skipping - SSH results already exist", session_id)
                    results[idx] = {
                        "comment": row.get("Ticket Comment", "Already processed"),
                        "status": row.get("Power Status", ""),
                        "ticket_status": row.get("Ticket Status", ""),
                        "power": row.get("Powers", "")
                    }
                    completed += 1
                    update_progress(completed, total_steps, f"Processing SSH ({completed}/{total_steps})", session_id)
                    continue
                
                # Skip rows with missing critical data
                if pd.isna(ticket) or pd.isna(service_id) or pd.isna(subject):
                    results[idx] = {"comment": "Missing critical data (ID/Subscription/Subject)", "status": "Error", "ticket_status": "Error", "power": ""}
                    continue
                
                ticket = str(ticket).strip()
                service_id = str(service_id).strip()
                subject = str(subject).strip()
            except (KeyError, AttributeError) as col_err:
                print_message(f"Column access error at row {idx}: {str(col_err)}", session_id)
                results[idx] = {"comment": f"Data error: {str(col_err)}", "status": "Error", "ticket_status": "Error", "power": ""}
                continue
            
            ne, location = extract_ne_and_location(subject)
            if ne and location:
                cleaned_ne = ne.replace("RB - ", "").strip().lower()
                if cleaned_ne in olt_data:
                    ne_type, host = olt_data[cleaned_ne]
                    semaphore = host_semaphores.get(host)
                    if "ridgeways_248.53" in cleaned_ne:
                        ssh_email = "noc"
                        ssh_password = "Faiba@543"
                    elif cleaned_ne in ["karen olt 1", "kericho olt"]:
                        ssh_email = "support"
                        ssh_password = "Support@2024"
                    else:
                        ssh_email = username
                        ssh_password = password
                    if semaphore:
                        semaphore.acquire()
                    try:
                        future = executor.submit(
                            ssh_task, idx, ticket, service_id, cleaned_ne, location, ne_type, host,
                            ssh_email, ssh_password, subject, session_id, 'los', semaphore
                        )
                        future_to_info[future] = (idx, semaphore)
                    except Exception as submit_err:
                        print_message(f"Failed to submit SSH task for ticket {ticket}: {str(submit_err)}", session_id)
                        results[idx] = {"comment": f"Task submission failed: {str(submit_err)}", "status": "Error", "ticket_status": "Error", "power": ""}
                        if semaphore:
                            semaphore.release()
                else:
                    results[idx] = {"comment": "NE not found in OLT data", "status": "Error", "ticket_status": "Error", "power": ""}
            else:
                results[idx] = {"comment": "Could not parse NE/Location", "status": "Error", "ticket_status": "Error", "power": ""}

        for future in concurrent.futures.as_completed(future_to_info):
            idx, semaphore = future_to_info[future]
            try:
                comment, status, ticket_status, power = future.result()
                results[idx] = {
                    "comment": comment, "status": status, "ticket_status": ticket_status, "power": power
                }
            except Exception as e:
                results[idx] = {
                    "comment": str(e), "status": "Error", "ticket_status": "Error", "power": ""
                }
            completed += 1
            update_progress(completed, total_steps, f"Processing SSH ({completed}/{total_steps})", session_id)

    for idx, res in results.items():
        if idx in df.index:
            df.at[idx, "Ticket Comment"] = res.get("comment", "")
            df.at[idx, "Status"] = res.get("status", "")
            df.at[idx, "Ticket Status"] = res.get("ticket_status", "")
            # 'power' key comes from the SSH task result payload; write consistently
            df.at[idx, "Power Levels"] = res.get("power", "")
            df.at[idx, "Powers"] = res.get("power", "")
            df.at[idx, "Power Status"] = res.get("status", "")
            # Conditional action assignment based on mode
            if res["status"] == "Power levels are optimal":
                df.at[idx, "Action"] = "Process"
            elif mode == 'los' and res["status"] in ["Low Rx", "Extreme Low Rx"]:
                # LOS mode: transition tickets with Low/Extreme Low Rx
                df.at[idx, "Action"] = "Transition"
            else:
                # LOWRX mode or other power statuses: skip
                df.at[idx, "Action"] = "Skip"

    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"los_processed_{uuid.uuid4()}.xlsx")
    df.to_excel(output_path, index=False)
    print_message(f"SSH processing complete. Updated file: {output_path}", session_id)

    to_process = [r for r in results.values() if r["status"] == "Power levels are optimal"]
    return df, output_path, to_process

def lowrx_load_and_filter(file_path, session_id):
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
            cat1_norm = df[col_category1].str.lower().fillna('')
            df = df[cat1_norm.str.contains("low rx")]

        if col_category3 in df.columns:
            cat3_norm = df[col_category3].str.lower().fillna('')
            df = df[cat3_norm.str.contains("low rx")]

        if col_subject in df.columns:
            subject_norm = df[col_subject].str.lower().fillna('')
            df = df[subject_norm.str.contains("low rx")]

        # Validate all required columns are present
        required_cols = ["ID", "Subscription", "Subject", "Creation", "Category 1", "Category 3", "Customer", "_Assign"]
        missing_required = [col for col in required_cols if col not in df.columns]
        if missing_required:
            print_message(f"Error: Required columns missing: {', '.join(missing_required)}", session_id)
            return None
        
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
    # Same logic as LOS SSH but with mode='lowrx' for different action assignment
    return los_process_ssh(df, username, password, session_id, mode='lowrx')

# Double Tickets logic
ALLOWED_CATEGORY1 = {
    "LOS Fiber Cut-Cable Maintenance Incident",
    "Extreme Low RX-Cable Maintenance Incident",
    "Low RX-Cable Maintenance Incident",
    "GPON LOSi-Cable Maintenance Incident",
    "GPON LOFi-Cable Maintenance Incident",
    "GPON Sfi-Cable Maintenance Incident",
}

def safe_to_datetime(series, colname="Creation"):
    s = pd.to_datetime(series, errors="coerce")
    if s.isna().all():
        return None
    return s

def check_duplicate_tickets_for_subscription(df, subscription, current_ticket_id, current_category1):
    current_base_category = None
    if pd.notna(current_category1):
        current_cat_lower = str(current_category1).lower()
        # Change: Treat "Extreme Low Rx" and "Low Rx" as distinct categories for duplication check
        if "extreme low rx" in current_cat_lower:
            current_base_category = "extreme_low_rx"
        elif "low rx" in current_cat_lower:
            current_base_category = "low_rx"
        elif "los" in current_cat_lower:
            current_base_category = "los"
        # Other categories like lofi, losi, sfi are not part of the transition logic,
        # so we can ignore them here to be more specific.

    if current_base_category is None:
        return None, []

    same_subscription = df[df['Subscription'] == subscription].copy()

    def matches_category(cat1):
        if pd.isna(cat1):
            return False
        cat_lower = str(cat1).lower()
        # Change: Perform an exact category match instead of a broad "contains" check.
        if current_base_category == "extreme_low_rx":
            return "extreme low rx" in cat_lower
        elif current_base_category == "low_rx":
            return "low rx" in cat_lower and "extreme" not in cat_lower
        return False # Only check for low rx types in this context

    same_subscription['matches_category'] = same_subscription['Category 1'].apply(matches_category)
    duplicates = same_subscription[same_subscription['matches_category']]

    if len(duplicates) <= 1:
        return None, []

    if 'Creation' in duplicates.columns:
        duplicates = duplicates.sort_values('Creation')

    oldest_ticket = duplicates.iloc[0]
    newer_tickets = duplicates.iloc[1:]

    return oldest_ticket, newer_tickets

def double_load_and_filter(file_path, apply_rules, session_id):
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
            if "Customer" in df.columns:
                mask = ~df["Customer"].astype(str).str.contains("JAMIIL LIMITED", case=False, na=False)
                df = df[mask]
            if "Category 3" in df.columns:
                mask = df["Category 3"].astype(str).str.strip().str.upper().str.startswith("GPON")
                df = df[mask]
            if "Category 1" in df.columns:
                df = df[df["Category 1"].astype(str).isin(ALLOWED_CATEGORY1)]

        print_message("Checking for duplicate tickets...", session_id)

        processed_tickets = set()
        duplicate_groups = {}

        for idx, row in df.iterrows():
            ticket_id = row['ID']
            if ticket_id in processed_tickets:
                continue

            subscription = row['Subscription']
            category1 = row.get('Category 1', '')

            oldest_ticket, newer_tickets = check_duplicate_tickets_for_subscription(
                df, subscription, ticket_id, category1
            )

            if oldest_ticket is not None and len(newer_tickets) > 0:
                oldest_id = oldest_ticket['ID']
                processed_tickets.add(oldest_id)
                for _, newer_row in newer_tickets.iterrows():
                    processed_tickets.add(newer_row['ID'])

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

# ---------------------------------------------------------------------------
# UPDATED CRM UPDATE LOGIC (FROM OLD VERSION)
# ---------------------------------------------------------------------------

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

        # # Wait for the form to be visible
        # try:
        #     wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "form[data-doctype='Issue']")))
        #     time.sleep(0.1)  # Additional wait for form to fully render
        # except:
        #     pass  # Continue even if form selector not found

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

                # Replace "LOS" with power status and add a concise power-level snippet
                # Extract a short power string (ONU/OLT receive values) from `power_levels`
                powers_short = ""
                try:
                    if power_levels:
                        # Prefer the explicit 'ONU Rx: ... OLT Rx: ...' substring
                        m_short = re.search(r'ONU Rx:\s*[^\s,;]+.*?OLT Rx:\s*[^\s,;]+', power_levels)
                        if m_short:
                            powers_short = m_short.group(0).strip()
                        else:
                            # Strip any leading 'Ticket ...:' prefix if present
                            m_ticket = re.search(r'Ticket\s*[^:]+:\s*(.*)', power_levels)
                            powers_short = m_ticket.group(1).strip() if m_ticket else power_levels.strip()
                except Exception:
                    powers_short = power_levels if power_levels else ""

                new_subject = original_subject.replace("LOS", f"{status_label} {powers_short}")

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
                time.sleep(0.1)
                print_message(f"✅ Set resolution", session_id)
                print_message(f"✅ Resolution {resolution_text}", session_id)
                # Read back the updated resolution text from the Quill editor and log it
                #  try:
                #     # Prefer plaintext via innerText; fall back to textContent or innerHTML
                #     res_preview = driver.execute_script(
                #         "var editor = arguments[0]; return (editor.innerText || editor.textContent || editor.innerHTML)" ,
                #         resolution_box
                #     )
                #     if res_preview is None:
                #         res_preview = ""
                #     res_preview = str(res_preview).strip()
                #     # Log a preview (cap at 300 chars to avoid huge logs)
                #     preview_msg = res_preview if len(res_preview) <= 300 else res_preview[:300] + '...'
                #     print_message(f"Updated resolution: {preview_msg}", session_id)
                # except Exception as e:
                #     print_message(f"Could not read updated resolution: {e}", session_id)
                # break
            except Exception as e:
                if resolution_attempt < max_resolution_attempts - 1:
                    print_message(f"⚠️ Retry resolution update (attempt {resolution_attempt + 1}/{max_resolution_attempts})", session_id)
                    time.sleep(0.1)
                else:
                    raise Exception(f"Failed to update resolution: {str(e)}")

        # Save with retry
        max_save_attempts = 3
        for save_attempt in range(max_save_attempts):
            try:
                save_button = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@data-label='Save']")))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", save_button)
                time.sleep(0.1)
                driver.execute_script("arguments[0].click();", save_button)
                time.sleep(0.1)  # Wait longer for save to complete

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

def transition_lowrx_tickets(df, session_id):
    driver = sessions[session_id]['driver']
    wait = sessions[session_id]['wait']
    processed = 0
    skipped_as_newer = set()

    tickets_to_transition = df[df['Power Status'].isin(['Low Rx', 'Extreme Low Rx'])]
    total = len(tickets_to_transition)

    if total == 0:
        print_message("No 'Low Rx' or 'Extreme Low Rx' tickets to transition.", session_id)
        return

    print_message(f"Found {total} tickets to transition. Duplicate check is disabled.", session_id)
    
    for idx, row in tickets_to_transition.iterrows():
        if sessions[session_id].get('paused', False):
             while sessions[session_id].get('paused', False): time.sleep(1)
        if sessions[session_id].get('stopped', False): return

        try:
            ticket = row["ID"]
            power_status = row["Power Status"]
            power_levels = row["Power Levels"]
            original_subject = row.get("Subject", "")

            status_label = "Extreme Low Rx" if power_status == "Extreme Low Rx" else "Low Rx"
            category = extreme_lowrx_transition_category if power_status == "Extreme Low Rx" else lowrx_transition_category

            # No duplicate check, just transition the ticket
            resolution = f"{status_label} detected. {power_levels}. Ticket transitioned for further optimization."
            print_message(f"Ticket {ticket}: Transitioning to {status_label} category.", session_id)
            update_crm_ticket(
                driver, wait, ticket, "In Process", category, resolution, session_id,
                is_closing=False, power_status=power_status,
                original_subject=original_subject, power_levels=power_levels
            )

            processed += 1
            update_progress(processed, total, f"Transitioning ({processed}/{total})", session_id)

        except Exception as e:
            print_message(f"Failed to transition {ticket}: {str(e)}", session_id)

def los_lowrx_update_crm(df, session_id, mode):
    driver = sessions[session_id]['driver']
    wait = sessions[session_id]['wait']
    available = sessions[session_id]['available_categories']
    options = los_category_options if mode == 'los' else lowrx_category_options
    processed = 0
    total = len(df[df['Action'] == 'Process'])
    
    for idx, row in df[df['Action'] == 'Process'].iterrows():
        if sessions[session_id].get('paused', False):
             while sessions[session_id].get('paused', False): time.sleep(1)
        if sessions[session_id].get('stopped', False): return

        ticket = row["ID"]
        power_levels = row["Power Levels"]
        try:
            category = get_non_repeating_category(available, options)
            # Prefer the full Power Levels text without the leading 'Ticket <id>: ' prefix.
            # If that yields an empty string (e.g. power_levels only contained the prefix),
            # fall back to the concise `Powers` value (if present). Finally default to 'Optimized'.
            resolution = "Optimized"
            if power_levels:
                candidate = power_levels.replace(f"Ticket {ticket}: ", "").strip()
                if candidate:
                    resolution = candidate
                else:
                    # Try concise powers column (short receive powers) if available on the row
                    pv = row.get("Powers") if isinstance(row, dict) else row.get("Powers", None)
                    if pv and str(pv).strip():
                        resolution = str(pv).strip()
            else:
                # If power_levels empty, still try the Powers column
                pv = row.get("Powers") if isinstance(row, dict) else row.get("Powers", None)
                if pv and str(pv).strip():
                    resolution = str(pv).strip()

            update_crm_ticket(driver, wait, ticket, "Closed", category, resolution, session_id, is_closing=True)
            processed += 1
            update_progress(processed, total, f"CRM Update ({processed}/{total})", session_id)
        except Exception as e:
            print_message(f"Failed to update {ticket}: {str(e)}", session_id)

def double_update_crm(df, service_tickets, session_id):
    driver = sessions[session_id]['driver']
    wait = sessions[session_id]['wait']
    processed = 0
    total = sum(len(v['newer']) for v in service_tickets.values())
    
    # Specific category from Old Code
    category_value = "Double Ticket-GPON-QoS Affecting-Extreme Low RX-Cable Maintenance Incident"

    for key, bundle in service_tickets.items():
        old_ticket = bundle['oldest']['ID']
        for _, newer_row in bundle['newer'].iterrows():
            if sessions[session_id].get('paused', False):
                 while sessions[session_id].get('paused', False): time.sleep(1)
            if sessions[session_id].get('stopped', False): return

            ticket = newer_row['ID']
            try:
                resolution = f"Existing TT: {old_ticket}"
                update_crm_ticket(driver, wait, ticket, "Duplicate", category_value, resolution, session_id, is_closing=True)
                
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

# ---------------------------------------------------------------------------
# FLASK ROUTES
# ---------------------------------------------------------------------------
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
    # New option for LOS/LowRx mode
    los_processing_option = request.form.get('los_processing_option', 'all') # Default to 'all'
    
    thread = threading.Thread(
        target=process_mode,
        args=(mode, file_path, username, password, session_id, apply_rules, dry_run, los_processing_option)
    )
    thread.daemon = True
    thread.start()
    return jsonify({'session_id': session_id, 'status': 'started'})

def process_mode(mode, file_path, username, password, session_id, apply_rules, dry_run, los_processing_option):
    print_message(f"Starting {mode.upper()} mode processing...", session_id)
    update_progress(0, 100, "Initializing...", session_id)
    
    ssh_path = None  # Initialize ssh_path variable
    
    if mode in ['los', 'lowrx']:
        if mode == 'los':
            df_filtered = los_load_and_filter(file_path, session_id)
            if df_filtered is None: return
            if check_if_already_processed(df_filtered, session_id):
                # SSH already completed - use existing results
                df = df_filtered
                to_process = df[df['Action'] == 'Process'].to_dict('records')
                # No ssh_path when skipping SSH
            else:
                # Need to run SSH processing
                df, ssh_path, to_process = los_process_ssh(df_filtered, username, password, session_id)
        else:
            df_filtered = lowrx_load_and_filter(file_path, session_id)
            if df_filtered is None: return
            if check_if_already_processed(df_filtered, session_id):
                # SSH already completed - use existing results
                df = df_filtered
                to_process = df[df['Action'] == 'Process'].to_dict('records')
                # No ssh_path when skipping SSH
            else:
                # Need to run SSH processing
                df, ssh_path, to_process = lowrx_process_ssh(df_filtered, username, password, session_id)

        num_to_process = len(to_process)
        num_lowrx = len(df[df['Power Status'] == 'Low Rx'])
        num_extreme_lowrx = len(df[df['Power Status'] == 'Extreme Low Rx'])

        print_message(f"Summary: {num_to_process} optimal, {num_lowrx} Low Rx, {num_extreme_lowrx} Extreme", session_id)

        if num_to_process > 0 or num_lowrx > 0 or num_extreme_lowrx > 0:
            if not login_and_process(username, password, session_id):
                print_message("CRM login failed. Aborting.", session_id)
                return

            if num_to_process > 0 and los_processing_option in ['all', 'los_only']:
                print_message(f"{num_to_process} optimal tickets ready for closure...", session_id)
                los_lowrx_update_crm(df, session_id, mode)

            if (num_lowrx > 0 or num_extreme_lowrx > 0) and los_processing_option in ['all', 'transition_only']:
                print_message(f"Transitioning {num_lowrx + num_extreme_lowrx} tickets...", session_id)
                transition_lowrx_tickets(df, session_id)
        else:
            print_message("No tickets to process based on the selected criteria.", session_id)
            update_progress(100, 100, "Complete", session_id)
            return

    else: # Double Ticket Mode
        df, service_tickets = double_load_and_filter(file_path, apply_rules, session_id)
        if df is None or not service_tickets:
            print_message("No duplicates found.", session_id)
            return
        
        num_to_process = sum(len(v['newer']) for v in service_tickets.values())
        if dry_run:
            print_message(f"Dry run: Would process {num_to_process} tickets.", session_id)
            update_progress(100, 100, "Dry run complete", session_id)
            return

        print_message(f"Processing {num_to_process} duplicate tickets...", session_id)
        if not login_and_process(username, password, session_id):
            print_message("CRM login failed. Aborting.", session_id)
            return
        double_update_crm(df, service_tickets, session_id)

    update_progress(100, 100, "Complete", session_id)
    if session_id in sessions and 'driver' in sessions[session_id]:
        try:
            sessions[session_id]['driver'].quit()
        except: pass
        sessions[session_id].pop('driver', None)
    sessions.pop(session_id, None)

@app.route('/report', methods=['POST'])
def report():
    if 'file' not in request.files: return jsonify({'error': 'No file'})
    file = request.files['file']
    if file.filename == '': return jsonify({'error': 'No file selected'})
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
    print_message(f"Starting {report_type.upper()} report...", session_id)
    update_progress(0, 100, "Initializing...", session_id)
    try:
        in_path = Path(file_path)
        out_filename = f"{in_path.stem}_{report_type}_report_{uuid.uuid4().hex[:8]}.xlsx"
        out_path = Path(app.config['UPLOAD_FOLDER']) / out_filename

        if report_type == 'gpon':
            stats = generate_report.process_workbook(in_path, out_path)
        elif report_type == 'enterprise':
            stats = enterprise_report.process_workbook(in_path, out_path)
        else:
            print_message("Error: Unknown report type", session_id)
            return

        print_message(f"Report generated: {out_path.name}", session_id)

        # Build a user-friendly filename like: "13th January Morning GPON Report.xlsx"
        def _ordinal(n: int) -> str:
            if 10 <= (n % 100) <= 20:
                suf = 'th'
            else:
                suf = {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')
            return f"{n}{suf}"

        def _time_segment(now: datetime) -> str:
            h = now.hour
            if 5 <= h <= 11:
                return 'Morning'
            if 12 <= h <= 16:
                return 'Afternoon'
            if 17 <= h <= 20:
                return 'Evening'
            return 'Night'

        now = datetime.now()
        day_label = _ordinal(now.day)
        month_label = now.strftime('%B')
        segment = _time_segment(now)
        type_label = report_type.upper() if report_type == 'gpon' else report_type.capitalize()
        friendly_base = f"{day_label} {month_label} {segment} {type_label} Report"
        friendly_name = secure_filename(friendly_base) + '.xlsx'
        dest = Path(app.config['UPLOAD_FOLDER']) / friendly_name
        # Avoid clobbering an existing file: add counter suffix if needed
        counter = 1
        while dest.exists():
            dest = Path(app.config['UPLOAD_FOLDER']) / secure_filename(f"{friendly_base} ({counter})")
            dest = dest.with_suffix('.xlsx')
            counter += 1

        try:
            shutil.copy2(out_path, dest)
            # Provide a download link (relative) to front-end; front-end can prefix host if needed
            print_message(f"Download report: /download/{dest.name}", session_id)
            # Emit structured event so front-end can automatically show a download button
            try:
                socketio.emit('report_ready', {'url': f"/download/{dest.name}", 'filename': dest.name}, room=session_id)
            except Exception:
                # Non-fatal: continue if emit fails
                pass
        except Exception as copy_err:
            print_message(f"Report ready but failed to create friendly copy: {copy_err}", session_id)
            # Still notify front-end of the original file so it can be downloaded
            try:
                if out_path.exists():
                    socketio.emit('report_ready', {'url': f"/download/{out_path.name}", 'filename': out_path.name}, room=session_id)
            except Exception:
                pass

        update_progress(100, 100, "Complete", session_id)
    except Exception as e:
        print_message(f"Error: {e}", session_id)
        update_progress(100, 100, "Error", session_id)
    finally:
        pass


@socketio.on('join_room')
def on_join(data):
    join_room(data['room'])

@app.route('/download/<path:filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

@app.route('/pause', methods=['POST'])
def pause_processing():
    data = request.json
    session_id = data.get('session_id')
    if session_id and session_id in sessions:
        sessions[session_id]['paused'] = True
        return jsonify({'status': 'success'})
    return jsonify({'status': 'error'})

@app.route('/resume', methods=['POST'])
def resume_processing():
    data = request.json
    session_id = data.get('session_id')
    if session_id and session_id in sessions:
        sessions[session_id]['paused'] = False
        return jsonify({'status': 'success'})
    return jsonify({'status': 'error'})

@app.route('/stop', methods=['POST'])
def stop_processing():
    data = request.json
    session_id = data.get('session_id')
    if session_id in sessions:
        session_data = sessions.get(session_id)
        if session_data:
            session_data['stopped'] = True
            
            # Shutdown the thread pool executor
            if 'executor' in session_data:
                try:
                    # cancel_futures is available in Python 3.9+
                    session_data['executor'].shutdown(wait=False, cancel_futures=True)
                    print_message("SSH process scheduler stopped.", session_id)
                except Exception as e:
                    print_message(f"Error stopping SSH scheduler: {e}", session_id)

            # Quit the browser
            if 'driver' in session_data:
                try:
                    session_data['driver'].quit()
                    print_message("Browser session closed.", session_id)
                except Exception as e:
                    print_message(f"Error closing browser on stop: {e}", session_id)
        
        # Finally, remove the session
        sessions.pop(session_id, None)
        print_message("Session cleared.", session_id)
        socketio.emit('processing_stopped', {'message': 'Processing stopped and session cleared.'}, room=session_id)

    return jsonify({'status': 'success'})

if __name__ == '__main__':
    # For development only - use WSGI server (gunicorn) for production
    socketio.run(app, host='0.0.0.0', port=5000, debug=False)
    