from flask import Flask, render_template, request, jsonify, send_file
from flask_socketio import SocketIO, emit, join_room
import os
import sys
import threading
import time
import logging
import signal
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
import tempfile
from datetime import datetime

# Add reportgen to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'Report Gen'))
import app as reportgen_module
generate_report = reportgen_module
enterprise_report = reportgen_module

app = Flask(__name__)
app.secret_key = 'gpon_assistant_secret_key_change_me'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

socketio = SocketIO(app, cors_allowed_origins="*", async_mode='threading')

# {session_id: {'driver': None, 'wait': None, 'available_categories': [], 'paused': False, 'stopped': False, 'ssh_errors': []}}
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
OLT_DATA_FILE = "OLT_DATA1.xlsx"





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

def analyze_ssh_error(error_msg, host, ticket=None, olt_data=None):
    """Analyze SSH error and return error details with probable cause."""
    # Look up OLT name from host IP
    olt_name = 'Unknown'
    if olt_data:
        for ne_name, (ne_type, ne_host) in olt_data.items():
            if ne_host == host:
                olt_name = ne_name
                break

    error_details = {
        'timestamp': datetime.now().isoformat(),
        'ticket': ticket,
        'host': host,
        'olt_name': olt_name,
        'error_type': 'Unknown',
        'error_message': str(error_msg),
        'probable_cause': 'Unknown error',
        'severity': 'Medium'
    }

    error_str = str(error_msg).lower()

    # Connection timeout errors
    if 'netmikotimeoutexception' in error_str or 'no existing session' in error_str:
        error_details['error_type'] = 'Connection Timeout'
        error_details['probable_cause'] = 'Network connectivity issue, OLT overloaded, or firewall blocking SSH'
        error_details['severity'] = 'High'

    # Authentication errors
    elif 'authentication failed' in error_str or 'permission denied' in error_str:
        error_details['error_type'] = 'Authentication Failure'
        error_details['probable_cause'] = 'Incorrect SSH credentials or account locked/disabled'
        error_details['severity'] = 'High'

    # Host unreachable errors
    elif 'connection refused' in error_str or 'no route to host' in error_str:
        error_details['error_type'] = 'Host Unreachable'
        error_details['probable_cause'] = 'OLT device unreachable, IP address incorrect, or device powered off'
        error_details['severity'] = 'High'

    # SSH protocol errors
    elif 'ssh' in error_str and ('protocol' in error_str or 'key' in error_str):
        error_details['error_type'] = 'SSH Protocol Error'
        error_details['probable_cause'] = 'SSH service disabled, incompatible SSH version, or key exchange failure'
        error_details['severity'] = 'Medium'

    # Command execution errors
    elif 'command not found' in error_str or 'invalid command' in error_str:
        error_details['error_type'] = 'Command Execution Error'
        error_details['probable_cause'] = 'Incorrect OLT type/model or firmware incompatibility'
        error_details['severity'] = 'Medium'

    # Parsing errors
    elif 'parse' in error_str or 'format' in error_str:
        error_details['error_type'] = 'Data Parsing Error'
        error_details['probable_cause'] = 'Unexpected OLT response format or firmware version differences'
        error_details['severity'] = 'Low'

    return error_details

def record_ssh_error(session_id, error_details):
    """Record SSH error for the session."""
    if session_id not in sessions:
        return

    if 'ssh_errors' not in sessions[session_id]:
        sessions[session_id]['ssh_errors'] = []

    sessions[session_id]['ssh_errors'].append(error_details)

    # Emit error stats update to frontend
    error_stats = get_ssh_error_stats(session_id)
    socketio.emit('ssh_error_update', error_stats, room=session_id)

def get_ssh_error_stats(session_id):
    """Get SSH error statistics for the session."""
    if session_id not in sessions or 'ssh_errors' not in sessions[session_id]:
        return {'total_errors': 0, 'errors_by_type': {}, 'errors_by_host': {}, 'errors_by_olt': {}, 'recent_errors': []}

    errors = sessions[session_id]['ssh_errors']
    stats = {
        'total_errors': len(errors),
        'errors_by_type': {},
        'errors_by_host': {},
        'errors_by_olt': {},
        'errors_by_severity': {},
        'recent_errors': errors[-10:]  # Last 10 errors
    }

    for error in errors:
        # Count by error type
        error_type = error.get('error_type', 'Unknown')
        stats['errors_by_type'][error_type] = stats['errors_by_type'].get(error_type, 0) + 1

        # Count by host
        host = error.get('host', 'Unknown')
        stats['errors_by_host'][host] = stats['errors_by_host'].get(host, 0) + 1

        # Count by OLT name
        olt_name = error.get('olt_name', 'Unknown')
        stats['errors_by_olt'][olt_name] = stats['errors_by_olt'].get(olt_name, 0) + 1

        # Count by severity
        severity = error.get('severity', 'Medium')
        stats['errors_by_severity'][severity] = stats['errors_by_severity'].get(severity, 0) + 1

    return stats

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

def read_olt_data(session_id=None):
    """Read OLT_DATA1.xlsx once and cache it."""
    global OLT_DATA_CACHE, OLT_DATA_MTIME

    file_path = OLT_DATA_FILE
    if not os.path.exists(file_path):
        if session_id:
            print_message(f"Error: {OLT_DATA_FILE} not found in root directory!", session_id)
        return {}

    try:
        mtime = os.path.getmtime(file_path)
    except OSError as e:
        if session_id:
            print_message(f"Error reading OLT data mtime: {e}", session_id)
        return {}

    with OLT_DATA_LOCK:
        if OLT_DATA_CACHE is not None and OLT_DATA_MTIME == mtime:
            if session_id:
                print_message("OLT data loaded from cache.", session_id)
            return OLT_DATA_CACHE

        try:
            df = pd.read_excel(file_path, dtype=str)
            required_cols = {"NE", "Type", "Host"}
            missing = required_cols - set(df.columns)
            if missing:
                if session_id:
                    print_message(f"Error: OLT data missing columns: {', '.join(missing)}", session_id)
                return {}

            ne_data = {}
            for _, row in df.iterrows():
                ne_raw = row.get("NE", "")
                if pd.isna(ne_raw) or str(ne_raw).strip() == "":
                    continue
                ne = str(ne_raw).strip().lower()
                ne_type = row.get("Type", "")
                host = row.get("Host", "")
                if pd.isna(host) or str(host).strip() == "":
                    continue
                ne_data[ne] = (str(ne_type).strip(), str(host).strip())

            OLT_DATA_CACHE = ne_data
            OLT_DATA_MTIME = mtime
            if session_id:
                print_message(f"OLT data loaded successfully ({len(ne_data)} NE entries).", session_id)
            return ne_data

        except Exception as e:
            if session_id:
                print_message(f"Error reading OLT data: {str(e)}", session_id)
            return {}


def _normalize_olt_value(value):
    return str(value).strip()


def _normalize_olt_ne(value):
    return str(value).strip().lower()


def add_olt_entry(ne, ne_type, host, session_id=None):
    file_path = OLT_DATA_FILE
    with OLT_DATA_LOCK:
        if os.path.exists(file_path):
            try:
                df = pd.read_excel(file_path, dtype=str)
            except Exception as e:
                if session_id:
                    print_message(f"Error loading existing OLT file: {e}", session_id)
                return False, f"Failed to read existing OLT database: {e}", 0
        else:
            df = pd.DataFrame(columns=["NE", "Type", "Host"])

        df = df.astype(object).where(pd.notnull(df), "")
        df["NE"] = df["NE"].apply(_normalize_olt_value)
        df["Type"] = df["Type"].apply(_normalize_olt_value)
        df["Host"] = df["Host"].apply(_normalize_olt_value)

        new_ne = _normalize_olt_value(ne)
        new_type = _normalize_olt_value(ne_type)
        new_host = _normalize_olt_value(host)
        if not new_ne or not new_type or not new_host:
            return False, "NE, Type and Host are all required.", len(df)

        normalized_existing = df["NE"].str.lower().tolist()
        if new_ne.lower() in normalized_existing:
            return False, f"OLTs entry for '{new_ne}' already exists.", len(df)

        new_row = pd.DataFrame([{"NE": new_ne, "Type": new_type, "Host": new_host}])
        df = pd.concat([df, new_row], ignore_index=True)

        try:
            df.to_excel(file_path, index=False)
            # Reset cached data so next read reflects the new entry
            global OLT_DATA_CACHE, OLT_DATA_MTIME
            OLT_DATA_CACHE = None
            OLT_DATA_MTIME = None
            if session_id:
                print_message(f"Added new OLT entry: {new_ne} / {new_type} / {new_host}", session_id)
            return True, f"Added new OLT entry for '{new_ne}'.", len(df)
        except Exception as e:
            if session_id:
                print_message(f"Error saving OLT file: {e}", session_id)
            return False, f"Failed to save OLT entry: {e}", len(df)


def update_olt_entry(index, ne, ne_type, host, session_id=None):
    """Update an OLT entry by index."""
    file_path = OLT_DATA_FILE
    try:
        index = int(index)
    except (ValueError, TypeError):
        return False, "Invalid OLT index.", 0
    
    with OLT_DATA_LOCK:
        if not os.path.exists(file_path):
            return False, "OLT database not found.", 0
        
        try:
            df = pd.read_excel(file_path, dtype=str)
        except Exception as e:
            if session_id:
                print_message(f"Error loading OLT file: {e}", session_id)
            return False, f"Failed to read OLT database: {e}", 0
        
        if index < 0 or index >= len(df):
            return False, f"OLT index {index} out of range.", len(df)
        
        new_ne = _normalize_olt_value(ne)
        new_type = _normalize_olt_value(ne_type)
        new_host = _normalize_olt_value(host)
        
        if not new_ne or not new_type or not new_host:
            return False, "NE, Type and Host are all required.", len(df)
        
        # Check for duplicates (excluding current entry)
        df_normalized = df.astype(object).where(pd.notnull(df), "")
        df_normalized["NE"] = df_normalized["NE"].apply(_normalize_olt_value)
        
        for i, existing_ne in enumerate(df_normalized["NE"]):
            if i != index and existing_ne.lower() == new_ne.lower():
                return False, f"OLT entry for '{new_ne}' already exists.", len(df)
        
        # Update the row
        df.loc[index, "NE"] = new_ne
        df.loc[index, "Type"] = new_type
        df.loc[index, "Host"] = new_host
        
        try:
            df.to_excel(file_path, index=False)
            # Reset cached data
            global OLT_DATA_CACHE, OLT_DATA_MTIME
            OLT_DATA_CACHE = None
            OLT_DATA_MTIME = None
            if session_id:
                print_message(f"Updated OLT entry {index}: {new_ne} / {new_type} / {new_host}", session_id)
            # Notify all connected sessions about OLT data change
            socketio.emit('olt_data_changed', {'action': 'updated', 'ne': new_ne}, namespace='/', room=None)
            return True, f"Updated OLT entry '{new_ne}'.", len(df)
        except Exception as e:
            if session_id:
                print_message(f"Error saving OLT file: {e}", session_id)
            return False, f"Failed to update OLT entry: {e}", len(df)


def delete_olt_entry(index, session_id=None):
    """Delete an OLT entry by index."""
    file_path = OLT_DATA_FILE
    try:
        index = int(index)
    except (ValueError, TypeError):
        return False, "Invalid OLT index.", 0
    
    with OLT_DATA_LOCK:
        if not os.path.exists(file_path):
            return False, "OLT database not found.", 0
        
        try:
            df = pd.read_excel(file_path, dtype=str)
        except Exception as e:
            if session_id:
                print_message(f"Error loading OLT file: {e}", session_id)
            return False, f"Failed to read OLT database: {e}", len(df) if 'df' in locals() else 0
        
        if index < 0 or index >= len(df):
            return False, f"OLT index {index} out of range.", len(df)
        
        # Get the NE name before deletion for logging
        olt_name = df.loc[index, "NE"] if "NE" in df.columns else f"Index {index}"
        
        # Delete the row
        df = df.drop(index).reset_index(drop=True)
        
        try:
            df.to_excel(file_path, index=False)
            # Reset cached data
            global OLT_DATA_CACHE, OLT_DATA_MTIME
            OLT_DATA_CACHE = None
            OLT_DATA_MTIME = None
            if session_id:
                print_message(f"Deleted OLT entry: {olt_name}", session_id)
            # Notify all connected sessions about OLT data change
            socketio.emit('olt_data_changed', {'action': 'deleted', 'ne': olt_name}, namespace='/', room=None)
            return True, f"Deleted OLT entry '{olt_name}'.", len(df)
        except Exception as e:
            if session_id:
                print_message(f"Error saving OLT file: {e}", session_id)
            return False, f"Failed to delete OLT entry: {e}", len(df)

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
# setting environment variables, e.g.:
#
#     export SSH_CONN_TIMEOUT=60
#     export SSH_READ_TIMEOUT=0   # 0 = read until device stops outputting (no cap)
#
#  The Netmiko library itself will suggest increasing 'conn_timeout' when a
#  ``NetmikoTimeoutException`` is raised with a message like "No existing
#  session".  _run_ssh_commands now catches that specific case and retries once
#  with a larger timeout.
SSH_CONN_TIMEOUT = int(os.getenv('SSH_CONN_TIMEOUT', '60'))  # increased from 30
SSH_READ_TIMEOUT = int(os.getenv('SSH_READ_TIMEOUT', '120'))   # changed from 0 to 120 seconds to prevent hanging
SSH_MAX_CONNECTION_ATTEMPTS = 3  # increased from 2 for more resilience
# Per-host connection stagger delay (seconds) to avoid flooding OLT TCP stack
SSH_HOST_STAGGER_DELAY = float(os.getenv('SSH_HOST_STAGGER_DELAY', '1.5'))


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
                'timeout': 60,           # increased from 30
                'banner_timeout': 60,    # increased from 30
                'auth_timeout': 60,      # increased from 30
                'conn_timeout': conn_timeout,
                'session_log': None,
                'disabled_algorithms': {'pubkeys': []},
            }
            conn = ConnectHandler(**device)

            # Wait for OLT to be ready after connection
            time.sleep(7)

            # Enter enable mode: send 'enable' then 'zxr10' as password
            conn.send_command_timing('enable', read_timeout=15)
            time.sleep(1)
            conn.send_command_timing('zxr10', read_timeout=15)
            time.sleep(1)

            # Send show commands and collect output.
            # read_timeout=120 means Netmiko will timeout after 2 minutes per command
            # This prevents hanging on unresponsive OLTs while still allowing
            # enough time for busy OLTs with large output tables.
            output = ''
            for cmd in commands:
                try:
                    result = conn.send_command_timing(
                        cmd,
                        read_timeout=SSH_READ_TIMEOUT,
                        delay_factor=3,   # extra delay between timing polls
                    )
                    output += result + '\n'
                except Exception as cmd_error:
                    # If a single command fails, log it but continue with other commands
                    logging.warning(f"Command '{cmd}' failed on {host}: {cmd_error}")
                    output += f"ERROR: Command '{cmd}' failed: {cmd_error}\n"
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
            # retry with a larger value (double each time)
            if 'No existing session' in msg and attempt < SSH_MAX_CONNECTION_ATTEMPTS:
                logging.warning(
                    "SSH to %s timed out (conn_timeout=%s, attempt %s/%s). retrying with larger timeout",
                    host,
                    current_timeout,
                    attempt,
                    SSH_MAX_CONNECTION_ATTEMPTS,
                )
                attempt += 1
                current_timeout = int(current_timeout * 1.5)  # gentler growth than doubling
                time.sleep(SSH_HOST_STAGGER_DELAY * attempt)  # back-off before retry
                continue
            # otherwise give up and propagate error string
            logging.error(f"SSH to {host} failed with NetmikoTimeoutException: {e}")
            return f"ERROR:{type(e).__name__}: {e}"
        except Exception as e:
            logging.error(f"SSH to {host} failed with exception: {e}")
            return f"ERROR:{type(e).__name__}: {e}"
    # should not reach here
    return None


def ssh_task(row_index, ticket, service_id, ne, location, ne_type, host, ssh_email, ssh_password, subject, session_id, mode='los', semaphore=None, olt_data=None):
    # Overall timeout for the entire SSH task (5 minutes)
    SSH_TASK_TIMEOUT = 300

    def timeout_handler():
        raise TimeoutError(f"SSH task timed out after {SSH_TASK_TIMEOUT} seconds for ticket {ticket}")

    timer = threading.Timer(SSH_TASK_TIMEOUT, timeout_handler)
    timer.start()

    def parse_db(value_raw):
        if value_raw is None:
            raise ValueError("Empty power value")
        v = str(value_raw)
        if "(" in v:
            v = v.split("(", 1)[0]
        v = v.replace("dBm", "").strip()
        return float(v)

    try:
        # Semaphore is acquired HERE (inside the worker thread) so the per-host
        # slot is only held for the actual SSH work, not while the task waits in
        # the executor queue.  This was previously acquired in the main thread,
        # which caused premature slot exhaustion and contributed to conn_timeout
        # failures when many tickets share the same OLT host.
        if semaphore:
            semaphore.acquire()
            # Small stagger so concurrent workers don't hammer the same OLT at
            # exactly the same moment.
            time.sleep(SSH_HOST_STAGGER_DELAY)
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
            error_msg = output[6:]
            print_message(f"Ticket {ticket}: SSH failed on {host} - {error_msg}", session_id)

            # Record SSH error
            error_details = analyze_ssh_error(error_msg, host, ticket, olt_data)
            record_ssh_error(session_id, error_details)

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

    except TimeoutError as e:
        # Handle task timeout specifically
        error_msg = str(e)
        print_message(f"SSH task timeout for ticket {ticket}: {error_msg}", session_id)
        logging.error(f"SSH task timeout for ticket {ticket}: {error_msg}")

        # Record timeout as SSH error
        error_details = analyze_ssh_error(error_msg, host, ticket, olt_data)
        record_ssh_error(session_id, error_details)

        return "Timeout Error", "Error", "Error", ""

    except Exception as e:
        print_message(f"SSH error for ticket {ticket}: {str(e)}", session_id)
        logging.error(f"SSH task failed for ticket {ticket}", exc_info=True)
        return "Error", "Error", "Error", ""

    finally:
        # Cancel the timeout timer
        try:
            timer.cancel()
        except Exception:
            pass

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
            df = df[~df[col_customer].isin(["JAMII TELECOMMUNICATIONS LIMITED", "JAMIIL LIMITED"])]
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
    # Limit to 3 concurrent SSH sessions per OLT host (was 5).
    # OLT devices have limited SSH server capacity; exceeding it causes
    # "No existing session" / conn_timeout failures.
    per_host_limit = 3
    host_semaphores = {host: threading.Semaphore(per_host_limit) for host in unique_hosts}
    max_workers = 20  # reduced from 30 to avoid OS thread/socket pressure

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
                        ssh_email = "support"
                        ssh_password = "Faiba@543"
                    if "RIDGEWAYS_C620_248.53" in cleaned_ne:
                        ssh_email = "support"
                        ssh_password = "Faiba@543"
                    elif cleaned_ne in ["karen olt 1", "kericho olt"]:
                        ssh_email = "support"
                        ssh_password = "Support@2024"
                    else:
                        ssh_email = username
                        ssh_password = password
                    # NOTE: semaphore.acquire() is now done INSIDE ssh_task
                    # (inside the worker thread) so the slot is only held
                    # while SSH work is actually in progress, not while the
                    # task waits in the executor queue.
                    try:
                        future = executor.submit(
                            ssh_task, idx, ticket, service_id, cleaned_ne, location, ne_type, host,
                            ssh_email, ssh_password, subject, session_id, 'los', semaphore, olt_data
                        )
                        future_to_info[future] = (idx, semaphore)
                    except Exception as submit_err:
                        print_message(f"Failed to submit SSH task for ticket {ticket}: {str(submit_err)}", session_id)
                        results[idx] = {"comment": f"Task submission failed: {str(submit_err)}", "status": "Error", "ticket_status": "Error", "power": ""}
                        # semaphore was not acquired here (it's acquired inside ssh_task),
                        # so do NOT release it on submission failure.
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
            df = df[~df[col_customer].isin(["JAMII TELECOMMUNICATIONS LIMITED", "JAMIIL LIMITED"])]

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
        driver.set_page_load_timeout(30)
        driver.get(ticket_url)

        # Wait for the page to actually finish loading by waiting for a known
        # element instead of a fixed sleep — the fixed 0.5 s was far too short
        # for the CRM and caused stale/missing element errors on slow networks.
        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.layout-main")))
        except Exception:
            time.sleep(3)  # fallback if selector doesn't match

        # Verify we're on the correct ticket page
        if ticket not in driver.current_url:
            raise Exception(f"Failed to navigate to ticket {ticket}")

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
                break  # success — exit the retry loop
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
                time.sleep(0.3)
                driver.execute_script("arguments[0].click();", save_button)
                # Wait for the save to complete: look for a success indicator or
                # wait for the save button to become stale/re-enabled.
                time.sleep(1.5)

                # Check for save success indicators
                try:
                    success_msg = driver.find_elements(By.CSS_SELECTOR, ".indicator.green, .msgprint")
                    if success_msg:
                        print_message(f"✅ Saved ticket {ticket} (success indicator found)", session_id)
                    else:
                        print_message(f"✅ Saved ticket {ticket}", session_id)
                except Exception:
                    print_message(f"✅ Saved ticket {ticket}", session_id)
                break
            except Exception as e:
                if save_attempt < max_save_attempts - 1:
                    print_message(f"Retry save (attempt {save_attempt + 1}/{max_save_attempts})", session_id)
                    time.sleep(1)
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
        if sessions.get(session_id, {}).get('paused', False):
            while sessions.get(session_id, {}).get('paused', False):
                time.sleep(1)
        if sessions.get(session_id, {}).get('stopped', False):
            return

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
        if sessions.get(session_id, {}).get('paused', False):
            while sessions.get(session_id, {}).get('paused', False):
                time.sleep(1)
        if sessions.get(session_id, {}).get('stopped', False):
            return

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
            if sessions.get(session_id, {}).get('paused', False):
                while sessions.get(session_id, {}).get('paused', False):
                    time.sleep(1)
            if sessions.get(session_id, {}).get('stopped', False):
                return

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
    sessions[session_id] = {'ssh_errors': []}
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    # Use .get() with defaults to avoid KeyError if a form field is missing
    mode = request.form.get('mode', 'los')
    username = request.form.get('username', '')
    password = request.form.get('password', '')
    apply_rules = request.form.get('apply_rules', 'on') == 'on'
    dry_run = request.form.get('dry_run', 'off') == 'on'
    # New option for LOS/LowRx mode
    los_processing_option = request.form.get('los_processing_option', 'all')  # Default to 'all'
    
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

        num_to_process = len(to_process) if to_process else 0
        num_lowrx = len(df[df['Power Status'] == 'Low Rx']) if df is not None and 'Power Status' in df.columns else 0
        num_extreme_lowrx = len(df[df['Power Status'] == 'Extreme Low Rx']) if df is not None and 'Power Status' in df.columns else 0

        print_message(f"Summary: {num_to_process} optimal, {num_lowrx} Low Rx, {num_extreme_lowrx} Extreme", session_id)

        if num_to_process > 0 or num_lowrx > 0 or num_extreme_lowrx > 0:
            if df is None:
                print_message("Error: No data available for CRM update.", session_id)
                return
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

@app.route('/olt/add', methods=['POST'])
def add_olt():
    ne = request.form.get('ne', '').strip()
    ne_type = request.form.get('type', '').strip()
    host = request.form.get('host', '').strip()

    if not ne or not ne_type or not host:
        return jsonify({'status': 'error', 'message': 'NE, Type and Host are required.', 'count': 0})

    success, message, count = add_olt_entry(ne, ne_type, host)
    if not success:
        return jsonify({'status': 'error', 'message': message, 'count': count})

    return jsonify({'status': 'success', 'message': message, 'ne': ne, 'count': count})

@app.route('/olt/count', methods=['GET'])
def olt_count():
    olt_data = read_olt_data()
    return jsonify({'count': len(olt_data)})

@app.route('/olt/list', methods=['GET'])
def list_olts():
    try:
        if not os.path.exists(OLT_DATA_FILE):
            return jsonify({'olts': []})
        df = pd.read_excel(OLT_DATA_FILE, dtype=str)
        olts = []
        for idx, row in df.iterrows():
            olts.append({
                'index': int(idx),
                'ne': str(row.get('NE', '')).strip(),
                'type': str(row.get('Type', '')).strip(),
                'host': str(row.get('Host', '')).strip()
            })
        return jsonify({'olts': olts})
    except Exception as e:
        return jsonify({'olts': [], 'error': str(e)})

@app.route('/olt/update/<int:index>', methods=['POST'])
def update_olt(index):
    ne = request.form.get('ne', '').strip()
    ne_type = request.form.get('type', '').strip()
    host = request.form.get('host', '').strip()

    if not ne or not ne_type or not host:
        return jsonify({'status': 'error', 'message': 'NE, Type and Host are required.'})

    success, message, count = update_olt_entry(index, ne, ne_type, host)
    if not success:
        return jsonify({'status': 'error', 'message': message, 'count': count})

    return jsonify({'status': 'success', 'message': message, 'count': count})

@app.route('/olt/delete/<int:index>', methods=['POST'])
def delete_olt(index):
    success, message, count = delete_olt_entry(index)
    if not success:
        return jsonify({'status': 'error', 'message': message, 'count': count})

    return jsonify({'status': 'success', 'message': message, 'count': count})

@app.route('/report', methods=['POST'])
def report():
    if 'file' not in request.files: return jsonify({'error': 'No file'})
    file = request.files['file']
    if file.filename == '': return jsonify({'error': 'No file selected'})
    session_id = str(uuid.uuid4())
    sessions[session_id] = {'ssh_errors': []}
    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(file_path)
    report_type = request.form.get('report_type', 'gpon')
    if report_type not in ('gpon', 'enterprise'):
        return jsonify({'error': f"Unknown report_type: {report_type}"})
    thread = threading.Thread(target=process_report_mode, args=(report_type, file_path, session_id))
    thread.daemon = True
    thread.start()
    return jsonify({'session_id': session_id, 'status': 'started'})


@app.route('/report_preview', methods=['POST'])
def report_preview():
    """Lightweight preview using the report engine *_preview() to show stats before full generation."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file'})
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'})
    report_type = request.form.get('report_type', 'gpon')
    if report_type not in ('gpon', 'enterprise'):
        return jsonify({'error': f"Unknown report_type: {report_type}"})

    filename = secure_filename(file.filename)
    # Use temp file for preview only (deleted after)
    suffix = os.path.splitext(filename)[1] or '.xlsx'
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, dir=app.config['UPLOAD_FOLDER']) as tmp:
        tmp_path = tmp.name
        file.save(tmp_path)

    in_path = Path(tmp_path)
    try:
        if report_type == 'gpon':
            info = generate_report.gpon_preview(in_path)
        else:
            info = generate_report.enterprise_preview(in_path)
        info['report_type'] = report_type
        return jsonify({'ok': True, 'stats': info})
    except Exception as e:
        return jsonify({'error': str(e)})
    finally:
        try:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)
        except Exception:
            pass


def process_report_mode(report_type, file_path, session_id):
    print_message(f"Starting {report_type.upper()} report...", session_id)
    update_progress(0, 100, "Initializing...", session_id)

    # Progress adapter from the report engine to live UI updates
    _phase_weights = {
        'load': (0, 10), 'filter': (10, 25), 'gpon': (25, 45), 'enterprise': (25, 45),
        'buckets': (45, 60), 'double': (55, 65), 'styling': (60, 80), 'summary': (75, 85),
        'write': (80, 95), 'done': (95, 100)
    }
    def _report_progress(phase: str, current: int = 0, total: int = 1):
        try:
            base = _phase_weights.get(phase.lower(), (40, 80))
            sub = int((current / max(1, total)) * (base[1] - base[0])) if total else 0
            pct = min(99, base[0] + sub)
            msg = f"{phase.replace('_',' ').title()} ({current}/{total})"
            update_progress(pct, 100, msg, session_id)
            if phase in ('write', 'done', 'styling'):
                print_message(f"Engine: {phase}...", session_id)
        except Exception:
            pass

    try:
        in_path = Path(file_path)
        out_filename = f"{in_path.stem}_{report_type}_report_{uuid.uuid4().hex[:8]}.xlsx"
        out_path = Path(app.config['UPLOAD_FOLDER']) / out_filename

        if report_type == 'gpon':
            stats = generate_report.gpon_process(in_path, out_path, progress_callback=_report_progress)
        elif report_type == 'enterprise':
            stats = enterprise_report.enterprise_process(in_path, out_path, progress_callback=_report_progress)
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

@app.route('/export_ssh_errors', methods=['POST'])
def export_ssh_errors():
    """Export SSH error data to Excel file."""
    data = request.json
    session_id = data.get('session_id')

    if not session_id or session_id not in sessions:
        return jsonify({'error': 'Invalid session'})

    errors = sessions[session_id].get('ssh_errors', [])
    if not errors:
        return jsonify({'error': 'No SSH errors to export'})

    # Convert errors to DataFrame
    df = pd.DataFrame(errors)

    # Reorder columns for better readability
    columns = ['timestamp', 'ticket', 'host', 'olt_name', 'error_type', 'severity', 'probable_cause', 'error_message']
    df = df[columns]

    # Generate filename with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"ssh_errors_{timestamp}.xlsx"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

    # Export to Excel
    df.to_excel(filepath, index=False)

    return jsonify({'filename': filename, 'url': f'/download/{filename}'})

@app.route('/get_ssh_errors', methods=['POST'])
def get_ssh_errors():
    """Get SSH error statistics for a session."""
    data = request.json
    session_id = data.get('session_id')

    if not session_id:
        return jsonify({'error': 'No session ID provided'})

    error_stats = get_ssh_error_stats(session_id)
    return jsonify(error_stats)

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
    