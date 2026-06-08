# GPON Project - Windows ChromeDriver Troubleshooting Guide

## Issue: "WinError 193 - not a valid Win32 application"

This error occurs when Selenium fails to execute the ChromeDriver on Windows. Common causes:

1. **Wrong Python Architecture** - 32-bit Python trying to run 64-bit ChromeDriver (or vice versa)
2. **Corrupted ChromeDriver Download** - Cache contains invalid binary
3. **Missing Chrome Browser** - Selenium can't find Chrome installation
4. **Chrome Installation Issues** - Chrome is not properly installed

## Quick Fix (Recommended)

### Option 1: Automatic Fix Script (Easiest)

Run the included fix script:

```batch
python fix_chromedriver.py
```

Or use the convenience script:

```batch
fix_and_run.bat
```

This will:
- Clear the ChromeDriver cache
- Verify Chrome is installed
- Check Python architecture
- Upgrade Selenium and WebDriver Manager
- Start the application

### Option 2: Manual Fix Steps

#### Step 1: Clear ChromeDriver Cache

**Windows:**
```powershell
# Open PowerShell and run:
Remove-Item -Path "$env:USERPROFILE\.wdm" -Recurse -Force -ErrorAction SilentlyContinue
```

**Linux/macOS:**
```bash
rm -rf ~/.wdm
```

#### Step 2: Verify Python Architecture

```bash
python -c "import struct; print(f'{struct.calcsize(\"P\") * 8}-bit')"
```

Should show `64-bit` (recommended) or `32-bit`.

#### Step 3: Install 64-bit Python (if needed)

If you're running 32-bit Python, download and install 64-bit Python:
- https://www.python.org/downloads/

#### Step 4: Verify Chrome Installation

Make sure Google Chrome is installed:
- https://www.google.com/chrome/

**Windows:** Check in one of these paths:
- `C:\Program Files\Google\Chrome\Application\chrome.exe`
- `C:\Program Files (x86)\Google\Chrome\Application\chrome.exe`
- `%APPDATA%\Local\Google\Chrome\Application\chrome.exe`

#### Step 5: Reinstall Dependencies

```bash
pip install --upgrade selenium webdriver-manager
```

#### Step 6: Run the Application

```batch
python -m waitress --port=5000 --host=0.0.0.0 wsgi:application
```

## Advanced Troubleshooting

### Enable Debug Logging

Edit `app.py` and add more verbose logging:

```python
import logging
logging.basicConfig(
    level=logging.DEBUG,  # Changed from INFO
    filename='app.log',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s'
)
```

### Check Chrome Version Compatibility

```bash
# Check your Chrome version:
# Windows: Go to Chrome menu > Settings > About Chrome
# It should be version 120+ for best compatibility

# Install specific ChromeDriver version:
python -c "from webdriver_manager.chrome import ChromeDriverManager; print(ChromeDriverManager().install())"
```

### Use Chrome in Headless Mode (if needed)

If Chrome window causes issues, enable headless mode in `app.py`:

Find this line in the `initialize_chrome_driver` function:
```python
# chrome_options.add_argument("--headless") # Optional: Keep commented for visibility
```

Change to:
```python
chrome_options.add_argument("--headless")  # Enable headless mode
```

### Check System Requirements

Verify your system meets these requirements:

- **Python:** 3.8+ (64-bit recommended)
- **Chrome:** Latest version from https://www.google.com/chrome/
- **RAM:** At least 2GB free
- **Disk Space:** At least 500MB free in %APPDATA%

### View Detailed Error Log

```bash
# Check the application log for more details:
type app.log

# On Linux/macOS:
cat app.log
```

## Still Having Issues?

### 1. Full Clean Reinstall

```bash
# Remove Python virtual environment (if using one)
rmdir /s venv

# Clear pip cache
pip cache purge

# Remove ChromeDriver cache
Remove-Item -Path "$env:USERPROFILE\.wdm" -Recurse -Force -ErrorAction SilentlyContinue

# Reinstall all dependencies
pip install -r requirements.txt

# Run the application
python -m waitress --port=5000 --host=0.0.0.0 wsgi:application
```

### 2. Use Virtual Environment (Recommended)

```bash
# Create virtual environment
python -m venv venv

# Activate it
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run application
python -m waitress --port=5000 --host=0.0.0.0 wsgi:application
```

### 3. Alternative: Use Development Server

If WSGI server continues to fail, fall back to Flask's development server:

```bash
python app.py
```

**Note:** This is not recommended for production but useful for testing.

## Windows 10/11 Specific Issues

### If Getting Permission Errors:

```powershell
# Run as Administrator:
# 1. Right-click PowerShell
# 2. Select "Run as Administrator"
# 3. Run the application command
```

### If Chrome Window Won't Display:

Add this option to `chrome_options` in `initialize_chrome_driver`:

```python
chrome_options.add_argument("--window-size=1440,900")
```

## Performance Tips

1. **Disable GPU** (already enabled in code)
2. **Disable Sandbox** (already enabled for headless mode)
3. **Disable Extensions** - Not typically an issue
4. **Use Headless Mode** - Faster, less resource-intensive

## Verification Checklist

- [ ] Python is 64-bit (`python --version` shows 3.8+)
- [ ] Chrome is installed (`chrome.exe` or `google-chrome` command works)
- [ ] ChromeDriver cache is cleared (`.wdm` folder deleted)
- [ ] All dependencies are installed (`pip list | grep -E "selenium|webdriver-manager|waitress"`)
- [ ] Ports 5000 not in use (`netstat -ano | findstr :5000`)
- [ ] No firewall blocking port 5000

## Still Need Help?

Check these resources:
- Selenium Documentation: https://www.selenium.dev/documentation/
- WebDriver Manager: https://github.com/SergeyPirogov/webdriver_manager
- Flask-SocketIO: https://flask-socketio.readthedocs.io/
- Waitress Documentation: https://docs.pylonsproject.org/projects/waitress/

## Contact

If the issue persists, provide the following information:

1. Output of `python fix_chromedriver.py`
2. Full error message from `app.log`
3. Output of `python -c "import platform; print(platform.platform())"`
4. Output of `python --version`
5. Chrome version (Chrome Menu > Settings > About Chrome)
