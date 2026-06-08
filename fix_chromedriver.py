#!/usr/bin/env python
"""
Utility script to clear ChromeDriver cache and fix Windows compatibility issues.
Run this if you get "WinError 193 - not a valid Win32 application" errors.
"""

import os
import sys
import shutil
import platform
import subprocess
from pathlib import Path

def find_correct_chromedriver():
    """Find the correct ChromeDriver executable in the cache."""
    wdm_path = os.path.expanduser("~/.wdm")
    
    if not os.path.exists(wdm_path):
        return None
    
    # Search for chromedriver executable
    if platform.system() == "Windows":
        pattern = "chromedriver.exe"
    else:
        pattern = "chromedriver"
    
    for root, dirs, files in os.walk(wdm_path):
        for file in files:
            if file == pattern:
                full_path = os.path.join(root, file)
                if os.path.isfile(full_path):
                    return full_path
    
    return None

def validate_chromedriver_integrity():
    """Validate that the cached ChromeDriver is not corrupted."""
    chromedriver_path = find_correct_chromedriver()
    
    if not chromedriver_path:
        print("⚠ ChromeDriver not found in cache")
        return False
    
    try:
        # Check file size (real chromedriver.exe is typically 50+ MB)
        file_size = os.path.getsize(chromedriver_path)
        min_size = 10 * 1024 * 1024  # 10 MB minimum
        
        if file_size < min_size:
            print(f"✗ ChromeDriver appears corrupted (size: {file_size} bytes, expected > {min_size})")
            return False
        
        print(f"✓ ChromeDriver found and healthy ({file_size / 1024 / 1024:.1f} MB)")
        return True
    except Exception as e:
        print(f"✗ Error validating ChromeDriver: {e}")
        return False

def clear_chromedriver_cache():
    """Clear the webdriver-manager cache."""
    cache_path = os.path.expanduser("~/.wdm")
    
    print(f"Checking ChromeDriver cache at: {cache_path}")
    
    if os.path.exists(cache_path):
        try:
            shutil.rmtree(cache_path)
            print(f"✓ Successfully cleared ChromeDriver cache")
        except Exception as e:
            print(f"✗ Error clearing cache: {e}")
            return False
    else:
        print(f"✓ Cache directory does not exist (already clean)")
    
    return True

def verify_chrome_installation():
    """Verify Chrome is installed on the system."""
    print("\nVerifying Chrome installation...")
    
    chrome_paths = {
        "Windows": [
            "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
            os.path.expanduser("~\\AppData\\Local\\Google\\Chrome\\Application\\chrome.exe"),
        ],
        "Darwin": [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        ],
        "Linux": [
            "/usr/bin/google-chrome",
            "/usr/bin/chromium",
            "/usr/bin/chromium-browser",
        ]
    }
    
    system = platform.system()
    paths = chrome_paths.get(system, [])
    
    for path in paths:
        if os.path.exists(path):
            # Get Chrome version
            try:
                if platform.system() == "Windows":
                    output = subprocess.check_output(
                        [path, "--version"],
                        stderr=subprocess.DEVNULL,
                        text=True
                    ).strip()
                    print(f"✓ Chrome found: {output}")
                else:
                    print(f"✓ Chrome found at: {path}")
            except Exception as e:
                print(f"✓ Chrome found at: {path}")
            return True
    
    print("✗ Chrome not found. Please install Google Chrome from https://www.google.com/chrome/")
    return False

def check_python_architecture():
    """Check Python architecture (32-bit vs 64-bit)."""
    print("\nChecking Python configuration...")
    
    arch = platform.architecture()
    bits = arch[0]
    system = platform.platform()
    
    print(f"  Python Architecture: {bits}")
    print(f"  System: {system}")
    print(f"  Python Version: {sys.version}")
    
    # Check if Python is 64-bit
    if "32bit" in bits:
        print("⚠ WARNING: You are running 32-bit Python.")
        print("  Recommendation: Install 64-bit Python for better compatibility")
        print("  Download from: https://www.python.org/downloads/")
    else:
        print("✓ Running 64-bit Python (recommended)")
    
    return True

def upgrade_dependencies():
    """Upgrade selenium and webdriver-manager."""
    print("\nUpgrading dependencies...")
    
    packages = ["selenium", "webdriver-manager"]
    
    for package in packages:
        try:
            print(f"  Upgrading {package}...")
            subprocess.check_call([
                sys.executable, "-m", "pip", "install", 
                "--upgrade", package, "-q"
            ])
            print(f"  ✓ {package} upgraded")
        except Exception as e:
            print(f"  ✗ Error upgrading {package}: {e}")
            return False
    
    return True

def main():
    """Main cleanup routine."""
    print("=" * 60)
    print("ChromeDriver Cache Cleaner and Windows Compatibility Fix")
    print("=" * 60)
    
    # Check architecture
    check_python_architecture()
    
    # Verify Chrome installation
    if not verify_chrome_installation():
        print("\n✗ Chrome is not installed. Please install it first.")
        return False
    
    # Check current cache integrity
    print("\nValidating current ChromeDriver cache...")
    if validate_chromedriver_integrity():
        print("\n✓ Your ChromeDriver cache appears healthy!")
        print("  If you're still getting errors, try clearing and reinstalling.")
        response = input("  Clear and reinstall ChromeDriver? (y/n): ").strip().lower()
        if response != 'y':
            print("  Skipping cache clear.")
            return True
    
    # Clear cache
    if not clear_chromedriver_cache():
        print("\n⚠ Could not clear all cache, but continuing...")
    
    # Upgrade dependencies
    print("\nUpgrading Selenium and WebDriver Manager...")
    if not upgrade_dependencies():
        print("\n⚠ Some dependencies failed to upgrade")
    
    print("\n" + "=" * 60)
    print("✓ Cleanup complete!")
    print("=" * 60)
    print("\nYou can now run your application:")
    print("  python -m waitress --port=5000 --host=0.0.0.0 wsgi:application")
    print("\nOr use:")
    print("  run_app.bat    (Windows)")
    print("  ./run_app.sh   (Linux/macOS)")
    
    return True

if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\nAborted by user.")
        sys.exit(1)
