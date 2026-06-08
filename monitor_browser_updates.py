"""
Browser Monitor - Attaches to existing Chrome session and monitors field updates
"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

print("="*80)
print("BROWSER MONITOR - Monitoring CRM Updates")
print("="*80)
print("\nThis script will connect to the Chrome browser and monitor field updates.")
print("Make sure the app is running and processing tickets.\n")

# Setup Chrome driver with debugging port
options = Options()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--start-maximized')

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)

try:
    # Navigate to CRM to get session
    print("Opening CRM...")
    driver.get("https://intranet.jtl.co.ke/app/issue")
    time.sleep(3)
    
    print("\n" + "="*80)
    print("MONITORING MODE ACTIVE")
    print("="*80)
    print("\nInstructions:")
    print("1. Keep this window open")
    print("2. Run your main app in another window")
    print("3. This script will check field values every 2 seconds")
    print("4. Press Ctrl+C to stop monitoring\n")
    
    last_ticket = None
    last_status = None
    last_subject = None
    last_category4 = None
    
    while True:
        try:
            current_url = driver.current_url
            
            # Check if we're on a ticket page
            if "/app/issue/ISS-" in current_url:
                ticket_id = current_url.split("/app/issue/")[1].split("?")[0]
                
                # New ticket detected
                if ticket_id != last_ticket:
                    print(f"\n{'='*80}")
                    print(f"📋 NEW TICKET DETECTED: {ticket_id}")
                    print(f"{'='*80}")
                    last_ticket = ticket_id
                    last_status = None
                    last_subject = None
                    last_category4 = None
                    time.sleep(1)  # Wait for page to load
                
                # Check Status field
                try:
                    status_element = driver.find_element(By.CSS_SELECTOR, "select[data-fieldname='status']")
                    current_status = status_element.get_attribute("value")
                    if current_status != last_status:
                        print(f"\n🔄 STATUS CHANGED: '{last_status}' → '{current_status}'")
                        last_status = current_status
                except:
                    pass
                
                # Check Subject field
                try:
                    subject_element = driver.find_element(By.CSS_SELECTOR, "input[data-fieldname='subject']")
                    current_subject = subject_element.get_attribute("value")
                    if current_subject != last_subject:
                        print(f"\n📝 SUBJECT CHANGED:")
                        print(f"   Old: {last_subject[:80] if last_subject else 'None'}...")
                        print(f"   New: {current_subject[:80]}...")
                        last_subject = current_subject
                except:
                    pass
                
                # Check Category 4 field
                try:
                    cat4_element = driver.find_element(By.CSS_SELECTOR, "input[data-fieldname='custom_category_4']")
                    current_cat4 = cat4_element.get_attribute("value")
                    if current_cat4 != last_category4:
                        print(f"\n📂 CATEGORY 4 CHANGED: '{last_category4}' → '{current_cat4}'")
                        last_category4 = current_cat4
                except:
                    pass
                
                # Check for save button clicks
                try:
                    # Look for any success messages or indicators
                    success_indicators = driver.find_elements(By.CSS_SELECTOR, ".indicator.green, .msgprint")
                    if success_indicators:
                        for indicator in success_indicators:
                            if indicator.is_displayed():
                                text = indicator.text
                                if text and "Saved" in text:
                                    print(f"\n💾 SAVE DETECTED: {text}")
                except:
                    pass
            
            else:
                if last_ticket:
                    print(f"\n⬅️  Left ticket {last_ticket}")
                    last_ticket = None
                    last_status = None
                    last_subject = None
                    last_category4 = None
            
            time.sleep(2)  # Check every 2 seconds
            
        except KeyboardInterrupt:
            print("\n\n⏹️  Monitoring stopped by user")
            break
        except Exception as e:
            # Silently continue on errors
            time.sleep(2)
    
except Exception as e:
    print(f"\n❌ Error: {e}")
    import traceback
    traceback.print_exc()
finally:
    print("\nClosing monitor...")
    driver.quit()

