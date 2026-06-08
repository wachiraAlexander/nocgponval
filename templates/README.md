# GPON Assistant Web App

## Quick Setup (No Admin Privileges Needed)
1. Install Python 3.10+ from python.org.
2. Create project folder, add files from above.
3. Run `pip install -r requirements.txt` (in project folder).
4. Place your `OLT_DATA1.xlsx` in the project root.
5. Run `python app.py`.
6. Open browser: http://localhost:5000 (on same PC) or http://YOUR_SERVER_IP:5000 (LAN).

## LAN Access
- Server runs on all interfaces (0.0.0.0:5000) – accessible from any LAN device.
- No firewall changes needed if port 5000 is open (default on most routers).
- Multiple users: Each upload gets a unique session.

## Features
- **LOS/Low Rx**: Upload Excel, SSH to OLTs, filter, auto-close optimal tickets in CRM.
- **Double Tickets**: Filter duplicates, dry-run option, mark as Duplicate in CRM.
- Real-time logs/progress via WebSockets.
- Downloads processed files automatically logged (check console or extend UI).

## Notes
- Chrome must be installed on server PC for Selenium.
- SSH/CRM use provided creds; secure in production.
- Logs in `app.log`.
- Extend `/download/<filename>` for file downloads if needed.

For issues: Check `app.log` or console.