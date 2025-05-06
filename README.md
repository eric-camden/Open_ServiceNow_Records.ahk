# Open_ServiceNow_Records.ahk
ServiceNow Quick Access Tool This AutoHotkey script streamlines access to ServiceNow records.
---

# **ServiceNow Quick Access Tool**
This **AutoHotkey (AHK) script** streamlines access to **ServiceNow records** by allowing users to quickly **search tickets**, **lookup users/groups**, and **customize their configurations**. It automatically applies **your company name** to ServiceNow URLs, ensures correct **ticket number formatting**, and lets you **save preferences** like ticket length and startup behavior.

---

## **Features**
✅ **Quick Ticket Search** – Instantly open ServiceNow tickets using predefined formats (`INC#######`, `CHG########`).  
✅ **User & Group Lookup** – Search for users, groups, and configuration items (CIs) within ServiceNow.  
✅ **Configurable Settings** – Set **your company name**, ticket number length, and startup preferences.  
✅ **Persistent Config Storage** – Automatically **stores preferences** in an **INI file** for future use.  
✅ **Tray Menu Reset** – Clear stored settings anytime by right-clicking the system tray icon.  
✅ **Customizable UI** – Modify ticket formats (7-10 digits) and control startup behavior.

---

## **Installation**
1. Install **AutoHotkey** (if not already installed) – [Download Here](https://www.autohotkey.com/).  
2. Clone or download this repository to your local machine.  
3. Run the script (`Open ServiceNow Records.ahk`).  
4. Configure your settings in the GUI (if prompted).  
5. Start searching tickets effortlessly!

---

## **Usage**
- **Hotkey:** `CTRL + SHIFT + O` → Opens the **ServiceNow ticket menu**.  
- **GUI Options:**  
  - Enter **company name** (applies to ServiceNow URLs).  
  - Select **ticket number format** (7-10 digits).  
  - Choose whether to **disable the startup popup** (can reset via tray menu).  
- **Tray Menu Options:**  
  - **Reset Settings** – Clears stored configurations.  
  - **Exit** – Closes the script.

---

## **Configuration Settings**
All settings are stored in `CompanySettings.ini` (located in the script directory). The file contains:
```ini
[Settings]
CompanyName=mycompany
TicketLength=7
NeverAskAgain=0
```
Modify these values manually or via the **GUI** for seamless customization.

---

## **Troubleshooting**
- **GUI not appearing?** Ensure `NeverAskAgain=0` in `CompanySettings.ini`.  
- **Incorrect URL formatting?** Double-check the **CompanyName** field.  
- **Ticket number mismatch?** Adjust **TicketLength** in the configuration.

---

## **License**
This script is licensed under the **MIT License**, allowing free usage, modification, and distribution with attribution.

---


