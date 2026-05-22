# Outlook-Auto-Categorizer
A Python-based Outlook inbox automation tool that automatically categorizes and routes emails into organized folders using custom business logic.
Built using Python + Outlook COM automation (`pywin32`) for enterprise productivity workflows.
<img width="998" height="736" alt="emailscript" src="https://github.com/user-attachments/assets/3d7b3269-9c37-4020-a032-f1189226d19a" />
---
# Features
- Automatically categorizes Outlook emails
- Moves emails into matching folders
- Prioritizes emails using rule hierarchy
- Detects Microsoft Teams notifications
- Separates internal company communications
- Identifies case updates and Dayforce emails
- Flags emails from important colleagues
- Skips non-email Outlook items safely
- Processes only unorganized inbox emails
---
# Categories
| Category | Detection Logic |
|---|---|
| Department | Emails CCing `Department` |
| Company News | Emails from internal communications senders |
| Teams Messages | Emails containing `Microsoft Teams` in body |
| Cases | Subject contains `case status update` |
| Dayforce | Subject or sender contains `dayforce` |
| Important | Emails from or CCing key colleagues |
| Other | All remaining emails |
---
# Folder Structure
Create these folders directly under your Outlook Inbox:
```text
Inbox
│
├── Department
├── Company News
├── Teams Messages
├── Cases
├── Dayforce
├── Important
└── Other

⸻

Requirements
Windows
Microsoft Outlook Desktop App
Python 3.x
Python library:
pip install pywin32

⸻

Installation
Clone the repository:
git clone https://github.com/yourusername/outlook-auto-categorizer.git
Navigate into the project folder:
cd outlook-email-categorizer
Install dependencies:
pip install pywin32
Run the script:
python categorizeemails.py

⸻

How It Works
The script:
Connects to Outlook using COM automation
Reads emails from the Inbox
Applies prioritization rules
Assigns Outlook categories
Moves emails into matching folders
Only emails still located in the Inbox are processed.

⸻

Priority Order
The categorization logic uses priority ordering:
Department
Company News
Teams Messages
Cases
Dayforce
Important
Other
This ensures high-priority operational emails are categorized correctly first.

⸻

Technologies Used
Python
pywin32
Outlook COM Automation

⸻

Disclaimer
This project was created for educational and productivity purposes.
Use responsibly within your organization’s IT and security policies.
