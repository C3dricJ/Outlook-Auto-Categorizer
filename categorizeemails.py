import win32com.client
# -----------------------------------
# CONNECT TO OUTLOOK
# -----------------------------------
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# Inbox
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
# Sort newest first
messages.Sort("[ReceivedTime]", True)
# -----------------------------------
# OUTLOOK FOLDERS
# (Must exist under Inbox)
# -----------------------------------
folders = {
   "Department": inbox.Folders["Department"],
   "Company News": inbox.Folders["Company News"],
   "Teams Messages": inbox.Folders["Teams Messages"],
   "Cases": inbox.Folders["Cases"],
   "Dayforce": inbox.Folders["Dayforce"],
   "Important": inbox.Folders["Important"],
   "Other": inbox.Folders["Other"]
}
# -----------------------------------
# IMPORTANT COLLEAGUES
# -----------------------------------
IMPORTANT_PEOPLE = [
   "john doe",
   "jane doe"
]
# -----------------------------------
# FUNCTION TO CATEGORIZE + MOVE EMAIL
# -----------------------------------
def categorize_and_move(message, category_name):
   try:
# Assign Outlook category
       message.Categories = category_name
       message.Save()
# Move to folder
       message.Move(folders[category_name])
       print(f"[{category_name}] Moved -> {message.Subject}")
   except Exception as e:
       print(f"Error processing email: {e}")
# -----------------------------------
# MAIN EMAIL LOOP
# -----------------------------------
for message in messages:
   try:
# Only process actual mail items
       if message.Class != 43:
           continue
# Only process emails still in Inbox
       if message.Parent.Name != "Inbox":
           continue
# Safely get Outlook fields
       subject = str(getattr(message, "Subject", "")).lower()
       sender = str(getattr(message, "SenderName", "")).lower()
       cc = str(getattr(message, "CC", "")).lower()
       body = str(getattr(message, "Body", "")).lower()
# Combined searchable fields
       combined_text = f"{sender} {cc}"
# -----------------------------------
# Department (HIGHEST PRIORITY)
# -----------------------------------
       if "department" in cc:
           categorize_and_move(message, "Department")
# -----------------------------------
# Company NEWS
# -----------------------------------
       elif (
           "Company internal communications" in sender
           or "Company inside news" in sender
       ):
           categorize_and_move(message, "Company News")
# -----------------------------------
# MICROSOFT TEAMS
# -----------------------------------
       elif "microsoft teams" in body:
           categorize_and_move(message, "Teams Messages")
# -----------------------------------
# CASE STATUS UPDATES
# -----------------------------------
       elif "case status update" in subject:
           categorize_and_move(message, "Cases")
# -----------------------------------
# DAYFORCE
# -----------------------------------
       elif (
           "dayforce" in subject
           or "dayforce" in sender
       ):
           categorize_and_move(message, "Dayforce")
# -----------------------------------
# IMPORTANT PEOPLE
# LOWER PRIORITY
# -----------------------------------
       elif any(person in combined_text for person in IMPORTANT_PEOPLE):
           categorize_and_move(message, "Important")
# -----------------------------------
# EVERYTHING ELSE
# -----------------------------------
       else:
           categorize_and_move(message, "Other")
   except Exception as e:
       print(f"Skipping message due to error: {e}")