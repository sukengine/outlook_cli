# Outlook CLI Skill

Automate Microsoft Outlook (Classic) operations via COM automation. This skill provides comprehensive coverage of Outlook email and calendar operations through the Outlook COM API.

## Prerequisites

### Outlook Classic Installation

**Windows:**
1. Install Microsoft 365 Apps for Enterprise or Office 2021/2019
2. During installation, select "Outlook (Classic)" option
3. For Microsoft 365, you may need to use the Office Deployment Tool:

```xml
<!-- configuration.xml for Outlook Classic -->
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Access" />
      <ExcludeApp ID="Excel" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="OneNote" />
      <ExcludeApp ID="PowerPoint" />
      <ExcludeApp ID="Publisher" />
      <ExcludeApp ID="Teams" />
      <ExcludeApp ID="Word" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
```

```powershell
# Download and run Office Deployment Tool
setup.exe /configure configuration.xml
```

**Verify Installation:**
```powershell
# Check if Outlook is installed
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
```

### Python Environment Setup

```bash
# Create virtual environment
python -m venv .venv

# Activate (Windows PowerShell)
.\.venv\Scripts\Activate.ps1

# Install pywin32 for COM automation
pip install pywin32

# Install additional dependencies
pip install python-dateutil pytz
```

### PowerShell Setup (Alternative)

```powershell
# No additional setup needed - uses built-in COM support
# Verify COM access works:
$outlook = New-Object -ComObject Outlook.Application
$outlook.Version
```

---

## Quick Reference

| Operation | COM Object | Key Method |
|-----------|------------|------------|
| Create email | `MailItem` | `Application.CreateItem(0)` |
| Send email | `MailItem` | `MailItem.Send()` |
| Create appointment | `AppointmentItem` | `Application.CreateItem(1)` |
| Create meeting | `AppointmentItem` | `MeetingStatus = 1`, `Send()` |
| Create task | `TaskItem` | `Application.CreateItem(3)` |
| Create contact | `ContactItem` | `Application.CreateItem(2)` |
| Get folder | `Folder` | `Namespace.GetDefaultFolder()` |
| Add attachment | `Attachment` | `Attachments.Add(path)` |
| Add recipient | `Recipient` | `Recipients.Add(email)` |
| Create rule | `Rule` | `Rules.Create(name)` |

---

## A. OUTLOOK EMAIL OPERATIONS

### 1. Start Outlook and Access Inbox

**Use Case:** Launch Outlook programmatically and access the default Inbox folder.

**Python:**
```python
import win32com.client

# Create Outlook Application object
outlook = win32com.client.Dispatch("Outlook.Application")

# Get MAPI namespace
namespace = outlook.GetNamespace("MAPI")

# Access default Inbox folder
inbox = namespace.GetDefaultFolder(6)  # olFolderInbox = 6

# Display the folder
inbox.Display()

# Get folder info
print(f"Folder: {inbox.Name}")
print(f"Item count: {inbox.Items.Count}")
```

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)  # olFolderInbox
$inbox.Display()
Write-Host "Inbox contains $($inbox.Items.Count) items"
```

**Folder Constants:**
| Constant | Value | Description |
|----------|-------|-------------|
| `olFolderInbox` | 6 | Inbox |
| `olFolderSentMail` | 5 | Sent Items |
| `olFolderDrafts` | 16 | Drafts |
| `olFolderDeletedItems` | 3 | Deleted Items |
| `olFolderOutbox` | 4 | Outbox |
| `olFolderCalendar` | 9 | Calendar |
| `olFolderContacts` | 10 | Contacts |
| `olFolderTasks` | 13 | Tasks |

---

### 2. Create and Display New Email

**Use Case:** Generate a new email message and display it to the user.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")

# Create new mail item (olMailItem = 0)
mail = outlook.CreateItem(0)

# Set basic properties
mail.Subject = "Meeting Reminder"
mail.Body = "Please remember our meeting tomorrow at 10 AM."
mail.To = "recipient@example.com"

# Display the email
mail.Display()

# Or send directly
# mail.Send()
```

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)  # olMailItem

$mail.Subject = "Meeting Reminder"
$mail.Body = "Please remember our meeting tomorrow at 10 AM."
$mail.To = "recipient@example.com"

$mail.Display()
# $mail.Send()  # To send directly
```

---

### 3. Configure Email Accounts for Sending

**Use Case:** Enumerate available email accounts and send from a specific account.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Enumerate all accounts
for account in outlook.Session.Accounts:
    print(f"Account: {account.DisplayName}")
    print(f"  Type: {account.AccountType}")
    print(f"  SMTP: {account.SmtpAddress}")

# Send from specific account
mail = outlook.CreateItem(0)
mail.Subject = "Test from specific account"
mail.To = "recipient@example.com"

# Set the sending account
for account in outlook.Session.Accounts:
    if account.SmtpAddress == "specific@domain.com":
        mail.SendUsingAccount = account
        break

mail.Send()
```

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application

# List all accounts
$outlook.Session.Accounts | ForEach-Object {
    Write-Host "Account: $($_.DisplayName)"
    Write-Host "  SMTP: $($_.SmtpAddress)"
}

# Send from specific account
$mail = $outlook.CreateItem(0)
$mail.Subject = "Test from specific account"
$mail.To = "recipient@example.com"

$account = $outlook.Session.Accounts | Where-Object { $_.SmtpAddress -eq "specific@domain.com" }
if ($account) {
    $mail.SendUsingAccount = $account
}
$mail.Send()
```

---

### 4. Add and Configure Message Recipients

**Use Case:** Add recipients with specific types (To, CC, Bcc).

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

mail.Subject = "Multi-recipient Email"

# Method 1: Direct assignment
mail.To = "to1@example.com; to2@example.com"
mail.CC = "cc@example.com"
mail.BCC = "bcc@example.com"

# Method 2: Using Recipients collection (more control)
recipient_to = mail.Recipients.Add("to@example.com")
recipient_to.Type = 1  # olTo = 1

recipient_cc = mail.Recipients.Add("cc@example.com")
recipient_cc.Type = 2  # olCC = 2

recipient_bcc = mail.Recipients.Add("bcc@example.com")
recipient_bcc.Type = 3  # olBCC = 3

# Resolve recipients against address book
mail.Recipients.ResolveAll()

mail.Display()
```

**Recipient Type Constants:**
| Constant | Value | Description |
|----------|-------|-------------|
| `olTo` | 1 | To recipient |
| `olCC` | 2 | Carbon Copy |
| `olBCC` | 3 | Blind Carbon Copy |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)

$mail.Subject = "Multi-recipient Email"

# Add recipients with types
$to = $mail.Recipients.Add("to@example.com")
$to.Type = 1  # olTo

$cc = $mail.Recipients.Add("cc@example.com")
$cc.Type = 2  # olCC

$bcc = $mail.Recipients.Add("bcc@example.com")
$bcc.Type = 3  # olBCC

$mail.Recipients.ResolveAll()
$mail.Display()
```

---

### 5. Attach Files to Email

**Use Case:** Add file attachments with optional captions.

**Python:**
```python
import win32com.client
import os

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

mail.Subject = "Report Attached"
mail.To = "manager@example.com"
mail.Body = "Please find the attached report."

# Add single attachment
attachment_path = r"C:\Reports\monthly_report.xlsx"
mail.Attachments.Add(attachment_path)

# Add attachment with position (embedded in body)
# Position 0 = hide, 1 = first character position
mail.Attachments.Add(r"C:\Images\logo.png", 1)  # olByValue = 1

# Add multiple attachments
files_to_attach = [
    r"C:\Reports\report1.pdf",
    r"C:\Reports\report2.xlsx",
    r"C:\Data\data.csv"
]

for file_path in files_to_attach:
    if os.path.exists(file_path):
        mail.Attachments.Add(file_path)

mail.Display()
```

**Attachment Position Constants:**
| Constant | Value | Description |
|----------|-------|-------------|
| `olByValue` | 1 | Attachment is copied with message |
| `olByReference` | 4 | Attachment is a link |
| `olEmbeddedItem` | 5 | OLE embedded attachment |
| `olOLE` | 6 | OLE attachment |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)

$mail.Subject = "Report Attached"
$mail.To = "manager@example.com"
$mail.Body = "Please find the attached report."

# Add attachments
$mail.Attachments.Add("C:\Reports\monthly_report.xlsx")
$mail.Attachments.Add("C:\Images\logo.png")

$mail.Display()
```

---

### 6. Retrieve and Navigate Messages

**Use Case:** Access specific emails from folders.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # olFolderInbox

# Get all items
items = inbox.Items

# Sort by received date (descending)
items.Sort("[ReceivedTime]", True)

# Get first item
first_email = items.GetFirst()
print(f"Subject: {first_email.Subject}")
print(f"From: {first_email.SenderName}")
print(f"Received: {first_email.ReceivedTime}")

# Get specific item by index
second_email = items.Item(2)

# Iterate through items
count = 0
for item in items:
    if count >= 10:  # Limit to first 10
        break
    print(f"{count + 1}. {item.Subject}")
    count += 1

# Access item properties
print(f"Unread: {first_email.UnRead}")
print(f"Importance: {first_email.Importance}")  # 0=Low, 1=Normal, 2=High
```

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)

$items = $inbox.Items
$items.Sort("[ReceivedTime]", $true)

# Get first 5 emails
$items | Select-Object -First 5 | ForEach-Object {
    Write-Host "Subject: $($_.Subject)"
    Write-Host "From: $($_.SenderName)"
    Write-Host "Received: $($_.ReceivedTime)"
    Write-Host "---"
}
```

---

### 7. Search and Filter Inbox Items

**Use Case:** Find specific emails using filters.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)

# Method 1: Find method (returns first match)
items = inbox.Items
found_item = items.Find('[Subject] = "Need your advice"')
if found_item:
    print(f"Found: {found_item.Subject}")

# Method 2: Restrict method (returns filtered collection)
# Filter by subject containing text
filtered_items = items.Restrict('[Subject] LIKE "%urgent%"')

# Filter by unread
unread_items = items.Restrict('[UnRead] = True')

# Filter by sender
sender_items = items.Restrict('[SenderName] = "John Doe"')

# Filter by date range
from datetime import datetime
date_filter = f'[ReceivedTime] >= "{datetime(2024, 1, 1).strftime("%m/%d/%Y %H:%M %p")}"'
date_items = items.Restrict(date_filter)

# Filter by attachment
attachment_filter = items.Restrict('@SQL="urn:schemas:httpmail:hasattachment" = 1')

# Count results
print(f"Unread emails: {unread_items.Count}")

# Iterate filtered results
for item in filtered_items:
    print(f"Subject: {item.Subject}")
```

**Common Filter Properties:**
| Property | Example |
|----------|---------|
| `[Subject]` | `'[Subject] = "Test"'` |
| `[SenderName]` | `'[SenderName] = "John"'` |
| `[UnRead]` | `'[UnRead] = True'` |
| `[ReceivedTime]` | `'[ReceivedTime] >= "01/01/2024"'` |
| `[Importance]` | `'[Importance] = 2'` (High) |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)
$items = $inbox.Items

# Find unread emails
$unread = $items.Restrict('[UnRead] = True')
Write-Host "Unread count: $($unread.Count)"

# Find by subject
$found = $items.Find('[Subject] = "Important"')
if ($found) {
    Write-Host "Found: $($found.Subject)"
}
```

---

### 8. Process Offline Folders (.ost)

**Use Case:** Work with cached Exchange data.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Access stores
for store in namespace.Stores:
    print(f"Store: {store.DisplayName}")
    print(f"  FilePath: {store.FilePath}")
    print(f"  Is Cached: {store.IsCached}")

    # Get root folder
    try:
        root = store.GetRootFolder()
        print(f"  Root: {root.Name}")

        # Enumerate subfolders
        for folder in root.Folders:
            print(f"    - {folder.Name} ({folder.Items.Count} items)")
    except Exception as e:
        print(f"  Error accessing store: {e}")

# Work with items in reverse order
inbox = namespace.GetDefaultFolder(6)
items = inbox.Items
items.Sort("[ReceivedTime]", False)  # Ascending for reverse iteration

item = items.GetLast()
while item:
    print(f"Subject: {item.Subject}")
    item = items.GetPrevious()
```

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# List all stores
$namespace.Stores | ForEach-Object {
    Write-Host "Store: $($_.DisplayName)"
    Write-Host "  Cached: $($_.IsCached)"
}
```

---

### 9. Enumerate Mail Folders and Stores

**Use Case:** Navigate folder hierarchy.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

def enumerate_folders(folder, depth=0):
    """Recursively enumerate all folders"""
    indent = "  " * depth
    print(f"{indent}{folder.Name} ({folder.Items.Count} items)")

    try:
        for subfolder in folder.Folders:
            enumerate_folders(subfolder, depth + 1)
    except:
        pass

# Enumerate all stores and their folders
for store in namespace.Stores:
    print(f"\n=== Store: {store.DisplayName} ===")
    try:
        root = store.GetRootFolder()
        enumerate_folders(root)
    except Exception as e:
        print(f"  [Error: {e}]")

# Access specific folder by path
inbox = namespace.GetDefaultFolder(6)
subfolder = inbox.Folders["Archive"]
if subfolder:
    print(f"Archive folder: {subfolder.Items.Count} items")

# Create new folder
new_folder = inbox.Folders.Add("MyNewFolder")
```

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

function Get-Folders {
    param($Folder, $Depth = 0)
    $indent = "  " * $Depth
    Write-Host "$indent$($Folder.Name) ($($Folder.Items.Count) items)"
    $Folder.Folders | ForEach-Object {
        Get-Folders -Folder $_ -Depth ($Depth + 1)
    }
}

$inbox = $namespace.GetDefaultFolder(6)
Get-Folders -Folder $inbox
```

---

### 10. Automate Inbox Rules

**Use Case:** Create and manage email rules programmatically.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Get rules collection
rules = namespace.DefaultStore.GetRules()

# List existing rules
print("Existing Rules:")
for rule in rules:
    print(f"  - {rule.Name} (Enabled: {rule.Enabled})")

# Create a new rule
new_rule = rules.Create("Move Newsletter Emails", 0)  # 0 = olRuleReceive

# Add condition: from specific sender
from_condition = new_rule.Conditions.From
from_condition.Enabled = True
from_condition.Recipients.Add("newsletter@company.com")
from_condition.Recipients.ResolveAll()

# Add action: move to folder
move_action = new_rule.Actions.MoveToFolder
move_action.Enabled = True

# Get destination folder
inbox = namespace.GetDefaultFolder(6)
newsletters_folder = inbox.Folders.Add("Newsletters") if "Newsletters" not in [f.Name for f in inbox.Folders] else inbox.Folders["Newsletters"]
move_action.Folder = newsletters_folder

# Save rules
rules.Save()
print(f"Created rule: {new_rule.Name}")

# Enable/disable rule
new_rule.Enabled = True
rules.Save()

# Delete rule
# rules.Remove("Move Newsletter Emails")
# rules.Save()
```

**Rule Type Constants:**
| Constant | Value | Description |
|----------|-------|-------------|
| `olRuleReceive` | 0 | Apply to incoming messages |
| `olRuleSend` | 1 | Apply to outgoing messages |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$rules = $namespace.DefaultStore.GetRules()

# List rules
$rules | ForEach-Object {
    Write-Host "$($_.Name) - Enabled: $($_.Enabled)"
}

# Create new rule
$newRule = $rules.Create("Auto Reply", 0)
$newRule.Conditions.Subject.Enabled = $true
$newRule.Conditions.Subject.Text = @("URGENT")
$rules.Save()
```

---

### 11. Open the "Select Names" Dialog

**Use Case:** Let users select recipients from address book.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Create Select Names dialog
dialog = namespace.GetSelectNamesDialog()

# Configure dialog
dialog.Title = "Select Recipients"
dialog.AllowMultipleSelection = True
dialog.ShowOnlyInitialAddressList = False

# Set default address list
dialog.InitialAddressList = namespace.AddressLists("Global Address List")

# Show dialog
dialog.Display()

# Get selected recipients
if dialog.Recipients.Count > 0:
    print("Selected recipients:")
    for recipient in dialog.Recipients:
        print(f"  - {recipient.Name} ({recipient.Address})")

    # Use in email
    mail = outlook.CreateItem(0)
    mail.Recipients = dialog.Recipients
    mail.Display()
```

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

$dialog = $namespace.GetSelectNamesDialog()
$dialog.Title = "Select Recipients"
$dialog.AllowMultipleSelection = $true
$dialog.Display()

$dialog.Recipients | ForEach-Object {
    Write-Host "Selected: $($_.Name)"
}
```

---

### 12. Create New Contacts

**Use Case:** Add contacts programmatically.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Get contacts folder
contacts = namespace.GetDefaultFolder(10)  # olFolderContacts

# Create new contact
contact = outlook.CreateItem(2)  # olContactItem

# Set contact properties
contact.FirstName = "John"
contact.LastName = "Doe"
contact.FullName = "John Doe"
contact.Email1Address = "john.doe@example.com"
contact.CompanyName = "ACME Corp"
contact.JobTitle = "Software Engineer"
contact.BusinessTelephoneNumber = "+1-555-123-4567"
contact.MobileTelephoneNumber = "+1-555-987-6543"
contact.BusinessAddress = "123 Main St, City, State 12345"

# Save contact
contact.Save()
print(f"Created contact: {contact.FullName}")

# Find existing contact
contacts_items = contacts.Items
existing = contacts_items.Find('[FullName] = "John Doe"')
if existing:
    print(f"Found existing: {existing.FullName}")
```

**Contact Properties:**
| Property | Description |
|----------|-------------|
| `FirstName`, `LastName` | Name parts |
| `FullName` | Full display name |
| `Email1Address`, `Email2Address` | Email addresses |
| `CompanyName` | Organization |
| `JobTitle` | Position |
| `BusinessTelephoneNumber` | Work phone |
| `MobileTelephoneNumber` | Mobile phone |
| `BusinessAddress` | Work address |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application

$contact = $outlook.CreateItem(2)  # olContactItem
$contact.FirstName = "John"
$contact.LastName = "Doe"
$contact.Email1Address = "john.doe@example.com"
$contact.CompanyName = "ACME Corp"
$contact.Save()

Write-Host "Created contact: $($contact.FullName)"
```

---

### 13. Delegate Tasks via Email

**Use Case:** Create and assign tasks to others.

**Python:**
```python
import win32com.client
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch("Outlook.Application")

# Create new task
task = outlook.CreateItem(3)  # olTaskItem

# Set task properties
task.Subject = "Review quarterly report"
task.Body = "Please review and provide feedback on the attached quarterly report."
task.DueDate = datetime.now() + timedelta(days=7)
task.Importance = 2  # High
task.Status = 0  # olTaskNotStarted

# Assign task to someone
task.Assign()
task.Recipients.Add("colleague@example.com")
task.Recipients.ResolveAll()

# Send task request
task.Send()
print(f"Task assigned: {task.Subject}")

# Process received task request (on recipient side)
inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)
for item in inbox.Items:
    if item.Class == 52:  # olTaskRequestItem
        # Get associated task
        associated_task = item.GetAssociatedTask(True)
        if associated_task:
            print(f"Task request: {associated_task.Subject}")

            # Accept the task
            associated_task.Respond(0)  # olTaskAccept = 0
            # Or decline
            # associated_task.Respond(1)  # olTaskDecline = 1
```

**Task Status Constants:**
| Constant | Value | Description |
|----------|-------|-------------|
| `olTaskNotStarted` | 0 | Not started |
| `olTaskInProgress` | 1 | In progress |
| `olTaskComplete` | 2 | Completed |
| `olTaskWaiting` | 3 | Waiting on someone |
| `olTaskDeferred` | 4 | Deferred |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application

$task = $outlook.CreateItem(3)  # olTaskItem
$task.Subject = "Review quarterly report"
$task.DueDate = (Get-Date).AddDays(7)
$task.Assign()
$task.Recipients.Add("colleague@example.com")
$task.Send()

Write-Host "Task assigned: $($task.Subject)"
```

---

## B. OUTLOOK CALENDAR OPERATIONS

### 1. Create Single Appointment

**Use Case:** Create a one-time calendar appointment.

**Python:**
```python
import win32com.client
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch("Outlook.Application")

# Create new appointment
appointment = outlook.CreateItem(1)  # olAppointmentItem

# Set appointment properties
appointment.Subject = "Team Meeting"
appointment.Body = "Weekly team sync-up meeting"
appointment.Location = "Conference Room A"

# Set date/time
start_time = datetime.now().replace(hour=10, minute=0, second=0)
appointment.Start = start_time
appointment.End = start_time + timedelta(hours=1)
appointment.Duration = 60  # Duration in minutes

# Set additional properties
appointment.ReminderSet = True
appointment.ReminderMinutesBeforeStart = 15
appointment.BusyStatus = 2  # olBusy
appointment.Importance = 1  # Normal

# Save appointment
appointment.Save()
print(f"Created appointment: {appointment.Subject}")
print(f"  Start: {appointment.Start}")
print(f"  End: {appointment.End}")
```

**Appointment Properties:**
| Property | Description |
|----------|-------------|
| `Subject` | Appointment title |
| `Body` | Description/notes |
| `Location` | Meeting location |
| `Start` | Start datetime |
| `End` | End datetime |
| `Duration` | Length in minutes |
| `ReminderSet` | Enable reminder |
| `ReminderMinutesBeforeStart` | Reminder timing |
| `BusyStatus` | Free/Busy/Tentative/OOF |
| `Importance` | Low/Normal/High |
| `Sensitivity` | Normal/Private/Confidential |

**BusyStatus Constants:**
| Constant | Value | Description |
|----------|-------|-------------|
| `olFree` | 0 | Free |
| `olTentative` | 1 | Tentative |
| `olBusy` | 2 | Busy |
| `olOutOfOffice` | 3 | Out of Office |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application

$apt = $outlook.CreateItem(1)  # olAppointmentItem
$apt.Subject = "Team Meeting"
$apt.Location = "Conference Room A"
$apt.Start = (Get-Date).AddHours(1)
$apt.Duration = 60
$apt.ReminderSet = $true
$apt.ReminderMinutesBeforeStart = 15
$apt.Save()

Write-Host "Created: $($apt.Subject)"
```

---

### 2. Send Meeting Requests

**Use Case:** Create a meeting with attendees and send invitations.

**Python:**
```python
import win32com.client
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch("Outlook.Application")

# Create meeting request
meeting = outlook.CreateItem(1)  # olAppointmentItem

# Set meeting properties
meeting.Subject = "Project Kickoff Meeting"
meeting.Location = "Main Conference Room"
meeting.Body = "Please join us for the project kickoff meeting."

# Set date/time
start_time = datetime.now().replace(hour=14, minute=0, second=0) + timedelta(days=1)
meeting.Start = start_time
meeting.End = start_time + timedelta(hours=2)
meeting.Duration = 120

# Mark as meeting (not just appointment)
meeting.MeetingStatus = 1  # olMeeting

# Add required attendees
meeting.RequiredAttendees = "manager@example.com; lead@example.com"

# Add optional attendees
meeting.OptionalAttendees = "observer@example.com"

# Or use Recipients collection for more control
recipient1 = meeting.Recipients.Add("manager@example.com")
recipient1.Type = 1  # olRequired

recipient2 = meeting.Recipients.Add("observer@example.com")
recipient2.Type = 2  # olOptional

meeting.Recipients.ResolveAll()

# Set reminder
meeting.ReminderSet = True
meeting.ReminderMinutesBeforeStart = 30

# Send meeting request
meeting.Send()
print(f"Meeting sent: {meeting.Subject}")
```

**MeetingStatus Constants:**
| Constant | Value | Description |
|----------|-------|-------------|
| `olNonMeeting` | 0 | Appointment without attendees |
| `olMeeting` | 1 | Meeting with attendees |
| `olMeetingReceived` | 3 | Received meeting request |
| `olMeetingCanceled` | 5 | Canceled meeting |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application

$meeting = $outlook.CreateItem(1)
$meeting.Subject = "Project Kickoff"
$meeting.Location = "Conference Room"
$meeting.Start = (Get-Date).AddDays(1).AddHours(14)
$meeting.Duration = 120
$meeting.MeetingStatus = 1  # olMeeting
$meeting.RequiredAttendees = "manager@example.com"
$meeting.Send()

Write-Host "Meeting sent: $($meeting.Subject)"
```

---

### 3. Respond to Meeting Invitations

**Use Case:** Accept, decline, or tentatively accept meeting requests.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)  # olFolderInbox

# Find meeting requests
for item in inbox.Items:
    if item.Class == 53:  # olMeetingRequestItem (MeetingItem)
        print(f"Meeting request: {item.Subject}")
        print(f"  Organizer: {item.SenderName}")
        print(f"  When: {item.GetAssociatedAppointment(False).Start}")

        # Get associated appointment
        appointment = item.GetAssociatedAppointment(False)

        # Respond to meeting
        # Option 1: Accept
        accept_response = appointment.Respond(3, True)  # olMeetingAccepted, SendResponse
        accept_response.Send()

        # Option 2: Decline
        # decline_response = appointment.Respond(4, True)  # olMeetingDeclined

        # Option 3: Tentatively accept
        # tentative_response = appointment.Respond(2, True)  # olMeetingTentative

        break
```

**Meeting Response Constants:**
| Constant | Value | Description |
|----------|-------|-------------|
| `olMeetingTentative` | 2 | Tentatively accept |
| `olMeetingAccepted` | 3 | Accept |
| `olMeetingDeclined` | 4 | Decline |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $namespace.GetDefaultFolder(6)

$inbox.Items | Where-Object { $_.Class -eq 53 } | ForEach-Object {
    $apt = $_.GetAssociatedAppointment($false)
    Write-Host "Meeting: $($apt.Subject)"

    # Accept
    $response = $apt.Respond(3, $true)  # olMeetingAccepted
    $response.Send()
}
```

---

### 4. Create Recurring Appointments

**Use Case:** Set up repeating calendar events.

**Python:**
```python
import win32com.client
from datetime import datetime

outlook = win32com.client.Dispatch("Outlook.Application")

# Create recurring appointment
appointment = outlook.CreateItem(1)  # olAppointmentItem
appointment.Subject = "Weekly Team Standup"
appointment.Location = "Meeting Room B"
appointment.Start = datetime.now().replace(hour=9, minute=30)

# Get recurrence pattern
pattern = appointment.GetRecurrencePattern()

# Set recurrence type
pattern.RecurrenceType = 1  # olRecursWeekly

# Configure weekly recurrence
pattern.DayOfWeekMask = 62  # Mon-Fri: 2+4+8+16+32 = 62
pattern.Interval = 1  # Every week
pattern.Duration = 30  # 30 minutes

# Set end date or number of occurrences
# Option 1: End by date
pattern.PatternEndDate = datetime(2025, 12, 31)

# Option 2: End after N occurrences
# pattern.Occurrences = 52

# Save the recurring appointment
appointment.Save()

# Clear recurrence (convert to single appointment)
# appointment.ClearRecurrencePattern()
# appointment.Save()

print(f"Created recurring: {appointment.Subject}")
print(f"  IsRecurring: {appointment.IsRecurring}")
```

**RecurrenceType Constants:**
| Constant | Value | Description |
|----------|-------|-------------|
| `olRecursDaily` | 0 | Daily |
| `olRecursWeekly` | 1 | Weekly |
| `olRecursMonthly` | 2 | Monthly |
| `olRecursMonthNth` | 3 | Monthly (Nth day) |
| `olRecursYearly` | 5 | Yearly |
| `olRecursYearNth` | 6 | Yearly (Nth day) |

**DayOfWeekMask Constants:**
| Day | Value |
|-----|-------|
| Sunday | 1 |
| Monday | 2 |
| Tuesday | 4 |
| Wednesday | 8 |
| Thursday | 16 |
| Friday | 32 |
| Saturday | 64 |
| Weekdays (Mon-Fri) | 62 |
| All days | 127 |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application

$apt = $outlook.CreateItem(1)
$apt.Subject = "Weekly Team Standup"
$apt.Start = (Get-Date).AddDays(1)

$pattern = $apt.GetRecurrencePattern()
$pattern.RecurrenceType = 1  # olRecursWeekly
$pattern.DayOfWeekMask = 62  # Mon-Fri
$pattern.Duration = 30

$apt.Save()
Write-Host "Created recurring appointment"
```

---

### 5. Share and Export Calendars

**Use Case:** Export calendar to iCalendar format or share via email.

**Python:**
```python
import win32com.client
import os

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Get calendar folder
calendar = namespace.GetDefaultFolder(9)  # olFolderCalendar

# Create CalendarSharing object
calendar_sharing = calendar.GetCalendarExporter()

# Configure export settings
calendar_sharing.CalendarDetail = 2  # olFullDetails
calendar_sharing.IncludeAttachments = True
calendar_sharing.IncludePrivateDetails = False
calendar_sharing.RestrictToWorkingHoursOnly = False

# Set date range
from datetime import datetime, timedelta
calendar_sharing.StartDate = datetime.now()
calendar_sharing.EndDate = datetime.now() + timedelta(days=30)

# Export to iCalendar file
ics_path = r"C:\Exports\calendar_export.ics"
calendar_sharing.SaveAsICal(ics_path)
print(f"Calendar exported to: {ics_path}")

# Forward calendar via email
mail = calendar_sharing.ForwardAsICal(0)  # olCalendarMailFormatDailySchedule
mail.Subject = "My Calendar"
mail.To = "colleague@example.com"
mail.Display()
```

**CalendarDetail Constants:**
| Constant | Value | Description |
|----------|-------|-------------|
| `olFreeBusyOnly` | 0 | Free/Busy only |
| `olFreeBusyAndSubject` | 1 | Free/Busy + Subject |
| `olFullDetails` | 2 | Full details |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendar = $namespace.GetDefaultFolder(9)

$exporter = $calendar.GetCalendarExporter()
$exporter.CalendarDetail = 2  # olFullDetails
$exporter.StartDate = Get-Date
$exporter.EndDate = (Get-Date).AddDays(30)
$exporter.SaveAsICal("C:\Exports\calendar.ics")

Write-Host "Calendar exported"
```

---

### 6. Customize Calendar View

**Use Case:** Modify how calendar is displayed.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
calendar = namespace.GetDefaultFolder(9)

# Get current view
current_view = calendar.CurrentView
print(f"Current view: {current_view.Name}")
print(f"  Type: {current_view.ViewType}")

# Get all views
for view in calendar.Views:
    print(f"View: {view.Name} (Type: {view.ViewType})")

# Create custom calendar view
new_view = calendar.Views.Add("My Custom View", 5, 1)  # Name, Type (olCalendarView), SaveOption

# Configure calendar view
new_view.CalendarViewMode = 0  # olCalendarViewDay
new_view.DayWeekFont.Name = "Segoe UI"
new_view.DayWeekFont.Size = 10
new_view.DayWeekTimeFont.Name = "Segoe UI"
new_view.DayWeekTimeFont.Size = 8

# Apply and save
new_view.Apply()
new_view.Save()

# Switch to different view mode
# olCalendarViewDay = 0
# olCalendarViewWeek = 1
# olCalendarViewMonth = 2
# olCalendarViewMultiDay = 3
```

**CalendarViewMode Constants:**
| Constant | Value | Description |
|----------|-------|-------------|
| `olCalendarViewDay` | 0 | Day view |
| `olCalendarViewWeek` | 1 | Week view |
| `olCalendarViewMonth` | 2 | Month view |
| `olCalendarViewMultiDay` | 3 | Multi-day view |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendar = $namespace.GetDefaultFolder(9)

$views = $calendar.Views
$views | ForEach-Object {
    Write-Host "View: $($_.Name)"
}
```

---

### 7. Customize Timeline Views

**Use Case:** Configure timeline display for tasks/calendar.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
calendar = namespace.GetDefaultFolder(9)

# Find timeline view
for view in calendar.Views:
    if view.ViewType == 4:  # olTimelineView
        timeline = view
        break
else:
    # Create new timeline view
    timeline = calendar.Views.Add("My Timeline", 4, 1)

# Configure timeline
timeline.TimelineViewMode = 0  # olTimelineViewDay
timeline.MaxLabelWidth = 100
timeline.ShowWeekNumbers = True
timeline.GroupByFields = "Subject"

# Font settings
timeline.ItemFont.Name = "Calibri"
timeline.ItemFont.Size = 9
timeline.ScaleFont.Name = "Calibri"
timeline.ScaleFont.Size = 8

# Apply changes
timeline.Apply()
timeline.Save()
```

**TimelineViewMode Constants:**
| Constant | Value | Description |
|----------|-------|-------------|
| `olTimelineViewDay` | 0 | Day timeline |
| `olTimelineViewWeek` | 1 | Week timeline |
| `olTimelineViewMonth` | 2 | Month timeline |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$calendar = $namespace.GetDefaultFolder(9)

$timelineView = $calendar.Views | Where-Object { $_.ViewType -eq 4 }
if ($timelineView) {
    $timelineView.ShowWeekNumbers = $true
    $timelineView.Apply()
}
```

---

### 8. Manage Categories

**Use Case:** Create and use color categories for organization.

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# List all categories
print("Categories:")
for category in namespace.Categories:
    print(f"  - {category.Name} (Color: {category.Color}, Key: {category.ShortcutKey})")

# Create new category
# Note: Color values are predefined
# 0=Red, 1=Orange, 2=Yellow, 3=Green, 4=Blue, 5=Purple, etc.
new_category = namespace.Categories.Add("Urgent", 0)  # Red color

# Assign category to email
mail = outlook.CreateItem(0)
mail.Subject = "Urgent Task"
mail.Categories = "Urgent"
mail.Save()

# Assign multiple categories
mail.Categories = "Urgent; Red Category; Work"

# Remove category
# namespace.Categories.Remove("Urgent")

# Filter by category
inbox = namespace.GetDefaultFolder(6)
for item in inbox.Items:
    if "Urgent" in item.Categories:
        print(f"Urgent: {item.Subject}")
```

**Category Color Constants:**
| Constant | Value | Color |
|----------|-------|-------|
| `olCategoryColorRed` | 0 | Red |
| `olCategoryColorOrange` | 1 | Orange |
| `olCategoryColorYellow` | 2 | Yellow |
| `olCategoryColorGreen` | 3 | Green |
| `olCategoryColorBlue` | 4 | Blue |
| `olCategoryColorPurple` | 5 | Purple |

**PowerShell:**
```powershell
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# List categories
$namespace.Categories | ForEach-Object {
    Write-Host "Category: $($_.Name)"
}

# Create category
$namespace.Categories.Add("Urgent", 0)  # Red

# Assign to item
$mail = $outlook.CreateItem(0)
$mail.Subject = "Urgent Task"
$mail.Categories = "Urgent"
$mail.Save()
```

---

## C. ADDITIONAL OPERATIONS

### Working with Time Zones

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# List all time zones
for tz in namespace.TimeZones:
    print(f"{tz.Name} (Bias: {tz.Bias} minutes)")

# Get current time zone
current_tz = namespace.CurrentTimeZone
print(f"Current: {current_tz.Name}")

# Create appointment with specific time zone
appointment = outlook.CreateItem(1)
appointment.StartTimeZone = namespace.TimeZones.Item("Pacific Standard Time")
appointment.EndTimeZone = namespace.TimeZones.Item("Eastern Standard Time")
```

---

### Working with Search Folders

**Python:**
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Get search folders
search_folders = namespace.GetDefaultFolder(6).Folders

# Create search folder (via MAPI)
# Search folders are created differently - use AdvancedSearch
results = outlook.AdvancedSearch(
    "Inbox",
    'urn:schemas:httpmail:subject LIKE "%urgent%"',
    False,
    "UrgentSearch"
)
```

---

### Error Handling

**Python:**
```python
import win32com.client
import pythoncom

try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = "Test"
    mail.Send()
except pythoncom.com_error as e:
    print(f"COM Error: {e}")
    print(f"  HRESULT: {e.hresult}")
    print(f"  Message: {e.excepinfo[2] if e.excepinfo else 'Unknown'}")
except Exception as e:
    print(f"Error: {e}")
```

**PowerShell:**
```powershell
try {
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.Subject = "Test"
    $mail.Send()
}
catch {
    Write-Host "Error: $_"
}
finally {
    # Release COM objects
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
}
```

---

## Best Practices

### 1. Resource Management

**Always release COM objects:**
```python
import win32com.client
import pythoncom

def send_email():
    outlook = None
    mail = None
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = "Test"
        mail.To = "test@example.com"
        mail.Send()
    finally:
        if mail:
            del mail
        if outlook:
            del outlook
        # Force garbage collection
        import gc
        gc.collect()
```

### 2. Security Considerations

- Outlook may display security prompts for programmatic access
- Consider using Outlook Security Settings or Group Policy
- Use `Outlook.Application` with appropriate permissions

### 3. Performance Tips

- Use `Restrict()` instead of iterating all items
- Set `Items.Sort()` before processing large collections
- Use `GetFirst()`/`GetNext()` for sequential access
- Close inspectors/explorers when done

### 4. Threading

- COM objects are not thread-safe
- Create all objects on the same thread
- Use `pythoncom.CoInitialize()` before COM operations in threads

---

## Troubleshooting

### Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| `COM object not registered` | Outlook not installed | Install Outlook Classic |
| `Operation aborted` | Security prompt blocked | Configure security settings |
| `Object doesn't support property` | Wrong object type | Check `Class` property |
| `Cannot save item` | Invalid property value | Validate before setting |

### Debug Tips

```python
# Check object class
print(f"Class: {item.Class}")  # 43=MailItem, 53=MeetingItem

# List all properties
import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
for attr in dir(mail):
    if not attr.startswith('_'):
        print(attr)
```

---

## Related Resources

- [Microsoft Outlook VBA Reference](https://learn.microsoft.com/en-us/office/vba/api/overview/outlook)
- [Outlook Object Model](https://learn.microsoft.com/en-us/office/vba/api/outlook.application)
- [MAPI Constants](https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace)
