# Thunderbird to Outlook Converter

Free, open-source tool to migrate Thunderbird emails into Outlook (classic) via local PST storage.

Converts Thunderbird's mbox format → EML files → injects directly into Outlook using its own COM automation engine.

**Requirements:** Windows, Classic Outlook (not New Outlook), PowerShell, Python 3 (for mbox conversion step).

---

## Why This Exists

Every Thunderbird-to-Outlook converter I found was either paid or broken. This script uses Outlook's own COM interface to inject emails directly into a local PST folder — no third-party import wizards, no Exchange requirements, no paid tools.

---

## How It Works

### The Outlook Engine Trick

Outlook's COM automation model lets you programmatically create items in PST folders. The trick: create a `PostItem` (Type 6) in the target folder, then force its `MessageClass` to `"IPM.Note"`. This makes Outlook treat it as a standard email — headers, body, attachments, dates, senders — everything maps correctly.

This bypasses Outlook's restrictive import dialogs and gives full control over metadata mapping via MAPI property tags.

### Pipeline

```
Thunderbird mbox files
    ↓ (Python script)
Individual .eml files
    ↓ (PowerShell script)
Outlook PST folders via COM automation
```

---

## Step-by-Step Setup

### 1. Export from Thunderbird

Thunderbird stores emails in mbox format (no file extension, or `.msf` index files alongside). You need to extract these to individual `.eml` files.

**On Linux** (where Thunderbird data lives):

```bash
python3 convert_mbox_to_eml.py
```

Edit the `src` and `dst` paths in the script first:
- `src` = path to Thunderbird `Local Folders` directory
- `dst` = where you want the EML files output

The script walks all mbox files, converts each message to individual `.eml` files, preserving folder structure.

**On Windows** (if Thunderbird is installed there):

Same script works — just use Windows paths:
```python
src = 'C:\\Users\\YOURNAME\\AppData\\Roaming\\Thunderbird\\Profiles\\YOURPROFILE.default\\Mail\\Local Folders'
dst = 'C:\\Users\\Public\\Thunderbird_Staging'
```

### 2. Copy EML Files to Windows Staging Path

Move the converted EML folder structure to:

```
C:\Users\Public\Thunderbird_Staging
```

This is the default source path in the import script. Folder structure should look like:

```
Thunderbird_Staging\
├── Inbox\
│   ├── 0.eml
│   ├── 1.eml
│   └── ...
├── Sent\
│   ├── 0.eml
│   └── ...
└── CustomFolder\
    ├── 0.eml
    └── SubFolder\
        └── 0.eml
```

### 3. Create Target PST in Outlook

Before running the script, you need a local PST file set up in Outlook:

1. Open **Classic Outlook** (not New Outlook)
2. Go to **File → Info → Tools → Clean Up Old Items** or **File → Account Settings → Account Settings → Data Files**
3. Click **Add** to create a new Outlook Data File (.pst)
4. Save it somewhere permanent (e.g., `Documents\Thunderbird_Archive.pst`)
5. When prompted, name the root folder **`Thunderbird_Archive`**

> **IMPORTANT:** The root folder name must exactly match `$TargetPSTName` in the script. Default is `"Thunderbird_Archive"`.

6. Verify it appears in your Outlook sidebar before proceeding.

### 4. Configure the Import Script

Open `Import-ThunderbirdToOutlook.ps1` and check these values:

| Variable | Default | What to set |
|----------|---------|-------------|
| `$SourcePath` | `C:\Users\Public\Thunderbird_Staging` | Path to your EML files |
| `$TargetPSTName` | `Thunderbird_Archive` | Root folder name in Outlook sidebar |

### 5. Run the Import

**In PowerShell** (as your normal user, no admin needed):

```powershell
# Allow script execution for this session only
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# Run the import
.\Import-ThunderbirdToOutlook.ps1
```

> **CRITICAL:** Classic Outlook MUST be open and running while the script executes. The script uses Outlook's COM interface — if Outlook isn't running, it can't connect.

### 6. Keep Your PC Awake

Large imports take time. If your PC goes to sleep, the script will crash.

Open a **second PowerShell window** and run:

```powershell
.\Keep-Awake.ps1
```

This sends synthetic `F15` keypresses every 45 seconds to prevent sleep. Close it when import finishes.

---

## Handling Large Mailboxes

**If you have many thousands of emails**, Outlook can crash or hang if you try to import everything at once.

### Split Strategy

1. Don't dump all EML files into one staging folder
2. Instead, create batch subfolders:
   ```
   Thunderbird_Staging\
   ├── Batch1_Inbox_2020\
   ├── Batch1_Inbox_2021\
   ├── Batch2_Sent\
   └── Batch3_OldArchives\
   ```
3. Run the script once per batch
4. After each batch completes, verify in Outlook, then move to next

### Signs You Need to Split

- Outlook becomes unresponsive during import
- Script hangs on a specific folder
- Memory usage climbs steadily without stopping
- You see "Outlook is not responding" dialogs

### Recommended Batch Sizes

- Under 5,000 emails per run: usually fine
- 5,000–20,000: split into batches of ~5,000
- Over 20,000: definitely split, consider running overnight in smaller chunks

---

## Customization

### Change Target Folder Name

Edit `$TargetPSTName` in the script:

```powershell
$TargetPSTName = "My_Archive"
```

Must match the root folder name in your Outlook sidebar exactly.

### Change Source Path

Edit `$SourcePath`:

```powershell
$SourcePath = "D:\MyEmails\EML"
```

### Skip Read Status

Remove or comment out this block to keep emails as unread:

```powershell
try {
    $MailItem.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003", 1)
} catch {}
```

---

## Troubleshooting

| Error | Cause | Fix |
|-------|-------|-----|
| `Could not find a root folder named 'Thunderbird_Archive'` | PST not created or wrong name | Create PST in Outlook, verify name matches `$TargetPSTName` |
| `New-Object -ComObject Outlook.Application` fails | Classic Outlook not open | Open Classic Outlook, keep it running |
| Script hangs / Outlook not responding | Too many emails at once | Split into smaller batches |
| PC went to sleep, import stopped | Sleep interrupted COM connection | Use `Keep-Awake.ps1` in separate terminal |
| Attachments missing | Temp path issues | Check `$SandboxDir` is writable, not locked |
| Emails show as Post, not Email | `MessageClass` not set | Ensure `$MailItem.MessageClass = "IPM.Note"` line runs |

---

## Technical Deep Dive

### COM Automation Flow

```
1. Create ADODB.Stream → load .eml file bytes
2. Create CDO.Message → parse EML via DataSource.OpenObject
3. Create PostItem (Type 6) in target PST folder
4. Set MessageClass = "IPM.Note" → identity swap to email
5. Map fields: Subject, Body/HTMLBody, Date, From
6. Map attachments via CDO → SaveToFile → Outlook.Attachments.Add
7. Set read status via MAPI prop 0x0E070003
8. Save to PST local storage
9. Release COM objects to prevent memory leaks
```

### MAPI Property Tags Used

| Tag | Purpose |
|-----|---------|
| `0x0E060040` | PR_CLIENT_SUBMIT_TIME (sent date) |
| `0x0C1A001E` | PR_SENDER_NAME |
| `0x0C1F001E` | PR_SENDER_EMAIL |
| `0x0E070003` | PR_MSG_FLAGS (read/unread) |

### Why PostItem + MessageClass Swap?

Outlook's COM model restricts direct email creation in PST folders. `MailItem.Add()` requires a store with transport capabilities. `PostItem.Add()` (Type 6) works in any folder. Setting `MessageClass = "IPM.Note"` tells Outlook to render it as a standard email — headers, threading, and all display properties work correctly.

---

## File Structure

```
EML_to_Outlook\
├── README.md                          # This file
├── Import-ThunderbirdToOutlook.ps1    # Main import script
├── convert_mbox_to_eml.py            # Python mbox → EML converter
├── Keep-Awake.ps1                     # Prevents PC sleep during import
└── .gitignore
```

---

## License

Do whatever you want with this. If it saves you time, great.
