# Mailbox Mover

A friendly Windows desktop app for moving Thunderbird mail into a local Classic Outlook PST. It is built with [Wails](https://wails.io/) (Go + TypeScript), Tailwind CSS 4, and daisyUI 5.

The repository still includes the original command-line scripts for reference, but the recommended workflow is now the desktop app.

## What the app does

1. Lets you choose Thunderbird's `Local Folders` directory with a native folder picker.
2. Converts Thunderbird mbox files to individual EML messages without modifying the source.
3. Detects open Outlook stores/PSTs and lets you choose the destination.
4. Recreates the folder hierarchy and imports messages, dates, senders, bodies, and attachments.
5. Shows live progress, skipped-message counts, safe cancellation, and an optional keep-awake mode.

## Requirements

### To use the built app

- Windows 10 or 11
- Classic Outlook desktop, **not New Outlook**
- An Outlook PST created/attached before migration
- Microsoft Edge WebView2 Runtime (already included with most current Windows installations)

### To develop

Wails v2 requires Go 1.21+, Node/npm, and WebView2 on Windows. Follow the official [Wails installation guide](https://wails.io/docs/gettingstarted/installation):

```powershell
go install github.com/wailsapp/wails/v2/cmd/wails@latest
wails doctor
```

If `wails` is not found, add `%USERPROFILE%\go\bin` to `PATH` and reopen the terminal.

## Run in development

```powershell
wails dev
```

Wails installs frontend dependencies automatically. You can also install them directly:

```powershell
npm install --prefix frontend
```

## Build the desktop app

```powershell
wails build
```

The executable is written to:

```text
build\bin\MailboxMover.exe
```

For a Windows installer, install [NSIS](https://wails.io/docs/guides/windows-installer/) and run:

```powershell
wails build -nsis
```

## Use Mailbox Mover

1. Close Thunderbird so its mailbox files do not change during conversion.
2. Open **Classic Outlook**.
3. In Outlook, create or attach the destination PST from **File → Account Settings → Data Files**.
4. Start Mailbox Mover.
5. Select Thunderbird's `Local Folders` directory. It is commonly under:

   ```text
   %APPDATA%\Thunderbird\Profiles\<profile>\Mail\Local Folders
   ```

6. Select the destination Outlook data file.
7. Choose an empty temporary folder with enough free disk space.
8. Review the preferences and start the migration.

The app never changes the original Thunderbird files. Stop requests take effect after the current message finishes; messages already imported remain in Outlook.

## Tests

```powershell
go test ./...
npm run build --prefix frontend
```

## Legacy scripts

- `convert_mbox_to_eml.py` — original configurable Python mbox converter
- `Import-ThunderbirdToOutlook.ps1` — original standalone Outlook importer
- `Keep-Awake.ps1` — original standalone keep-awake helper
- `scripts/Import-ThunderbirdToOutlook.ps1` — parameterized importer embedded into the desktop app

## Important limitations

- PST import is Windows-only because it uses Outlook COM automation.
- Classic Outlook must remain open throughout the import.
- Very large mailboxes should be migrated in logical batches if Outlook becomes unstable.
- The Outlook importer depends on Windows CDO/ADODB COM components available in typical Classic Outlook environments.

## License

See [LICENSE](LICENSE).
