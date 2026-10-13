# Extract Parts

A PowerShell utility for extracting and flattening RAR archive sets, optionally compressing loose `.nsp` / `.xci` files, and safely moving outputs into a chosen destination folder.

This project was created and iteratively improved with GitHub Copilot as a coding assistant. It is provided as-is for education, automation, and personal use, and it should be used only with content you own or are otherwise legally entitled to process.

## What it does

- scans a source folder for RAR archive sets
- extracts archive contents with WinRAR
- flattens extracted files into a destination folder
- optionally converts loose `.nsp` / `.xci` files to `.nsz` / `.xcz` using NSZ
- can prompt before deleting original archives or loose source files
- records diagnostic logs and a transcript for troubleshooting

## Key safety behavior

- prompts interactively by default
- honors command-line overrides when provided
- avoids destructive cleanup in dry-run or no-delete modes
- verifies output before deleting original source files
- keeps archives unless the user explicitly confirms deletion

## Requirements

- Windows PowerShell 5.1+
- WinRAR CLI (`Rar.exe`)
- NSZ CLI and a compatible legal keys file (`prod.keys` or `keys.txt`)
- Python environment for the NSZ install flow, if needed

## Important legal note

- You must only process files you own or are legally allowed to convert or extract.
- The script does not provide or download keys.
- You are responsible for obtaining any required legal game keys in compliance with local law and platform terms.

## Quick start

1. Open PowerShell in the repository.
2. Run:

```powershell
.\Extract-Parts.ps1
```

3. Pick the source folder and destination folder when prompted.
4. Confirm any cleanup choices that appear.

## Example command-line usage

```powershell
.\Extract-Parts.ps1 -SourceFolder "D:\Roms" -OutputFolder "D:\Extracted" -DeleteArchives -DeleteStandalone
.\Extract-Parts.ps1 -NonInteractive -DryRun
.\Extract-Parts.ps1 -NonInteractive -DeleteArchives
```

## Notes

- If command-line flags are supplied, they take precedence over the interactive prompts.
- If you run in non-interactive mode, the script prefers safe defaults.
- The script logs details to the local app data logs directory and keeps a transcript for troubleshooting.

## License

This repository does not currently include a project license file. If you intend to redistribute or reuse it publicly, consider adding an explicit license before publishing broadly.
