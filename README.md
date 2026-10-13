# Archive Extract Convert

A PowerShell utility for extracting and flattening RAR archive sets, optionally compressing `.nsp` / `.xci` files, and copying outputs into a chosen destination folder. The script remains named `Extract-Parts.ps1` for compatibility with existing usage.

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
- makes `-DryRun` preview detected inputs without installing dependencies or creating destination folders/files
- avoids deleting source files in `-DryRun` or `-NoDelete` mode
- verifies output before deleting original source files
- keeps archives unless the user explicitly confirms deletion

## Requirements

- Windows PowerShell 5.1+
- WinRAR CLI (`Rar.exe`) for archive extraction
- NSZ CLI and a compatible legal keys file (`prod.keys` or `keys.txt`) for compression
- Python 3.8+ for the optional NSZ installation flow, if needed

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
.\Extract-Parts.ps1 -SourceFolder "D:\Roms" -OutputFolder "D:\Extracted" -NonInteractive -DryRun
.\Extract-Parts.ps1 -SourceFolder "D:\Roms" -OutputFolder "D:\Extracted" -NonInteractive -DeleteArchives
```

## Notes

- If command-line flags are supplied, they take precedence over the interactive prompts.
- In non-interactive mode, provide `-SourceFolder` and optionally `-OutputFolder`; duplicates are skipped, and missing dependencies or keys are not installed/selected through prompts. `-Force` re-enables prompts.
- The script logs details to the local app data logs directory and keeps a transcript for troubleshooting.
- `-DryRun` scans the selected folders and reports the number of archive sets and loose files. It does not install dependencies or extract, convert, copy, or delete files. Diagnostic logs may still be written.

## License

This project is distributed under the MIT License. See [LICENSE](./LICENSE) for the full text.
