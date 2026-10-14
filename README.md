# Archive Extract Convert

A PowerShell utility for extracting and flattening RAR archive sets, optionally compressing `.nsp` / `.xci` files, and copying outputs into a chosen destination folder. The script remains named `Extract-Parts.ps1` for compatibility with existing usage.

This project was created and iteratively improved with GitHub Copilot as a coding assistant. It is provided as-is for education, automation, and personal use, and it should be used only with content you own or are otherwise legally entitled to process.

## What it does

- scans a source folder for RAR archive sets
- extracts archive contents with WinRAR
- adaptively increases extraction and compression concurrency when queued work and system resources allow, with no fixed job-count ceiling in resource-monitoring mode
- starts compression for newly finished archive extractions while other conversions are still running, when an NSZ slot is available
- flattens extracted files into a destination folder
- optionally converts loose `.nsp` / `.xci` files to `.nsz` / `.xcz` using NSZ
- displays a color-coded live dashboard with extraction/conversion states, current steps, elapsed times, queue previews, and results
- separates the archive-set completion bar from the current-item progress bar, so extraction, copy, and SHA-256 progress can advance before an archive set is fully handled
- shows WinRAR progress percentages as they are emitted (including carriage-return updates) and NSZ percentages when available, with an approximate per-item ETA after enough progress has been observed; the ETA is for the current archive item/conversion, not the whole run
- keeps dashboard display values in a shared state object; only the centralized one-second refresh gate renders it, in place for interactive terminals and as snapshots when output is redirected
- large file copies and SHA-256 checks update dashboard state with byte progress while the refresh gate continues to render at one-second intervals
- displays current CPU, available-memory, disk-activity, free-space, and adaptive concurrency status in the dashboard
- lets an interactive user press `Q` to stop starting work; active jobs finish and source files are preserved
- reports when a tool has not emitted a progress update for five minutes; this is informational and does not automatically terminate long-running work
- removes stale, script-owned extraction and compression work folders at startup while preserving folders used by active runs
- cleans up active background jobs, owned work folders, and transcript logging on normal completion and early exit
- asks before deleting original RAR archives and, when compression is enabled, loose `.nsp` / `.xci` source files
- records diagnostic logs and a transcript for troubleshooting

## Key safety behavior

- prompts interactively by default
- honors command-line overrides when provided
- makes `-DryRun` preview detected inputs without installing dependencies or creating destination folders/files
- avoids deleting source files in `-DryRun` or `-NoDelete` mode
- verifies placed output contents with SHA-256 once before deleting each original RAR archive set, while still validating each source volume; if a later same-name file overwrites an earlier result, deletion of the earlier source is blocked
- excludes the script's temporary ownership marker from extracted files
- skips same-name files only when SHA-256 hashes match; otherwise, the incoming file replaces the existing one
- in adaptive mode, starts with one extraction and one conversion slot, then increases capacity only when queued work exists and CPU, available memory, disk activity, and free space have headroom
- resource-monitoring mode has no fixed extraction or conversion job-count ceiling; it increases capacity gradually, and running jobs are never terminated when capacity is reduced
- falls back to one extraction and one conversion slot with an explicit dashboard warning if Windows resource metrics cannot be read; processing continues conservatively
- reserves at least 10 GB or 5% of each local work volume (whichever is larger) before scaling concurrency up
- pauses new extraction and conversion jobs below that disk-space reserve in both adaptive and fixed modes; interactive runs resume when space is available, while non-interactive runs stop safely with a nonzero exit code
- supports `-FixedConcurrency` to use fixed limits of six extractions and two conversions instead of adaptive scheduling
- asks whether to compress NSP/XCI files at the start of an interactive run; `-NoCompression` skips compression and copies extracted files uncompressed
- keeps full NSZ verification by default; `-QuickVerify` skips the whole-NSP SHA256 pass and verifies NCA hashes only, trading some verification coverage for less verification work
- exposes NSZ's compression level with `-CompressionLevel` (default 18; lower values favor speed over compression ratio and can require more output disk space)
- stages compressed output on the destination volume and moves verified results into place to avoid a second full-size temporary copy
- keeps archives unless the user explicitly confirms deletion
- in non-interactive mode, preserves source files unless `-DeleteArchives` or `-DeleteStandalone` explicitly enables deletion
- returns a nonzero process exit code when extraction, conversion, copy, or disk-safety operations fail

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
.\Extract-Parts.ps1 -SourceFolder "D:\Roms" -OutputFolder "D:\Extracted" -FixedConcurrency
.\Extract-Parts.ps1 -SourceFolder "D:\Roms" -OutputFolder "D:\Extracted" -NoCompression
.\Extract-Parts.ps1 -SourceFolder "D:\Roms" -OutputFolder "D:\Extracted" -QuickVerify
.\Extract-Parts.ps1 -SourceFolder "D:\Roms" -OutputFolder "D:\Extracted" -CompressionLevel 12
```

## Notes

- If command-line flags are supplied, they take precedence over the interactive prompts.
- Interactive runs ask whether to compress detected `.nsp` / `.xci` files before checking for NSZ or keys. Compression remains enabled by default in non-interactive runs unless `-NoCompression` is passed.
- RAR files must be extracted to a seekable `.nsp` / `.xci` file before NSZ can convert them; NSZ does not accept RAR input directly. XCI compression produces `.xcz`. `-QuickVerify` still validates NCA hashes but skips the full NSP SHA256 verification; omit it for the default full verification. NSZ defaults to compression level 18; lower `-CompressionLevel` values trade compression ratio for speed, so check destination free space when changing it.
- In non-interactive mode, provide `-SourceFolder` and optionally `-OutputFolder`; identical same-name files are skipped, different same-name files are overwritten, and missing dependencies or keys are not installed/selected through prompts. `-Force` re-enables prompts.
- The script logs details to the local app data logs directory and keeps a transcript for troubleshooting. Repetitive NSZ progress output is omitted, and each conversion retains at most 200 non-progress output lines.
- Startup cleanup only targets folders matching the script's private temporary-folder naming pattern. It automatically removes stale folders marked by this version, preserves work folders owned by a live process or referenced by an active process, and reports cleanup failures. Unmarked folders from older versions require interactive confirmation; non-interactive runs preserve them. `-DryRun` does not remove temporary folders.
- The disk reserve is checked on the output and temporary volumes every 15 seconds. If space is below the reserve, interactive runs wait for free space to recover; non-interactive runs stop rather than wait indefinitely. Already-running work is allowed to finish. Press `Q` during an interactive run to stop queueing new jobs; no-progress notices are advisory because large conversions can legitimately take a long time.
- `-DryRun` scans the selected folders and reports the number of archive sets and loose files. It does not install dependencies or extract, convert, copy, or delete files. Diagnostic logs may still be written.

## License

This project is distributed under the MIT License. See [LICENSE](./LICENSE) for the full text.
