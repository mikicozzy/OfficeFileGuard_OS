# CrocoOfficeFileGuard

CrocoOfficeFileGuard is a small toolset that helps reduce accidental external disclosure of Office documents by marking files as "reserved" and requiring explicit user confirmation before sending such files to external recipients from Outlook.

This repository contains:
- `CrocoOfficeFileReserveToggle.Cli` — a CLI utility that toggles a custom Open XML document property named `reserved` on Office files (.docx, .xlsx, .pptx).
- `CrocoShellHandler` — a SharpShell-based Windows Explorer context-menu extension that launches the CLI for the selected file(s).
- `OutlookFileGuard_addin` — a VSTO Outlook add-in that intercepts outgoing email sends and prompts the user if reserved attachments are being sent to recipients outside the sender's SMTP domain.
- Documentation and notes in `ARCHITECTURE.md`.

Overview
- Toggle from Explorer: Right-click a supported Office file and invoke the context-menu command `Toggle Reserved`. This action sets or removes a CustomDocumentProperty named `reserved` in the file package (Open XML).
- Outlook guard: The VSTO add-in hooks `Application.ItemSend`. If the add-in detects that an outgoing email addressed to external recipients contains one or more files marked `reserved`, it prompts the sender to confirm the send. If the sender declines, the send is cancelled.

How it works (technical)
- CLI: Uses the Open XML SDK to open the package and add/remove a `CustomDocumentProperty` with Name=`reserved` and a boolean value. See `src/CrocoOfficeFileReserveToggle.Cli/Program.cs`.
- Shell extension: Implemented as a SharpShell `SharpContextMenu` (see `src/CrocoShellHandler/CrocoOfficeCommand.cs`). The menu item launches the CLI executable with the selected file path as argument.
- Outlook add-in: Implemented with VSTO. On `ItemSend`, the add-in:
  - Resolves the sender SMTP address (handles Exchange/MAPI fallback).
  - Scans recipients and treats any SMTP domain not matching the sender domain as external.
  - Inspects attachments and (when implemented) checks each attachment for the `reserved` property.
  - Prompts the user to explicitly confirm sending if reserved attachments are present and external recipients exist. See `src/OutlookFileGuard_addin/ThisAddIn.cs`.

Installation and setup (developer / tester)
1. Open the solution in Visual Studio 2026.
2. Build the solution: use __Build__ > __Build Solution__.
3. CLI deployment:
   - Place the CLI executable (`CrocoOfficeFileReserveToggle.Cli.exe`) in a stable location.
   - Update the shell extension or installer to point to the CLI path (the current SharpShell code launches a hard-coded path; see `CrocoShellHandler`).
4. Register the SharpShell extension:
   - Follow SharpShell documentation for registering the COM server or use the SharpShell Server for development.
   - Note: registering COM servers and making shell extensions available system-wide typically requires administrator privileges.
5. Outlook add-in:
   - Deploy or run the VSTO add-in from Visual Studio (or package it as required).
   - Trust and enable the add-in in Outlook if prompted.

Usage
- Toggle reserved state: In Windows Explorer, right-click a `.docx`, `.xlsx`, or `.pptx` file and select `Toggle Reserved`. The CLI will add the `reserved` custom property if missing, or remove it if already present.
- Send email: In Outlook, attach files as usual. If one or more attachments are marked `reserved` and any recipient domain differs from the sender's domain, the add-in will prompt for confirmation before sending.

Current limitations and **TODO**
- Attachment inspection is currently a placeholder in the add-in (`AreAttachementPrivate` in `ThisAddIn.cs`). Implement logic to:
  - Create a temporary copy of each attachment (if required).
  - Open the attachment as an Open XML package and check custom properties for `reserved`.
- The shell extension currently launches the CLI by a hard-coded executable path. Replace with a configurable path or embed the toggle logic into the extension to avoid brittle deployment.
- UX improvements:
  - Show the current reserved state in the context menu (e.g., menu text or a checked state).
  - Support multi-file toggling and progress/confirmation dialogs.
  - Replace debug MessageBox calls in the add-in with non-blocking logging and production-ready dialogs.
- SharpShell / Explorer notes:
  - Windows imposes limits on which items can appear in the primary context menu; experiments and limitations are documented in `src/CrocoShellHandler/Read.me`.

Security and operational notes
- This tool is a user-facing safety net and not a replacement for enterprise DLP or classification tools. Use it as an additional layer to reduce accidental leaks.
- COM registration, shell integration, and VSTO deployment can require admin privileges and may be constrained by corporate policy and Outlook security settings.
- Take care when temporarily storing attachments for inspection; remove temp copies securely after inspection.

Development pointers
- See `src/CrocoOfficeFileReserveToggle.Cli/Program.cs` for the Open XML implementation that creates and removes the `reserved` CustomDocumentProperty.
- Implement `AreAttachementPrivate` in `src/OutlookFileGuard_addin/ThisAddIn.cs` to call the same Open XML inspection logic used by the CLI.
- Replace MessageBox-based debug traces and finalize the confirmation UI (a typed-confirmation or stronger acknowledgement may be appropriate for higher assurance workflows).
- Follow repository coding standards and contribution guidelines when submitting changes.

License
- **TODO** to add a license file (e.g., `LICENSE` with MIT or other license) to the repository root before publishing.