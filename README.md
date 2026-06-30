# OfficeFileGuard

OfficeFileGuard is an open-source, privacy-first Windows toolset designed to reduce the risk of accidental disclosure of Microsoft Office documents.

Instead of relying solely on user awareness, OfficeFileGuard allows Office documents to be explicitly marked as **reserved**. Whenever a marked document is attached to an Outlook email addressed to recipients outside the sender's organization, the Outlook add-in requires explicit confirmation before the message is sent.

The project is intended as an additional safety layer for organizations and individuals who regularly exchange confidential Microsoft Office documents.

---

## Features

* Mark Microsoft Office documents as **reserved**.
* Windows Explorer context menu integration.
* Outlook add-in that detects reserved Office documents attached to outgoing emails.
* Automatic detection of external recipients.
* Confirmation dialog before sending confidential documents outside the organization.
* Uses Microsoft Office Open XML custom document properties.
* No proprietary document format.
* Fully open source.

---

## Privacy First

OfficeFileGuard is designed with a **privacy-first** approach.

* **All document inspection is performed locally on the user's computer.**
* **No documents, attachments, metadata, or email information are uploaded to external servers.**
* **No cloud services are used.**
* **No telemetry or user tracking is performed.**
* **No personal or organizational data is collected.**

OfficeFileGuard operates entirely offline. The application only reads the metadata required to determine whether an Office document has been marked as **reserved**, and all processing remains on the local machine.

---

## Repository Contents

The solution contains the following projects.

### Production projects

| Project                              | Description                                                                                              |
| ------------------------------------ | -------------------------------------------------------------------------------------------------------- |
| **CrocoOfficeFileGuard.Core**        | Core library implementing Office document inspection and management of the `reserved` document property. |
| **CrocoIconOverlayShellExtension**   | Windows Explorer shell extension providing context menu integration.                                     |
| **CrocoOfficeFileReserveToggle.Cli** | Command-line utility used to add or remove the `reserved` custom document property.                      |
| **OutlookFileGuard_addin**           | Microsoft Outlook VSTO add-in that inspects outgoing emails before sending.                              |

### Test project

| Project                             | Description                      |
| ----------------------------------- | -------------------------------- |
| **CrocoOfficeFileGuard.Core.Tests** | Unit tests for the core library. |

---

## How It Works

OfficeFileGuard stores a Boolean custom document property named **reserved** inside Microsoft Office Open XML documents (`.docx`, `.xlsx`, `.pptx`).

The workflow is straightforward:

1. The user marks a document as **reserved** using the Windows Explorer context menu.
2. The custom property is stored directly inside the Office document.
3. When Outlook sends an email, the add-in inspects Office attachments.
4. If one or more reserved documents are attached and at least one recipient belongs to a different SMTP domain, OfficeFileGuard asks the user for explicit confirmation before the email is sent.

The Office document itself remains fully compatible with Microsoft Office because the implementation relies entirely on the Open XML standard.

---

## Architecture

The solution is divided into independent components.

* **CrocoOfficeFileGuard.Core** implements Office document inspection and metadata management.
* **CrocoOfficeFileReserveToggle.Cli** modifies Office document properties.
* **CrocoIconOverlayShellExtension** invokes the CLI from the Windows Explorer context menu.
* **OutlookFileGuard_addin** performs email inspection before sending.
* **CrocoOfficeFileGuard.Core.Tests** validates the core library.

Additional implementation details are available in **ARCHITECTURE.md**.

---

## Building

### Requirements

* Visual Studio 2022 or later
* .NET Framework Developer Pack required by the solution
* Microsoft Office (for Outlook add-in development)
* NuGet Package Restore enabled

Open the solution in Visual Studio and build it normally:

```text
Build → Build Solution
```

---

## Usage

### Mark a document

Right-click a supported Microsoft Office document:

* `.docx`
* `.xlsx`
* `.pptx`

and select the **OfficeFileGuard** context menu command.

The command toggles the `reserved` document property.

### Send an email

Attach one or more Office documents in Outlook.

If reserved documents are detected and one or more recipients belong to an external SMTP domain, OfficeFileGuard asks for confirmation before the email is sent.

---

## Supported File Formats

* Microsoft Word (.docx)
* Microsoft Excel (.xlsx)
* Microsoft PowerPoint (.pptx)

Legacy binary Office formats (`.doc`, `.xls`, `.ppt`) are not supported.

---

## Project Status

OfficeFileGuard is actively maintained.

The current version provides a complete end-to-end workflow including:

* Office document marking
* Windows Explorer integration
* Outlook protection against accidental external disclosure
* Automated unit tests
* Local-only document inspection
* No telemetry
* No cloud connectivity

---

## Contributing

Contributions are welcome.

Bug reports, feature requests and pull requests are appreciated.

Please read **CONTRIBUTING.md** before submitting changes.

---

## Security

OfficeFileGuard is intended to reduce accidental information disclosure.

It is **not** a replacement for enterprise Data Loss Prevention (DLP), Information Protection, or document classification systems.

Security reporting instructions are available in **SECURITY.md**.

---

## License

OfficeFileGuard is licensed under the **Apache License 2.0**.

See the **LICENSE** file for details.
