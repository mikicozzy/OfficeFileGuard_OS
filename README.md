# OfficeFileGuard

OfficeFileGuard is an open-source Windows toolset designed to reduce the risk of accidental disclosure of Microsoft Office documents.

Instead of relying solely on user awareness, OfficeFileGuard allows Office documents to be explicitly marked as **reserved**. Whenever a marked document is attached to an Outlook email addressed to recipients outside the sender's organization, the Outlook add-in requires explicit confirmation before the message is sent.

The project is intended as an additional safety layer for organizations and individuals who regularly exchange confidential Office documents.

---

## Features

* Mark Microsoft Office documents as **reserved**.
* Windows Explorer context menu integration.
* Outlook add-in that detects reserved Office documents attached to outgoing emails.
* Automatic detection of external recipients.
* Confirmation dialog before sending confidential documents outside the organization.
* Open XML based implementation (no document format modifications).

---

## Repository contents

The solution currently contains the following projects.

### Production projects

| Project                              | Description                                                                                         |
| ------------------------------------ | --------------------------------------------------------------------------------------------------- |
| CrocoOfficeFileGuard.Core        | Core library implementing Office document inspection and reserved property management.              |
| CrocoIconOverlayShellExtension   | Windows Explorer shell extension providing context menu integration.                                |
| CrocoOfficeFileReserveToggle.Cli | Command-line utility used to add or remove the `reserved` custom document property.                 |
| OutlookFileGuard_addin           | Microsoft Outlook VSTO add-in that intercepts outgoing emails and checks attached Office documents. |

### Test project

| Project                             | Description                      |
| ----------------------------------- | -------------------------------- |
| CrocoOfficeFileGuard.Core.Tests | Unit tests for the core library. |

---

## How it works

OfficeFileGuard stores a Boolean custom document property named **reserved** inside Microsoft Office Open XML documents (`.docx`, `.xlsx`, `.pptx`).

The workflow is straightforward:

1. The user marks a document as **reserved** using the Windows Explorer context menu.
2. The custom property is stored directly inside the Office document.
3. When Outlook sends an email, the add-in inspects Office attachments.
4. If one or more reserved documents are attached and at least one recipient belongs to a different SMTP domain, OfficeFileGuard asks the user for explicit confirmation before the email is sent.

No proprietary document format is used; the implementation relies entirely on the Open XML standard.

---

## Architecture

The solution is divided into independent components.

* **Core library** implements document inspection and metadata management.
* **CLI utility** modifies Office document properties.
* **Explorer shell extension** invokes the CLI from the Windows context menu.
* **Outlook add-in** performs email inspection before sending.
* **Unit tests** validate the core library.

Additional implementation details are available in **ARCHITECTURE.md**.

---

## Building

### Requirements

* Visual Studio 2022 or later
* .NET Framework Developer Pack required by the solution
* Microsoft Office (for Outlook add-in development)
* NuGet Package Restore enabled

Clone the repository and build the solution normally:

```text
Build → Build Solution
```

---

## Usage

### Mark a document

Right-click a supported Office document:

* `.docx`
* `.xlsx`
* `.pptx`

and select the OfficeFileGuard context menu command.

The command toggles the `reserved` document property.

### Sending email

Attach one or more Office documents in Outlook.

If reserved documents are detected and one or more recipients belong to an external SMTP domain, OfficeFileGuard asks for confirmation before the email is sent.

---

## Supported file formats

* Microsoft Word (.docx)
* Microsoft Excel (.xlsx)
* Microsoft PowerPoint (.pptx)

Legacy binary Office formats are not supported.

---

## Project status

OfficeFileGuard is actively maintained.

The current version provides a complete end-to-end workflow including:

* document marking;
* Windows Explorer integration;
* Outlook protection against accidental external disclosure;
* automated tests for the core library.

---

## Contributing

Contributions are welcome.

Bug reports, feature requests and pull requests are appreciated.

Please read **CONTRIBUTING.md** before submitting changes.

---

## Security

OfficeFileGuard is intended to reduce accidental information disclosure.

It is **not** a replacement for enterprise Data Loss Prevention (DLP), document classification or information protection systems.

Security reporting instructions are available in **SECURITY.md**.

---

## License

This project is released under the MIT License.

See **LICENSE** for details.
