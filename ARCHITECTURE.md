# Architecture

## Overview

OfficeFileGuard is a modular Windows application designed to reduce the risk of accidental disclosure of Microsoft Office documents.

The solution consists of independent components that cooperate to provide:

* document marking;
* Windows Explorer integration;
* Explorer icon overlays;
* Outlook protection against accidental external email disclosure.

The business logic is centralized inside **CrocoOfficeFileGuard.Core**, while the remaining projects provide user interfaces and Windows integration.

---

# Solution Structure

```
OfficeFileGuard.sln
│
├── CrocoOfficeFileGuard.Core
├── CrocoOfficeFileReserveToggle.Cli
├── CrocoIconOverlayShellExtension
├── OutlookFileGuard_addin
└── CrocoOfficeFileGuard.Core.Tests
```

---

# Project Overview

## CrocoOfficeFileGuard.Core

**Type**

.NET Class Library

**Responsibilities**

* Manage the `reserved` state of Office documents.
* Read and write the Open XML custom document property.
* Maintain the Windows Registry cache used by the Explorer icon overlay.
* Detect supported Office document types.
* Provide a common API shared by the CLI, Shell Extension and Outlook Add-in.

This project contains all business logic and should remain independent from any user interface.

---

## CrocoOfficeFileReserveToggle.Cli

**Type**

.NET Console Application

**Responsibilities**

* Toggle the reserved state of Office documents.
* Invoke the core library from scripts or external applications.
* Provide a lightweight executable used by Windows Explorer integration.

The CLI contains no business logic; it delegates all operations to **CrocoOfficeFileGuard.Core**.

---

## CrocoIconOverlayShellExtension

**Type**

Windows Shell Extension

**Responsibilities**

* Display an icon overlay for Office documents marked as reserved.
* Query the Windows Registry cache for maximum Explorer performance.

Windows Explorer requests icon overlay information very frequently. Reading Office documents every time would significantly impact performance.

For this reason the shell extension relies exclusively on the Registry cache maintained by **CrocoOfficeFileGuard.Core**, avoiding expensive document inspection during normal Explorer operation.

---

## OutlookFileGuard_addin

**Type**

Microsoft Outlook VSTO Add-in

**Responsibilities**

* Intercept outgoing email messages.
* Detect Office document attachments.
* Verify whether attached Office documents are marked as reserved.
* Detect recipients belonging to external SMTP domains.
* Ask the user for confirmation before sending confidential documents outside the organization.

The add-in performs document inspection through **CrocoOfficeFileGuard.Core**.

---

## CrocoOfficeFileGuard.Core.Tests

**Type**

xUnit Test Project (.NET 9)

**Responsibilities**

* Validate the business logic implemented in **CrocoOfficeFileGuard.Core**.
* Verify document property handling.
* Validate Registry synchronization.
* Prevent regressions during future development.

---

# Reserved State Storage

OfficeFileGuard stores the reserved state using two complementary mechanisms.

## Office Document Property

The authoritative source is a Boolean custom document property named:

```
reserved
```

stored inside Microsoft Office Open XML documents.

This information travels with the document, regardless of where the file is copied.

---

## Windows Registry Cache

To guarantee fast Windows Explorer performance, OfficeFileGuard also maintains a Registry cache.

```
HKEY_CURRENT_USER
└── Software
    └── CrocoOfficeFileGuard
        └── ReservedFiles
```

Each reserved document is represented by a Registry value whose name is the full file path.

The Registry cache is used **only** by the Explorer icon overlay.

Whenever a document is marked or unmarked as reserved, both the document property and the Registry cache are updated atomically.

---

# Dependency Graph

```
                     +-------------------------------+
                     | CrocoOfficeFileGuard.Core     |
                     +-------------------------------+
                          ▲          ▲           ▲
                          │          │           │
                          │          │           │
        +-----------------+          │           +----------------------+
        │                            │                                  │
        │                            │                                  │
+-------------------------+   +----------------------+        +------------------------+
| ReserveToggle.Cli       |   | Outlook Add-in      |        | IconOverlayExtension   |
+-------------------------+   +----------------------+        +------------------------+

                         ▲
                         │
                +----------------------+
                | Core.Tests (xUnit)   |
                +----------------------+
```

---

# Build Requirements

* Visual Studio 2022 or later
* .NET Framework Developer Pack required by the solution
* Microsoft Office (for Outlook add-in development)
* VSTO Runtime
* NuGet Package Restore enabled

---

# Design Principles

OfficeFileGuard is built around a few key principles.

* **Single source of business logic** inside the Core library.
* **Privacy First**: all processing is performed locally.
* **No cloud services.**
* **No telemetry.**
* **No document upload.**
* **Performance-oriented Windows Explorer integration** through Registry caching.
* **Modular architecture**, allowing each component to evolve independently.

---

# Future Evolution

The modular architecture allows additional integrations to be developed without modifying the core business logic.

Possible future integrations include additional Office applications, Windows shell features, or other document workflows.
