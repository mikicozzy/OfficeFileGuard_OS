# Solution Architecture

## Overview

This repository contains a multi-project C# solution designed to enforce file privacy 
and prevent sensitive documents from being sent as email attachments to external organizations.
The solution includes core protection logic and integrations with Outlook add-ins and 
shell extensions to apply these rules at different interaction points.

The solution is organized in modular projects, each with a its responsibility.

---

## Solution Structure

CrocoOfficeFileGuard.sln
│
├─ CrocoOfficeFileReserveToggle.Cli
├─ ////////////////////CrocoShellHandler_NON_SERVE
└─ OutlookFileGuard_addin


---

## Projects Description

### 1. CrocoOfficeFileReserveToggle.Cli

**Type:** .NET Console Application  

**Responsibility:**
- Provides a command-line interface to toggle file 'reservation' state
- Acts as a low-level control tool
- Can be invoked manually or by Windows Explorer Contextual Menu (the latter requiring registration on the windows register)

**Key characteristics:**
- No UI
- Focused on core business logic
- Designed to be deterministic and script-friendly

---

### 2. CrocoShellHandler_NON_SERVE

**Type:** Shell integration module  
**Responsibility:**

- (Describe what it *should* do or why it exists)
- Currently marked as **NON_SERVE**:
  - experimental
  - deprecated
  - or placeholder for future shell integration

**Notes:**
- This project is currently not used in production
- Kept for reference or future development

---

### 3. OutlookFileGuard_addin

**Type:** Outlook VSTO Add-in  

**Responsibility:**
- Integrates file protection logic directly into Microsoft Outlook
- Intercepts and reacts to attachments of 'reserved' Office files
- Asks users to confirm or block sending 'reserved' Office files to external email domanins

**Key characteristics:**
- UI-driven
- Runs inside Outlook process
- Security-sensitive (code signing required)  TBV

---

## Dependency Flow

**TO BE CHANGED AFTER .Core LIB ADDED**
Outlook Add-in
↓
CLI / Core Logic


- UI components depend on core logic
- Core logic does NOT depend on UI components

---

## Windows Registry 
Per memorizzare i file riservati si userà la sezione HKEY_CURRENT_USER (per singolo utente) del registry di windows:

- HKEY_CURRENT_USER\Software\CrocoOfficeFileGuard\ReservedFiles
- valore: path completo del file riservato (ogni file è una voce separata)
- tipo: REG_SZ

HKEY_CURRENT_USER
 └─ Software
    └─ CrocoOfficeFileGuard
       └─ ReservedFiles
          ├─ C:\Docs\test.docx    = ""
          ├─ C:\Excel\data.xlsx  = ""
          └─ C:\Slides\demo.pptx  = ""

---

## Build & Deployment Notes

- Solution built using Visual Studio
- Outlook add-in requires:
  - code signing certificate
  - VSTO runtime

---

## Future Improvements

- TBA
