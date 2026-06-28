windows explorer overlay icon non consente nessun algoritmo di selezione (è tutto sincrono) 
accetta solo lettura di un booleano per decidere se mostrare o meno l'icona di overlay.

Allo scopo si userà il 'registry' di windows per memorizzare globalmente i path/nomi dei file reserved.
Il registry sarà aggiornato solo on-demand (quando ad un file viene modificato l'attributo
custom 'reserved'). 

**TODO:** 
- in futuro usare FileSystemWatcher per tenere traccia degli spostamenti/cancellazioni dei file riservati.
- in futuro implementare un sistema di scanning delle directory per aggiornare il registry in caso di mismatch.
- in futuro implementare un sistema di caching in memoria per evitare continue letture del registry.


**Adesso:** 
- Implementata la modifica del registry in CrocoOfficeFileReserveToggle.Cli
- Realizzare estensione della Shell di Windows implementando l'interfaccia IShellIconOverlayIdentifier
    - implementata la logica che legge il registry e decide se mostrare o meno l'icona di overlay (ReservedFileRegistry.IsOverlayNeeded)
	- utilizzare la lireria SharpShell che semplifica lo sviluppo della shell extension

### Windows Registry
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

Tale struttura sarà la single source of truth per l'iconoverlay.

### IconOverlayShellExtension project 
Nel project IconOverlayShellExtension si userà la libreria Microsoft.Win32.Registry per leggere il registry e decidere se mostrare o meno l'icona di overlay.
Il filtro delle estensioni supportate (docx, xlsx, pptx) va inserito nell'IconOverlayHandler.

    string ext = Path.GetExtension(filePath).ToLowerInvariant();
    if (ext != ".docx" && ext != ".xlsx" && ext != ".pptx")
        return false;


funzione:
    bool IsReserved(string filePath)
    {
        using var key = Registry.CurrentUser.OpenSubKey(
            @"Software\CrocoOfficeFileGuard\ReservedFiles"
        );

        return key?.GetValue(filePath) != null;
    }


## Architettura futura: verifica asincrona della custom property

### Problema
Il registry è la single source of truth per l'overlay icon, ma può diventare 
inconsistente se un file riservato viene spostato, rinominato o cancellato.

### Soluzione
Architettura a due livelli:

**Livello 1 — IsMemberOf (sincrono, veloce)**
- Legge solo il registry (comportamento attuale)
- Se il file ha estensione supportata (.docx, .xlsx, .pptx), aggiunge il path 
  a una ConcurrentQueue<string>
- Ritorna subito il risultato dal registry senza bloccare Explorer

**Livello 2 — Worker asincrono (parallelo)**
- Gira in background finché la DLL è caricata
- Legge i path dalla coda uno alla volta
- Per ogni path chiama IsReservedInFile (legge la custom property dal file Office)
- Se il risultato è diverso da quello nel registry, aggiorna il registry
- Notifica Explorer del cambiamento tramite SHChangeNotify (P/Invoke su shell32.dll)

### TODO
- [fatto] Step 1: modificare IsReservedInFile per usare MemoryStream invece di 
      salvare su disco (più veloce, nessun file temporaneo residuo in caso di crash)
- [fatto] Step 2: implementare ConcurrentQueue in IsMemberOf
- [fatto] Step 3: implementare il worker asincrono con CancellationToken
- [fatto] Step 4: implementare SHChangeNotify per il refresh di Explorer
