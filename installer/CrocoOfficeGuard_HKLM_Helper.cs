static void Main()
{
    // 1. scrivi HKLM
    Registry.SetValue(
        @"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\     !CrocoOffice",
        "",
        "{098AA80F-2F55-4AA8-B95F-8AF77120E479}"
    );

    // 2. exit subito
}

Main();