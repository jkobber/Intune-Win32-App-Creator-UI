Kurzinfo
Dieses Projekt erstellt .intunewin-Pakete ueber einen GUI-Wizard (WPF).
Git ist optional; falls nicht vorhanden, wird per ZIP von GitHub geladen.

Ablauf (kurz)
1) Tool wird geladen, Arbeitsordner wird unter C:\temp\IntuneWinBuilder erstellt
2) Setup-Dateien in Setup-folder ablegen, EntryPoint waehlen
3) IntuneWin erstellen, optional weitere Pakete
4) Export nach Downloads + Cleanup der Arbeitsordner

Arbeitsordner
Alle temporaeren Dateien liegen unter:
C:\temp\IntuneWinBuilder

Start
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File ".\\IntuneWinBuilder.ps1"
