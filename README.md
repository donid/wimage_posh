# wimage_posh

A Powershell version of the c't wimage batch script

**Eine Powershell 7 Portierung des c't wimage batch skripts.**

*Das Script ist für eher fortgeschrittene Anwender gedacht. Das c't Magazin oder der Heise-Verlag
stehen in keiner Verbindung zu dieser  Powershell Portierung !*

Dieses Script ist nur eine Alternative zur Datei **ct-WIMage.x64.bat**.
Das Einrichten des USB-Sticks muss vorher mit den c't Werkzeugen gemacht worden sein.
Dann kann **wimage_posh.ps1** einfach "neben" **ct-WIMage.x64.bat** kopiert und ausgeführt werden.

Bitte lesen Sie unbedingt die Anleitungen zu dem **Original** Skript, siehe <https://ct.de/wimage>

Hilfe anzeigen:

    get-help .\wimage_posh.ps1

### Features

- automatische Windows Defender Ausnahme beschleunigt den Vorgang stark
- keine Probleme durch unterschiedliche Betriebssystem-Sprachen
- Windows 11 wird im Image-Namen und der Description richtig benannt
- Code (hoffentlich) einfacher zu verstehen und zu modifizieren
- mehr Ausgaben, die Problemdiagnosen erleichtern
