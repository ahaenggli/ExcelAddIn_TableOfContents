# Excel AddIn: table of contents [![GitHub release (latest by date)](https://img.shields.io/github/v/release/ahaenggli/ExcelAddIn_TableOfContents?style=social)](https://github.com/ahaenggli/ExcelAddIn_TableOfContents)
[![paypal](https://www.paypalobjects.com/de_DE/CH/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=S2F6JC7DGR548&source=url)
<a href="https://www.buymeacoffee.com/ahaenggli" target="_blank"><img src="https://cdn.buymeacoffee.com/buttons/default-orange.png" alt="Buy Me A Coffee" height="50px" width="217px" ></a>

## Features
- Inhaltsverzeichnis über alle sichtbaren Register generieren  
    ![ButtonGenerieren](.github/img/screenshot_button.png)  
    ![TOC](.github/img/screenshot_button2.png) 
    - Neue/Gelöschte/Verschobene Register werden automatisch nachgeführt
    - Erstelldatum zu jedem neuen Register
    - Erste Spalte enthält Link um schnell ins Register zu kommen
    - Felder sind als Custom-Property auf ursprünglichen Registern gespeichert.  
    Bedeutet: Auch bei versehentlichem löschen des Inhaltsverzeichnisses bleiben die Infos erhalten! 

- Einstellungen zur Gestaltung des Inhaltsverzeichnisses  
    ![Einstellungen](.github/img/screenshot_settings.png)  
    - Optional: Vorlage kann auch nur für das aktuelle Excel angepasst werden
- Es können irgendwelche Zusatzinfos pro Register gespeichert werden.  
  Sind diese nicht im Inhaltsverzeichnis definiert, sind die Werte nirgends sonst direkt ersichtlich. 
  ![Info](.github/img/screenshot_customproperties.png)
- Auto-Update Funktion (in Info-Dialog deaktivierbar)  
    ![Info](.github/img/screenshot_info.png)

## Changelog
... [findet sich hier](CHANGELOG.md) ...

## Auto-Update?
Beim Starten von Excel, 1x pro 24h, wird auf GitHub die Version überprüft. Gibt es eine neuere Version, wird diese als zip-Datei heruntergeladen und entpackt. Beim nächsten Excel-Neustart wird die neuere Version dann via ClickOnce-Update nachgeführt.

Wenn Excel nie geschlossen und gestartet wird, wird auch kein Update installiert.

## Fehler melden
Es dürfen gerne [hier](https://github.com/ahaenggli/ExcelAddIn_TableOfContents/issues) in GitHub Issues erfasst werden.