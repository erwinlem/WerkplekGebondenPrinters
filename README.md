# Project doelen

Het doel is om gebruikers per werkplek hun eigen printer te laten kiezen, deze
geldt hierna voor iedereen die hier op deze werkplek inlogt. Hix kan door de
omschrijving uit te lezen bepalen welke type printers gekoppeld zijn en
hierdoor direct de juiste printer gebruiken voor een bepaald etiket.

# Uitgangspunten

* Zo simpel mogelijk.
* Generiek.
* Makkelijk te onderhouden.
* [Semantic Versioning](https://semver.org/)

# Implementatie

Het geheel is gebouwd in powershell, printers worden uitgelezen uit active
directory.  Als opslag worden files gebruikt met de hostname.

Powershell is echt te traag bij het opstarten, omgezet naar c#.

# Verbeterpunten

* Script maken wat aangeroepen kan worden als er gewisseld wordt van werkplek.
* Filteren op hostname zodat thuiswerkers hun machine namen er niet inkomen (bijv. wr_b81oiunmoul).
* Meer feedback over wat er op de achtergrond gebeurd (printers toevoegen gaat
  niet snel).
* Config file met instellingen.
* Details over printer bij selecteren (printserver/sticker type etc).
* Server naam meenemen zodat dezelfde printernamen gebruikt kunnen worden.

# Tips

## Opstarten

Standaard kan je get direct starten met `WerkplekGebondenPrinter.exe`
Je kan kiezen om andere loaders te gebruiken zoals active directory of sql :
`
WerkplekGebondenPrinter-0.1.0.exe -d  -l h:\wpg-client.txt --cwd u:\Werkplekgebondenprinter2 --PrinterLoader PrinterLoaderAD
`
