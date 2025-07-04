# Project doelen

Het doel is om gebruikers per werkplek hun eigen printer te laten kiezen, deze
geldt hierna voor iedereen die hier op deze werkplek inlogt. Hix kan door de
omschrijving uit te lezen bepalen welke type printers gekoppeld zijn en
hierdoor direct de juiste printer gebruiken voor een bepaald etiket.

# Uitgangspunten

* Zo simpel mogelijk.
* Generiek.
* Makkelijk te onderhouden.

# Implementatie

Het geheel is gebouwd in powershell, printers worden uitgelezen uit active
directory.  Als opslag worden files gebruikt met de hostname.

# Verbeterpunten

* Script maken wat aangeroepen kan worden als er gewisseld wordt van werkplek.
* Filteren op hostname zodat thuiswerkers.
* Meer feedback over wat er op de achtergrond gebeurd (printers toevoegen gaat
  niet snel).

