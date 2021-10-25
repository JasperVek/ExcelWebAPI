Design:

Het idee van de applicatie is dat je een IModel en een IItem hebt als interfaces.
Deze vul je in als concrete classes en initializeer je, en de reader heeft dan iets om mee te werken omdat het de types kan 
achterhalen van de IModel en IItem.
Een IModel is een container en een IItem is een rij met waardes binnen Excel.
Het kan zijn dat er verschillende niveaus nesting zijn dus vandaar deze 2 interfaces.

Uiteindelijk pas je voor een ander Excel bestand de Domein items aan, 
waar je de aanpassingen IModel en IItem laat gebruiken en de juiste implementaties maakt in een Reader.
Hierna moet je aangeven welk hoofd object je gebruikt en stuurt naar de Reader, 
in mijn geval AllInvestments, en de rest zou de Reader moeten kunnen doen om dit object te vullen.

Deze voorlopige versie probeert een IModel terug te geven als object, wat niet werkt in JSON, en dus heeft het geen nuttige output in JSON.
De GetFile methode van de API geeft in dit geval wel echt een IModel terug, met 2 IModels, die beide 3 IItems bevatten.

Het Excel bestand heeft 2 verschillende tabellen onder elkaar maar dit wordt nog niet uitgelezen, omdat ik nog niet heb geimplementeerd
dat het eerst de IModels zoekt in Excel, waarna het dit gebruikt als startpunt om voor de properties te zoeken.
Nu zoekt het de properties vanaf het begin van Excel.