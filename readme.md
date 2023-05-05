# WOWIPORT / OPENWOWI - Ausgabe von Mieterjubiläen
### Allgemein
Für bestimmte  Aktivitäten ist es notwendig, alle Mieter in einer Liste auszugeben, die insgesamt länger als X Jahre  
Mieter des Unternehmens sind.

### Filterkriterien
* Nur Wohnraum
* Nur natürliche Personen (keine Unternehmen)
* Summierte Mietdauer größer oder gleich X (Überlappungen durch mehrere gleichzeitige Verträge werden bereinigt)
* Gültiger Mietvertrag zum einstellbaren Stichtag

### Berechtigungen OPENWOWI
Der verwendete API-Key muss die folgenden Berechtigungen haben:
* Objektdaten
* Personen lesen
* Vertragsdaten mit persönlichen Details
* Vertragsdaten ohne persönliche Details

### Ausgabe
Excel-Liste mit den folgenden Spalten:
* Name
* Vorname
* Titel
* Anrede
* Straße
* PLZ
* Ort
* Geburtsdatum
* Personennummer in Wowiport
* Anzahl der Mietjahre

### Einstellbare Parameter
* Minimale Anzahl der Mietjahre
* Stichtag für aktiven Vertrag
* Filter für Wirtschaftseinheiten ("Startet mit...")
* Filter füür Objektnummern ("Objektnummer kleiner als...")

### Prozess
Die Anwendung ruft zunächst alle Mietverträge und alle Vertragsnehmer viw OPENWOWI aus Wowiport ab. Dies kann je  
nach Anzahl 20-30 Minuten dauern. Anschließend werden alle Mietverträge für die einzelnen Mieter aufaddiert. Es  
werden hierbei natürlich auch beendete Verträge berücksichtigt, da der Mieter im Bestand umgezogen sein kann.  
Sollte ein Mieter mehrere Verträge haben, so wird die Überlappung herausgerechnet. Beispiel: Ein Mieter ist vor  
10 Jahren einzogen und hat vor 5 Jahren eine Garage bekommen. Die gesamte Mietdauer ist dadurch natürlich nicht  
15 sondern trotzdem 10 Jahre.

### Installation
* Repository klonen oder Paket als ZIP herunterladen und extrahieren
* Datei .env.example kopieren und eine .env Datei erstellen
* Werte in der .env Datei anpassen
* via Kommandozeile in den Ordner navigieren
* pip install -r requirements.txt ausführen, um die notwendigen Pakete zu installieren
* Script mit python3 mietlaenge\mietlaenge.py aufrufen
* Datei output.xlsx wird im Hauptordner erstellt

### Datenschutzhinweise 
* Es wird ein Zwischenspeicher aus allen abgerufenen Daten von Wowiport erstellt. Nach dem Aufruf  
des Scriptes sollten die Dateien wowi_contractors.json und wowi_nv.json aus dem Ordner entfernt werden.
* Api-Key, Kennwort und Benutzername liegen im Klartext in der .env-Datei vor. Diese Datei sollte entweder  
vor unbefugtem Zugriff geschützt werden oder nach dem Abruf direkt wieder gelöscht werden.