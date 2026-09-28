# Nutzung / Usage

## Seiten anlegen

**Page → Create**: Abschnitt wählen, Titel, Inhalt und Format (*Plain Text*, *Markdown*, *HTML*).
Über *Binary Properties* (`data`, `data,foto` oder `*`) werden Binärdaten aus vorherigen Nodes angehängt
(z. B. *Webhook*, *Read Binary File*, *HTTP Request*, *Telegram*, *E-Mail*):

- `image/*` → eingebettetes Bild
- `video/*` → Upload nach OneDrive (Ordner *n8n OneNote Videos*) und Link in der Seite
  (Option *Video Handling* → *Attach as File* versucht stattdessen einen Datei-Anhang)
- alles andere (Audio, PDF, Office …) → Datei-Anhang

Ausgabe: Seiten-Objekt mit `id`, `title`, `webUrl`, `clientUrl` (öffnet die OneNote-App).

## Seiten lesen

**Page → Get**: Metadaten plus `html` und/oder `text`. Mit *Download Media* landen alle Bilder und Anhänge als
Binärdaten (`data0`, `data1`, …) im Item; die Liste steht in `media`. Alternativ **Media → Download From Page**
mit Filter (nur Bilder / nur Anhänge) und eigenem Präfix.

**Page → Get Many**: optional pro Abschnitt, Filter Titel/Erstellt/Geändert, Sortierung, *Return All* oder Limit.

## Seiten ändern

**Page → Update** – *Append*/*Prepend* zum Seiten-Ende/-Anfang. *Replace*/*Insert* arbeiten auf einem Element:
zuerst **Get** mit *Include Element IDs*, dann die Element-ID (`p:{…}`) angeben.
**Media → Add to Page** hängt nur Dateien (mit optionaler Beschriftung) an.

## Notizbücher & Abschnitte

**Notebook / Section → Get Many** liefert IDs und Namen für nachfolgende Nodes
(z. B. `{{ $json.id }}` als *By ID* im Dropdown-Feld: Modus **By ID** wählen oder Expression verwenden).

## Trigger

**OneNote System Trigger** fragt regelmäßig (Standard: Poll Times des Nodes) neue (*Page Created*) oder neue und
geänderte Seiten (*Page Created or Modified*) ab, optional mit Inhalt. Beim ersten Aktivieren startet er ab „jetzt".
Im Test-Modus wird die neueste Seite geliefert.

## Als KI-Tool

Der Node ist `usableAsTool`; N8N_COMMUNITY_PACKAGES_ALLOW_TOOL_USAGE=true muss gesetzt sein (in `docker-compose.yml` enthalten).
