# n8n-nodes-onenote-system

**Deutsch** · [English](#english)

n8n-Community-Nodes für **Microsoft OneNote mit privatem Microsoft-Konto**. Notizen anlegen, lesen und
ergänzen – mit Text, Markdown, HTML, Bildern, Audio, Dateien und Videos – in beide Richtungen.

## Funktionen

| Ressource | Operationen |
| --- | --- |
| **Notebook** | Liste, Details, Erstellen |
| **Section** | Liste (optional pro Notizbuch), Details, Erstellen |
| **Page** | Erstellen, Lesen (HTML/Text/Metadaten, optional inkl. Medien), Liste mit Filtern, Ergänzen/Ersetzen, Kopieren, Verschieben, Löschen |
| **Media** | Bilder/Audio/Dateien einer Seite als Binärdaten herunterladen · Binärdaten an eine Seite anhängen |
| **Trigger** | Polling: neue bzw. geänderte Seiten (gesamt, pro Notizbuch oder pro Abschnitt) |

- Notizbuch, Abschnitt und Seite wählst du per **Dropdown** (mit Suche) oder per **ID/Expression**; jedes Ergebnis
  enthält die IDs (`id`, `notebookId`, `sectionId`, `webUrl`, …) zum Weiterverwenden.
- **Inhalte:** Text, Markdown (Überschriften, Listen, To-dos, Tabellen, Code, Links, Bilder per URL) oder rohes HTML.
- **Bilder** werden eingebettet, **Audio/PDF/andere Dateien** als Anhang. **Videos** kann OneNote nicht einbetten:
  sie werden nach OneDrive hochgeladen und als Link in die Seite eingefügt (bis 250 MB; abschaltbar).
- Mehrere Binärfelder auf einmal (`data,foto`, oder `*` für alle).
- Automatische Wiederholung bei Drosselung (HTTP 429) der Microsoft-Graph-API.
- Keine Laufzeit-Abhängigkeiten – geeignet für die Verifizierung als Community-Node.

## Installation

**Aus npm** (nach dem Veröffentlichen): n8n → *Settings → Community Nodes → Install* → `n8n-nodes-onenote-system`.

**Lokal aus dem Quellcode (Docker):**

```bash
npm install
npm run build
docker compose up -d          # n8n auf http://localhost:5678
```

## Einrichtung des Zugangs

Für private Konten ist eine (kostenlose) App-Registrierung in Azure nötig – Schritt für Schritt in
[docs/setup-azure.md](docs/setup-azure.md). Kurz: Redirect-URI `http://localhost:5678/rest/oauth2-credential/callback`,
Client-ID + Client-Secret im n8n-Credential **OneNote OAuth2 API** eintragen, *Connect* klicken.

## Nutzung

Siehe [docs/usage.md](docs/usage.md) und die importierbaren Workflows in [examples/](examples/).

## Entwicklung & Veröffentlichung

```bash
npm install
npm run build     # kompiliert nach dist/
npm run lint      # n8n-Community-Node-Regeln
npm test          # Unit-Tests (Node test runner, Graph-API gemockt)
npm run dev       # n8n mit Hot-Reload (n8n-node dev)
npm run release   # Version anheben, Tag setzen, veröffentlichen (release-it)
```

Veröffentlichen: `npm login` und `npm publish` (der Paketname muss mit `n8n-nodes-` beginnen). Für die
n8n-Verifizierung wird das Paket über GitHub Actions mit npm-Provenance veröffentlicht.

## Grenzen der OneNote-API

- Kein Webhook/Änderungs-Feed → der Trigger arbeitet per Polling.
- Videos lassen sich nicht einbetten (siehe oben). Audio wird als Datei-Anhang gespeichert.
- Seiten-Inhalt kann nur angehängt/vorangestellt bzw. über Element-IDs ersetzt werden, nicht als Ganzes überschrieben.
- Sehr große Notizbücher in OneDrive (> 5000 Elemente) können von der API mit Fehlern abgelehnt werden.
- Ein Notizbuch/Abschnitt mit Passwortschutz ist über die API nicht lesbar.

---

## English

n8n community nodes for **Microsoft OneNote with a personal Microsoft account**: create, read and extend notes with
text, Markdown, HTML, images, audio, files and videos – in both directions.

- Resources: Notebook, Section, Page, Media, plus a polling **OneNote Trigger** (new/modified pages).
- Pick notebook/section/page from a searchable list or by ID/expression; outputs contain IDs and web URLs.
- Images are embedded, audio/PDF/other files attached; videos are uploaded to OneDrive and linked (OneNote cannot embed video).
- Install via *Settings → Community Nodes* (`n8n-nodes-onenote-system`) or build locally (`npm install && npm run build`,
  `docker compose up -d`).
- Setup of the Azure app registration: [docs/setup-azure.md](docs/setup-azure.md). Usage: [docs/usage.md](docs/usage.md).

License: MIT
