# Beispiel-Workflows / Example workflows

Import in n8n: *Workflows → Import from File*. Danach in jedem OneNote-Node das
Credential und Notizbuch/Abschnitt auswählen.

| Datei | Beschreibung |
| --- | --- |
| `01-webhook-to-note.json` | Webhook (POST `title`, `text`) → neue Seite (Markdown) |
| `02-web-clip.json` | Webhook (POST `url`) → Webseite laden → als Seite mit Titel speichern |
| `03-new-page-trigger.json` | Trigger: neue Seite → Inhalt als Text weiterverarbeiten |
| `04-read-page-with-media.json` | Seite lesen inkl. Bilder/Audio als Binärdaten |
| `05-image-to-note.json` | Webhook mit Datei-Upload (Bild/Audio/Video) → Seite |
| `06-device-login.json` | Einmaliger Geräte-Code-Login (Refresh-Token für das Credential *OneNote Device Login API*) |
