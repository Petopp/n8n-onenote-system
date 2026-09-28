# Azure-App für das private Microsoft-Konto einrichten / Azure app setup

Deutsch zuerst, English below.

## Deutsch

Die Microsoft-Graph-API erlaubt den Zugriff auf OneNote nur über eine registrierte App. Die Registrierung ist
kostenlos und geht mit einem privaten Konto.

1. Öffne <https://entra.microsoft.com> → **App-Registrierungen** → **Neue Registrierung**
   (alternativ <https://portal.azure.com> → *App registrations*). Melde dich mit deinem privaten Konto an.
   Falls kein Verzeichnis vorhanden ist, legt Microsoft beim ersten Aufruf ein kostenloses an.
2. **Name:** z. B. `n8n OneNote`.
3. **Unterstützte Kontotypen:** *Nur persönliche Microsoft-Konten* oder *Konten in allen Organisationsverzeichnissen
   und persönliche Microsoft-Konten*.
4. **Umleitungs-URI:** Plattform **Web**, Wert
   `http://localhost:5678/rest/oauth2-credential/callback`
   (bei anderer n8n-URL: `<N8N-URL>/rest/oauth2-credential/callback`; außerhalb von localhost ist HTTPS Pflicht).
5. **Registrieren.** Kopiere die **Anwendungs-(Client-)ID**.
6. **Zertifikate & Geheimnisse → Neuer geheimer Clientschlüssel** → den **Wert** sofort kopieren (wird nur einmal angezeigt).
7. **API-Berechtigungen** (delegiert, Microsoft Graph): `Notes.ReadWrite`, `offline_access`, `User.Read`
   und – für Video-Upload – `Files.ReadWrite`. Bei privaten Konten wird die Zustimmung beim Anmelden erteilt.
8. In n8n: **Credentials → New → OneNote OAuth2 API**
   - *Client ID* / *Client Secret* eintragen
   - *Account Type:* Personal Microsoft Account
   - *Enable Video Upload:* an, wenn du Videos nutzen willst
   - **Connect my account** → anmelden → Zugriff erlauben.

### n8n läuft nicht auf `localhost` (Server, LAN-IP, kein HTTPS)

Azure akzeptiert `http://` nur für `localhost`, sonst ist HTTPS mit Hostnamen Pflicht (keine IP-Adresse). n8n baut die
Redirect-URL fest aus `N8N_EDITOR_BASE_URL` – sie lässt sich weder im Credential noch in diesem Modul ändern.
Es gibt drei Wege:

**A) Device-Login (empfohlen, ohne Redirect-URI):** siehe Abschnitt
[Ohne Redirect-URI: Device-Login](#ohne-redirect-uri-device-login) unten. Funktioniert mit jeder n8n-Adresse.

**B) SSH-Tunnel, nur für den ersten Login:**

1. Tunnel vom Rechner mit dem Browser zum n8n-Server starten (Fenster offen lassen):
   ```bash
   ssh -L 5678:localhost:5678 benutzer@n8n-server
   ```
2. In der n8n-Konfiguration des Servers setzen und den Container **neu erzeugen** (`restart` liest die Umgebung nicht neu):
   ```
   N8N_EDITOR_BASE_URL=http://localhost:5678
   WEBHOOK_URL=http://localhost:5678/
   ```
   ```bash
   docker compose up -d --force-recreate n8n
   ```
3. n8n im Browser über `http://localhost:5678` öffnen. Das Feld *OAuth Redirect URL* im Credential beginnt jetzt mit
   `http://localhost:5678/…` – diese URL in Azure eintragen und **Connect my account** klicken.
4. Danach kann der Tunnel geschlossen werden; n8n erneuert das Token selbständig. Für Webhooks danach
   `WEBHOOK_URL` wieder auf die echte Adresse stellen.

**C) HTTPS mit Hostnamen** (z. B. Reverse-Proxy oder Tailscale): `N8N_EDITOR_BASE_URL=https://n8n.example.org` setzen und
`https://n8n.example.org/rest/oauth2-credential/callback` in Azure eintragen.

### Ohne Redirect-URI: Device-Login

Für n8n hinter einer LAN-IP oder ohne HTTPS. Statt der Weiterleitung nutzt du den Geräte-Code-Login von Microsoft
(wie bei Smart-TV-Apps). Der Node **OneNote Login Helper** erzeugt dabei einmalig ein Refresh-Token; das Credential
**OneNote Device Login API** erneuert es danach selbständig.

**Azure einmalig anpassen** (App-Registrierung → links **Authentifizierung**). Die Seite gibt es in zwei Varianten:

*Neue Oberfläche (Reiter oben auf der Seite):*

1. Reiter **Umleitungs-URI-Konfiguration** → **Plattform hinzufügen** → **Mobil- und Desktopanwendungen** →
   `https://login.microsoftonline.com/common/oauth2/nativeclient` ankreuzen → **Konfigurieren**.
2. Reiter **Einstellungen** → nach unten scrollen zu **Öffentliche Clientflows zulassen** → **Ja** → **Speichern**.

*Alte Oberfläche (eine lange Seite ohne Reiter):*

1. Unter *Plattformkonfigurationen* → **Plattform hinzufügen** → **Mobil- und Desktopanwendungen** → nativeclient-URI
   (siehe oben) ankreuzen → **Konfigurieren**.
2. Ganz nach unten scrollen zu **Erweiterte Einstellungen** → **Öffentliche Clientflows zulassen** → **Ja** → **Speichern**.

*Schalter nicht zu finden?* Links **Manifest** öffnen, `allowPublicClient` suchen, auf `true` setzen (bei älteren Manifesten
heißt das Feld `isFallbackPublicClient`) und **Speichern**. Die Bezeichnungen weichen je nach Sprache und Portal-Version leicht ab.

Ohne die Mobil-/Desktop-Plattform kommt bei privaten Konten `AADSTS70002 … must be marked as 'mobile'`.

Ein Client Secret und eine funktionierende Redirect-URI werden für den Device-Login nicht benötigt. Berechtigungen wie oben
(`Notes.ReadWrite`, `offline_access`, `User.Read`, optional `Files.ReadWrite`). Nach dem Speichern ein paar Minuten warten.

**Login durchführen:**

1. Neuen Workflow anlegen, Node **OneNote Login Helper** einfügen (oder `examples/06-device-login.json` importieren).
2. *Operation:* **Start Login**, *Client ID* eintragen, ausführen. Ausgabe: `verificationUri` (`https://microsoft.com/devicelogin`) und `userCode`.
3. Die Adresse im Browser öffnen, den Code eingeben, mit dem privaten Konto anmelden und zustimmen.
4. Im selben Node *Operation* auf **Finish Login** stellen (Device Code kommt automatisch aus Schritt 2) und ausführen.
   Ausgabe: `refreshToken`.
5. Credential **OneNote Device Login API** anlegen: *Client ID*, *Refresh Token* einfügen, speichern (der Test lädt ein Notizbuch).
6. In den OneNote-Nodes bei *Authentication* **Device Login (No Redirect URL)** wählen und das Credential auswählen.
7. Die Ausführung des Login-Workflows löschen – das Refresh-Token ist ein Geheimnis.

Das Refresh-Token verfällt bei Microsoft nach 90 Tagen ohne Nutzung. n8n speichert bei jeder Erneuerung das neue Token
im Credential; sobald der Workflow gelegentlich läuft, bleibt der Login bestehen. Nach längerer Pause den Login wiederholen.

### Fehlerbehebung

| Meldung | Ursache / Lösung |
| --- | --- |
| `AADSTS50011` Redirect-URI stimmt nicht | URI in Azure exakt wie in n8n angezeigt (OAuth Redirect URL) eintragen; bei IP-Adresse: Device-Login nutzen (s. o.) |
| `AADSTS70002` … must be marked as 'mobile' | Plattform *Mobil- und Desktopanwendungen* (nativeclient-URI) hinzufügen **und** *Öffentliche Clientflows zulassen* = Ja |
| `invalid_client` beim Device-Login | Wie oben: Mobil-/Desktop-Plattform und öffentliche Clientflows in Azure aktivieren |
| Finish Login: *Not confirmed yet* | Code im Browser noch nicht bestätigt: Adresse öffnen, Code eingeben, dann *Finish Login* erneut ausführen |
| `AADSTS700016` / `unauthorized_client` | Kontotyp der App passt nicht zum gewählten *Account Type* |
| `AADSTS7000218` | Client-Secret falsch oder abgelaufen (Wert, nicht die Secret-ID kopieren) |
| `10008` „items in the folder exceed limit" | OneNote-Notizbuch in OneDrive hat zu viele Elemente; Notizbuch aufteilen |
| Video-Upload: `accessDenied` | Credential mit aktiviertem *Enable Video Upload* neu verbinden |
| Secret läuft ab | Clientschlüssel haben maximal 24 Monate Laufzeit, danach neu erzeugen |

## English

1. <https://entra.microsoft.com> → **App registrations** → **New registration**, sign in with your personal account.
2. Supported account types: *Personal Microsoft accounts only* (or *any org directory + personal*).
3. Redirect URI (platform **Web**): `http://localhost:5678/rest/oauth2-credential/callback`
   (`<n8n-url>/rest/oauth2-credential/callback` otherwise; HTTPS required except for localhost).
4. Copy the **Application (client) ID**; create a **client secret** and copy its **value** immediately.
5. Delegated Graph permissions: `Notes.ReadWrite`, `offline_access`, `User.Read`, optionally `Files.ReadWrite` (video upload).
6. In n8n create the credential **OneNote OAuth2 API**, enter client ID/secret and click **Connect my account**.

### n8n not on `localhost` (server, LAN IP, no HTTPS)

Azure only allows `http://` for `localhost`; everything else needs HTTPS with a hostname (no IP address). n8n builds the
redirect URL from `N8N_EDITOR_BASE_URL`; it cannot be changed in a credential or node. Options:

- **Device login (recommended):** no redirect URI at all. In Azure (*Authentication*) add the platform *Mobile and desktop applications* (nativeclient URI) and set *Allow public client
  flows = Yes* (new portal: tab *Settings*; old portal: *Advanced settings* at the bottom; or `allowPublicClient: true` in the
  manifest). In n8n run the **OneNote Login Helper** node (*Start Login*, open the address, enter the code,
  *Finish Login*), put the refresh token and client ID into the credential **OneNote Device Login API**, and select
  *Authentication: Device Login* in the OneNote nodes. Delete the login execution afterwards.
- **SSH tunnel for the first login:** `ssh -L 5678:localhost:5678 user@server`, set `N8N_EDITOR_BASE_URL=http://localhost:5678`
  and `WEBHOOK_URL=http://localhost:5678/`, recreate the container (`docker compose up -d --force-recreate n8n`), open
  `http://localhost:5678`, register the shown redirect URL in Azure, click *Connect my account*.
- **HTTPS with a hostname:** set `N8N_EDITOR_BASE_URL=https://n8n.example.org` and register
  `https://n8n.example.org/rest/oauth2-credential/callback`.
