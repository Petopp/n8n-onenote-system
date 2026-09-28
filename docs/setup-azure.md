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

### Fehlerbehebung

| Meldung | Ursache / Lösung |
| --- | --- |
| `AADSTS50011` Redirect-URI stimmt nicht | URI in Azure exakt wie in n8n angezeigt (OAuth Redirect URL) eintragen |
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
