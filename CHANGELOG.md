# Changelog

## 0.2.0

- New authentication option **Device Login** (credential *OneNote Device Login API* + node *OneNote Login Helper*):
  works without a redirect URI, so it can be used with n8n on a LAN IP or without HTTPS.
- Docs: setup guide covers SSH tunnel, HTTPS and device login.

## 0.1.0

- Initial release: OneNote node (notebooks, sections, pages, media), OneNote polling trigger,
  OAuth2 credential for personal Microsoft accounts.
