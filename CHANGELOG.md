# Changelog

## 0.2.1

- Docs and error hints: Device Login needs the Azure platform "Mobile and desktop applications" (AADSTS70002).
- Docs: complete device login steps, where to find "Allow public client flows", authentication field in usage guide.
- Package no longer ships `tsconfig.tsbuildinfo` and a copy of `package.json` in `dist/`.
- CI fixed (lock file); new GitHub Actions workflow `publish.yml` publishes with npm provenance.

## 0.2.0

- New authentication option **Device Login** (credential *OneNote Device Login API* + node *OneNote Login Helper*):
  works without a redirect URI, so it can be used with n8n on a LAN IP or without HTTPS.
- Docs: setup guide covers SSH tunnel, HTTPS and device login.

## 0.1.0

- Initial release: OneNote node (notebooks, sections, pages, media), OneNote polling trigger,
  OAuth2 credential for personal Microsoft accounts.
