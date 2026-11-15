# ZSP thuisbatterij – globale landingspagina

Live via GitHub Pages. Gebouwd met `python3 build.py --live` uit het project zsp-landingspaginas; pas niets hier met de hand aan, behalve `lead-endpoint.js`.

- `index.html` – de globale landingspagina (geen varianten), met GTM-KWDXBZG7.
- `lead-endpoint.js` – adres van de Google Apps Script-webapp die de aanvragen ontvangt (eindigt op `/exec`). Leeg = de pagina neemt geen aanvragen aan.
- `google-sheet-leads.gs` – het script voor de Google Sheet "ZSP leads thuisbatterij" (installatie staat bovenaan het bestand).
