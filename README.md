# Horizont LHB d.o.o. – servisi

Web stranica servisa računala i igraćih konzola. Slike, video i PDF dokumenti učitavaju se izravno s Google Drivea.

## 1. Drive mape
Napravi tri mape: **Slike**, **Video** i **PDF**. Svaku podijeli kao **Svi s poveznicom → Preglednik**.

- Podmape postaju tabovi na stranici (npr. *Laptopi*, *Konzole*, *Prije i poslije*).
- Datoteke izravno u glavnoj mapi idu pod tab **Ostalo**.
- Novi sadržaj vidi se odmah nakon osvježavanja stranice.
- Prazna mapa = ta sekcija se ne prikazuje.

ID mape je dio linka iza `/folders/`. Upiši ga u `index.html`:

```js
const FOLDERS = {
  gallery: 'ID_MAPE_SLIKE',
  video:   'ID_MAPE_VIDEO',
  pdf:     'ID_MAPE_PDF'
};
```

## 2. API ključ
Koristi se isti Drive API ključ kao na Koncept Zdravlje. Zaključan je na `https://sbatinac.github.io/*`, pa radi i ovdje bez dodatnih postavki.

## 3. GitHub Pages
1. Repo `horizont-servisi` → Settings → Pages → Branch: `main` / root → Save.
2. Stranica: `https://sbatinac.github.io/horizont-servisi/`

## 4. Prijava za servis
Gumbi vode na Google formu *Servis računala i igraćih konzola – prijava za termin*.
