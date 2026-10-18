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

## 4. Obrazac za prijavu (Google tablica + mail)
Obrazac je dio stranice (sekcija *Servisni nalog*). Prijave šalje u Google tablicu preko Apps Scripta, a tebi stiže mail za svaku prijavu.

1. Na računu na kojem želiš prijave napravi novu **Google tablicu**, npr. *Horizont – prijave servisa*.
2. U tablici: **Proširenja → Apps Script**.
3. Obriši sve u `Code.gs` i zalijepi sadržaj datoteke `apps-script/Code.gs` iz ovog repoa.
4. U kodu upiši svoju adresu: `const NOTIFY_EMAIL = 'tvoj@mail.com';` → **Spremi**.
5. **Deploy → New deployment** → ikona zupčanika → **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
   → **Deploy** → dopusti pristup (Advanced → Go to … → Allow).
6. Kopiraj **Web app URL** (završava na `/exec`) i upiši ga u `index.html`:
   ```js
   const ENDPOINT = 'https://script.google.com/macros/s/…/exec';
   ```

Stupci u tablici nastaju sami pri prvoj prijavi. Ako kasnije mijenjaš kod skripte: **Deploy → Manage deployments → Edit → Version: New version**, inače ostaje stara verzija.

## 5. Admin stranica
`https://sbatinac.github.io/horizont-servisi/admin.html`, zaključana lozinkom.
Prikazuje posjete (danas, 7 dana, ukupno, graf 14 dana), sve servisne naloge i upite iz tablice. Klik na broj zove, klik na mail otvara poruku.

Lozinku upiši **samo u Apps Scriptu** (`const ADMIN_KEY = '...'`), nikad u ovaj repo, jer je repo javan. Nakon promjene: Implementiraj → Upravljanje implementacijama → ✏️ → Nova verzija.
