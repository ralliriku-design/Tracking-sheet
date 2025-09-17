# Shipment Tracking Toolkit — Päivityspaketti (Ralli Logistics, 2025-09-16_2336)

Tämä paketti täydentää aiemmin toimittamaani Apps Script -kokonaisuutta **Kaukokiito-seurannalla** ja pienillä korjauksilla.
Saat mukana myös päivitetyn CredentialsHub.html -paneelin sekä kevyen testiskriptin.

## Mitä uutta
- **Kaukokiito-tuki**: `TRK_trackKaukokiito()` + valikkotoiminto *Shipment → Päivitä STATUS → Kaukokiito*.
- **Carrier-normalisointi**: `canonicalCarrier_()` tunnistaa nyt `kaukokiito`, `kki`.
- **Päiväyksen tulkinta**: `parseDateFlexible_()` osaa varmasti `dd.MM.yyyy HH:mm` ja ISO -formaatit.
- **Script Properties**: `seedKnownAccountsAndKeys()` lisää `KAUKO_TRACK_URL` placeholderin; `readMissingProps_()` näyttää puuttuvat KAUKO-avaimet.
- **HTML-paneeli**: Päivitetty ohje KAUKO-avaimista.

## Tiedostot
1. **01_patch_kaukokiito_and_fixes.gs** — Pudota projektiin *uutena .gs-tiedostona*. Tämä sisältää **korvaavat** funktiot (katso tiedoston otsikkokommentti).
2. **02_CredentialsHub.html** — Luo Apps Scriptiin uusi HTML-tiedosto nimellä `CredentialsHub` ja korvaa sisällöllä.
3. **04_TestScript.gs** — Valinnaiset itse-testit ja demo-data (ei kutsu ulkoisia rajapintoja).

> Jos projektissasi on jo `CredentialsHub`-tiedosto ja apufunktiot (DELIVERED_KEYWORDS jne.), **älä duplikoi**. Tämä paketti ei sisällä DELIVERED_KEYWORDS:ia, olettaen että se on jo lisätty aiemmasta täydennys-paketista.

## Asennus
1. Avaa Google Sheets → Laajennukset → Apps Script.
2. Lisää file **01_patch_kaukokiito_and_fixes.gs** (File → New → Script). **Poista** vanhat samannimiset funktiot (tai anna tämän tiedoston olla listassa viimeisenä).
3. Luo/korvaa **CredentialsHub**-HTML sisällöllä **02_CredentialsHub.html**.
4. (Valinnainen) Lisää **04_TestScript.gs** testejä varten.
5. Paina „Save” ja suorita valikosta **Shipment → Tarkistimet → Integraatioavaimet** tai avaa **Shipment → Credentials Hub**.

## Asetukset (Script Properties)
Avaa **Shipment → Credentials Hub** ja täytä avaimet rivimuodossa `KEY=VALUE`:

**Kaukokiito (valitse A tai B):**
- **A) API**
  - `KAUKO_TRACK_URL=https://api.kaukokiito.example/track?code={code}` *(vaihda oikeaan)*
  - Ja joko `KAUKO_API_KEY=...` **tai** `KAUKO_BASIC=Base64(user:pass)`
- **B) HTML fallback**
  - `KAUKO_SCRAPE_URL=https://www.kaukokiito.fi/seuranta/{code}` *(tai muu julkinen seurantasivu)*

**Muut (esimerkit; monet teillä ovat jo käytössä):**
- `MH_TRACK_URL`, `MH_BASIC`
- `POSTI_TOKEN_URL`, `POSTI_BASIC`, `POSTI_TRACK_URL`
- `GLS_FI_TRACK_URL`, `GLS_FI_API_KEY` *TAI* `GLS_TOKEN_URL`, `GLS_BASIC`, `GLS_TRACK_URL`
- `DHL_API_KEY`, `DHL_TRACK_URL`
- `BRING_TRACK_URL`, `BRING_UID`, `BRING_KEY`, `BRING_CLIENT_URL`
- `BULK_BACKOFF_MINUTES_BASE=5`

## Käyttö
- **Kertapäivitys:** `Shipment → Päivitä STATUS (valittu carrier) → Kaukokiito`
- **Bulk-ajo:** `Shipment → Iso statusajo → Aloita (aktiivinen välilehti)` — Kaukokiito huomioidaan automaattisesti kuten muutkin.
- **Vaatii_toimenpiteitä:** toimii kuten ennen. Kun Kaukokiito-rivejä päivitetään, `Delivered date (Confirmed)` täyttyy kun status tulkitaan toimitetuksi.

## Itsetestit (ilman ulkoisia kutsuja)
1. Lisää **04_TestScript.gs** projektiin.
2. Aja `seedDemoData_()` — luo pienen testitaulukon *Adhoc_Tracking*.
3. Aja `test_LocalParsers_()` — varmistaa `parseDateFlexible_` ja Kaukokiito JSON/HTML-heuristiikan toimivuuden paikallisilla sampleilla.
4. (Ulkoiset testit) Aseta oikeat KAUKO_* -avaimet ja käytä valikosta *Päivitä STATUS → Kaukokiito*.

## Huomioita
- HTML-scrape fallback on tarkoituksella **kevyt** ja voi rikkoutua, jos sivun rakenne muuttuu. Suositus on käyttää virallista APIa.
- Patch ei sisällä koko pääskriptiä. Jos haluat **täyden yhdistetyn .gs** -tiedoston (one-file), sanot vain — koostan ja annan yhtenä pakettina.
- Jos bulk-ajossa tulee 429, `Retry-After` huomioidaan ja `NextAt`-kenttä täyttyy.

Onko tarvetta myös **Power BI** -raporttien automaattiselle Kaukokiito-checkille? Voin lisätä sen `powerBiArchiveNew()`-polkuun, kuten Postille on tehty.