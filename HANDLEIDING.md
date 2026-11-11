# Euro Superstore Boekhouding — Windows-app

Een eenvoudige boekhouding-app voor Windows 10 en 11. Facturen en Z-tickets
scannen (foto of PDF), je bestaande Excel importeren, verkoop inboeken,
overzicht en BTW-aangifte, en export naar Excel.

---

## De app op Windows zetten — in 3 stappen

Je bouwt de app één keer tot een installer. Daarna kan je papa hem
gewoon installeren en met een dubbelklik openen, net als elk programma.

### Stap 1 — Node.js installeren (één keer)
Ga naar **https://nodejs.org** en klik op de grote groene **LTS**-knop.
Installeer het (gewoon op "Volgende" blijven klikken).

### Stap 2 — De app bouwen
Dubbelklik op **`build.bat`** in deze map.
Een zwart venster opent en doet automatisch alles. De eerste keer duurt
dit een paar minuten (het haalt de onderdelen op via internet).

### Stap 3 — Installeren
Als het klaar is, opent de map **`dist`** vanzelf. Daar staan twee bestanden:

- **`Euro Superstore Boekhouding Setup 1.2.0.exe`** — dubbelklik om te
  installeren. Maakt een snelkoppeling op het bureaublad en in het startmenu.
- **`Boekhouding-portable.exe`** — start de app zonder te installeren
  (handig op een USB-stick).

Kopieer één van die twee naar de computer van je papa en klaar.

---

## Eerste keer opstarten

1. De app opent op het scherm **Instellingen**.
2. Plak daar de **Anthropic API-sleutel** (zie hieronder) en klik **Bewaren**.
3. Klaar — vanaf nu kan je facturen laten lezen.

> De app werkt **zonder internet** voor het inboeken en het overzicht.
> Internet is alleen nodig om een factuurfoto automatisch te laten lezen.

### Een API-sleutel aanmaken
1. Ga naar **https://console.anthropic.com**
2. Maak een account (of log in).
3. Ga naar **API Keys** → **Create Key**, kopieer de sleutel (begint met `sk-ant-...`).
4. Plak hem in de app bij Instellingen.

Het lezen van facturen verbruikt een klein beetje tegoed per factuur.

---

## Je bestaande Excel overzetten

Ga naar **Excel importeren**, kies je bestand en klik op **Inlezen**. Je ziet
eerst een overzicht van wat er gevonden is; pas daarna wordt er iets toegevoegd.

Wat er ingelezen wordt:

| Bestand | Tabblad | Wordt |
|---|---|---|
| SALE PURCHASE | `CASH REGISTER` | verkoop per dag, met de verdeling over 6%, 21% en 0% |
| SALE PURCHASE | `PURCHASE` | aankoop goederen |
| SALE PURCHASE | `EXPENSES` | kosten |
| Kasboek | de maandtabs | verkoop per dag, alleen het totaal |

De BTW-groepen volgen het bestand: **groep 1 = 6%, groep 2 = 21%, groep 3 = 0%**.
De BTW wordt overgenomen zoals ze in de Excel staat, niet herberekend, zodat
elk factuurtotaal exact gelijk blijft.

Niet ingelezen:

- **De uitgavenkant van het kasboek.** Die facturen staan al in PURCHASE en
  EXPENSES; opnieuw inlezen zou ze dubbel zetten.
- **Overboekingen naar de bank** (BC/MC, KBC). Dat is geen aankoop of kost.

Importeer je hetzelfde bestand twee keer, dan herkent de app de regels die er
al staan en slaat ze over. Een creditnota met hetzelfde bedrag als de factuur
blijft wel gewoon staan: daarvoor wordt ook de naam en de betaalwijze vergeleken.

### Een let op bij de winst

De app rekent de winst **exclusief BTW** en **na korting**, zoals het hoort.
Je HOME PAGE in Excel rekent met bedragen inclusief BTW en telt de kortingen
weer bij de aankopen op. Die twee cijfers zullen dus verschillen; dat is geen
fout. De BTW-bedragen komen wel exact overeen met `BTW ON SALE` en `BTW ON BUY`
op je HOME PAGE.

## Verkoop inboeken — drie manieren

Bij **Verkoop (Z)** kan je kiezen:

1. **Excel uploaden.** Een dagexport uit de kassa, je kasboek, of je SALE
   PURCHASE. De app toont eerst wat ze gevonden heeft en voegt pas toe als je
   bevestigt. Gaat het om één dag, dan kan je de cijfers ook eerst in het
   formulier laten zetten om ze na te kijken.
2. **Het kassaticket scannen.** Een foto of PDF; de app haalt de datum en de
   bedragen per departement eruit.
3. **Zelf intikken.** Alleen datum en totaal, of opgesplitst per BTW-groep.

Voor de Excel hoeft het geen vaste opmaak te zijn. De app zoekt zelf een
koprij met een kolom voor de datum en een kolom voor het bedrag. Herkend
worden onder meer `Datum`, `Date`, `Totaal`, `Bedrag`, `Amount`, en voor de
groepen `6%` / `21%` / `0%`, `DPT 1` / `DPT 2` / `DPT 3` of `Groep 1` / `2` / `3`.
De bedragen per groep worden gelezen als **inclusief BTW**, net als op het ticket.

Is er geen kolom per groep, dan wordt alleen het dagtotaal ingelezen. Zulke
dagen krijgen een oranje stip in de lijst, zodat je ziet waar de BTW-verdeling
nog ontbreekt; op het overzicht staat hoeveel dat er zijn.

## Waar staan de gegevens?

Alles wordt bewaard in één bestand op de computer van je papa:
`boekhouding.json` in de Windows-app-map. Via **Instellingen → Open data-map**
kom je er direct bij. Maak gerust af en toe een kopie als back-up.

---

## Iets aanpassen?

- De volledige app zit in **`renderer/index.html`** (uiterlijk + logica).
- Het venster en de opslag zitten in **`main.js`**.
- Wil je een nieuwe versie verspreiden? Verhoog `version` in `package.json`
  en bouw opnieuw met `build.bat`.
