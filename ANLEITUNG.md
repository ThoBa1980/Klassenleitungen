# Klassen-Ampel – Anleitung

So funktioniert es in drei Bausteinen:

1. **Microsoft Forms** ist die Ampel. Die Kolleg:innen öffnen sie per Link oder QR-Code auf dem Tablet und tippen dreimal: Doppelstunde, Farbe, Absenden.
2. **Excel (OneDrive)**: Forms sammelt alle Antworten automatisch in einer Tabelle.
3. **`Klassenampel_Auswertung.xlsx`** (diese Datei) rechnet die Wochen-, Monats- und Schlussabrechnung sowie die Zeitmuster aus.

Punkte: 🟢 Grün = +1, 🟡 Gelb = 0, 🔴 Rot = −1.

---

## Teil A – Das Formular anlegen (einmalig, ca. 10 Minuten, am PC)

1. Unter <https://forms.office.com> mit dem **Schulkonto** anmelden.
2. **„Neues Formular“** wählen. Titel zum Beispiel: `Ampel 7b`.
   Beschreibung: `Bitte am Ende jeder Doppelstunde ausfüllen. Verdrückt? Einfach noch einmal senden – der letzte Eintrag zählt.`
3. **Frage 1** über „Neue hinzufügen“ → **Auswahl**:
   - Fragetext genau: **`Doppelstunde`**
   - Optionen: `1./2. Stunde`, `3./4. Stunde`, `5./6. Stunde`
   - Schalter **„Erforderlich“** einschalten
4. **Frage 2** → **Auswahl**:
   - Fragetext genau: **`Ampel`**
   - Optionen: `🟢 Grün – lief gut`, `🟡 Gelb – okay`, `🔴 Rot – lief schlecht`
     (Die Emojis sind freiwillig. Wichtig ist nur, dass die Wörter **Grün**, **Gelb** und **Rot** vorkommen.)
   - **„Erforderlich“** einschalten
5. **Frage 3** → **Datum**:
   - Fragetext genau: **`Anderes Datum (nur ausfüllen, wenn nicht heute)`**
   - **Nicht** erforderlich. Normalerweise bleibt das Feld leer, und es zählt automatisch der Tag des Absendens.
6. **Einstellungen** (über das Zahnrad oder „…“ → Einstellungen):
   - ✅ **„Nur Personen in meiner Organisation können antworten“**
   - ⬜ **„Namen aufzeichnen“** ausschalten, damit die Bewertung anonym bleibt (siehe Datenschutz unten)
   - ⬜ **„Auf eine Antwort pro Person beschränken“** muss **AUS** sein, denn jede Lehrkraft bewertet ja viele Stunden.

> ⚠️ Die Fragetitel **`Doppelstunde`**, **`Ampel`** und **`Anderes Datum…`** müssen genau so beginnen, denn die Excel-Auswertung erkennt die Spalten an diesen Wörtern.

## Teil B – Den Zugang auf die Tablets bringen

1. Im Formular auf **„Antworten sammeln“** bzw. **„Teilen“** klicken.
2. **Link kopieren** und den Kolleg:innen schicken (Teams, Mail). Zusätzlich gibt es dort einen **QR-Code** zum Herunterladen: ausdrucken und im Klassenraum neben die Tafel oder an die Tür hängen.
3. **Tipp für „ohne viel Klicken“:** Auf dem Tablet den Link einmal öffnen und als Symbol auf den Startbildschirm legen:
   - **iPad (Safari):** Teilen-Symbol → „Zum Home-Bildschirm“
   - **Android (Chrome):** ⋮ → „Zum Startbildschirm hinzufügen“

   Danach reichen ein Tipp aufs Symbol und drei Tipps im Formular.

## Teil C – Auswertung (z. B. freitags, am PC)

1. In Forms auf **„Antworten“** → **„Ergebnisse in Excel öffnen“** klicken. Die Tabelle liegt in deinem OneDrive und wird automatisch aktualisiert.
2. Dort alles markieren (**Strg + A**) und kopieren (**Strg + C**).
3. `Klassenampel_Auswertung.xlsx` öffnen → Blatt **„Rohdaten“** → alten Inhalt löschen → Zelle **A1** anklicken → einfügen (**Strg + V**).
4. Fertig. Die Ergebnisse stehen hier:
   - **Übersicht**: Schlussabrechnung mit Punktestand, Grün/Gelb/Rot-Anteil, bester und schwächster Woche, bestem und schwächstem Tag, bester und schwächster Tageszeit
   - **Wochen**: Wochenabrechnung mit Diagrammen
   - **Monate**: Monatsabrechnung
   - **Zeitmuster**: farbige Tabelle Wochentag × Doppelstunde (grün = läuft gut, rot = läuft schlecht)
5. Beim **ersten Mal** im Blatt **„Einstellungen“** Klasse, Startdatum und **Enddatum (Halbjahresende)** eintragen. Unten im Blatt steht eine Prüfung, die „OK – Spalten gefunden“ anzeigen muss.

Zum Ausprobieren gibt es `Klassenampel_Auswertung_DEMO.xlsx` mit erfundenen Beispieldaten.

### Regeln der Auswertung
- Gezählt werden nur Montag bis Freitag zwischen Start- und Enddatum.
- Wird dieselbe Doppelstunde am selben Tag mehrfach bewertet, **zählt nur der letzte Eintrag**. Wer sich vertippt hat, sendet einfach noch einmal.
- Ferienwochen erscheinen mit 0 Bewertungen und ziehen den Durchschnitt **nicht** herunter.
- Aussagen zu einzelnen Zeitfenstern (z. B. „Freitag 5./6.“) beruhen anfangs auf wenigen Bewertungen. Erst nach mehreren Wochen sind sie aussagekräftig.

## Datenschutz (bitte kurz mit der Schulleitung abstimmen)
- Es wird **die Klasse als Ganzes** bewertet, keine einzelnen Schüler:innen. Mit „Namen aufzeichnen“ = AUS wird auch nicht gespeichert, **welche Lehrkraft** bewertet hat.
- Die Daten liegen im Microsoft-365-Konto der Schule. Ob dessen Nutzung für solche Zwecke freigegeben ist, regelt eure Schule bzw. das Bundesland.

## Quellen (Microsoft-Support)
- Wer das Formular ausfüllen darf: <https://support.microsoft.com/en-US/Forms/choose-who-can-fill-out-a-form-or-quiz>
- Namen nicht aufzeichnen: <https://support.microsoft.com/en-au/forms/set-up-your-survey-so-names-aren-t-recorded-when-collecting-responses>
- Formular teilen, Link und QR-Code: <https://support.microsoft.com/forms/send-a-form-and-collect-responses>
- Ergebnisse ansehen und in Excel öffnen: <https://support.microsoft.com/en-US/Forms/check-and-share-your-form-results>
- Automatische Synchronisierung Forms → Excel: <https://techcommunity.microsoft.com/blog/microsoftformsblog/introducing-forms-data-sync-to-excel/4036051>

---

## Textvorschlag für die Kolleg:innen

> Liebe Kolleginnen und Kollegen,
>
> bis zum Halbjahr möchten wir das Verhalten der 7b mit einer **Klassen-Ampel** begleiten. Bitte gebt am Ende eurer Doppelstunde kurz eine Rückmeldung (10 Sekunden):
> **Link/QR-Code öffnen → Doppelstunde antippen → Grün / Gelb / Rot → Absenden.**
>
> 🟢 lief gut (+1 Punkt) · 🟡 okay (0) · 🔴 lief schlecht (−1)
>
> Verdrückt? Einfach noch einmal senden, dann zählt der letzte Eintrag. Wer nachträglich bewertet, trägt bitte das Datum ein. Die Bewertung ist anonym.
>
> Link: _[hier einfügen]_
>
> Vielen Dank!
