# The Vandal Stickerprogramma – Mac bouwen via GitHub

Dit pakket bevat de Mac-versie 7 en een GitHub Actions-workflow. Je hebt zelf geen Mac nodig.

## Stap voor stap op Windows

1. Pak deze ZIP uit.
2. Ga naar https://github.com/new en maak een **Private** repository aan, bijvoorbeeld `the-vandal-stickers`.
3. Upload via **Add file > Upload files** de bestanden `sticker_app.py`, `producten.xlsx`, `requirements.txt` en `README_GITHUB.md`.
4. **Belangrijk:** GitHub's webupload van losse bestanden bewaart verborgen mappen niet vanzelf. Maak op GitHub de workflow aan via **Add file > Create new file** en vul als bestandsnaam exact in: `.github/workflows/build-macos.yml`. Kopieer de volledige inhoud uit de meegeleverde `.github/workflows/build-macos.yml` en klik **Commit changes**.
5. Open het tabblad **Actions**. Selecteer links **Bouw The Vandal Mac-app** en klik **Run workflow** (branch `main`).
6. Wacht tot de twee builds klaar zijn. Open de voltooide workflow-run en download onder **Artifacts** de Intel- of AppleSilicon-versie.
7. Pak de gedownloade GitHub-artifact ZIP uit. Daarin zit nog een ZIP met de eigenlijke `.app`. Pak ook die uit op de Mac en verplaats de `.app` naar Programma's.

## Welke versie?
- **AppleSilicon**: Mac met M1, M2, M3, M4 of nieuwer.
- **Intel**: oudere Mac met Intel-processor.

## Belangrijk
- De app is niet met een Apple Developer-certificaat ondertekend of genotariseerd. macOS kan de eerste start blokkeren. Open dan **Systeeminstellingen > Privacy en beveiliging > Open toch** na een eerste startpoging, mits je de herkomst vertrouwt.
- De app print via het netwerk naar Zebra ZT230, standaard `10.60.1.231:9100`; de Mac moet het adres kunnen bereiken.
- Printerinstellingen worden op de Mac opgeslagen in `~/Library/Application Support/TheVandalStickerprogramma/printerinstellingen.json`.
- De meegeleverde `producten.xlsx` wordt bij het bouwen ingebed. Andere Excel-bestanden kun je vanuit de app openen. Wijzigingen in de ingebedde lijst vereisen een nieuwe build.
- Dit is een testbuild: code is op syntax gecontroleerd, maar de macOS-app en printerverbinding zijn niet op een echte Mac getest.
- De GitHub-runner en de beschikbare macOS-versies kunnen door GitHub gewijzigd worden; controleer de Actions-log als een build mislukt.
