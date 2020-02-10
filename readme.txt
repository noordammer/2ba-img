*********************************************
* Titel : 2BA afbeeldingen download         *
* Auteur: Paul Noordam				    *
* Jaar	: 2019					    *
*********************************************

Inhoud:
- inleiding
- werkwijze programma
- aanpassingen
- problemen
- backup API inlog gegevens

==========
INLEIDING:
==========

Doel van dit programma is om afbeeldingen via 2BA en aanwezige webcrawlers te verzamelen.
Hiervoor word verbinding met de API van 2ba gelegd.
Resultaten worden per leverancier opgeslagen als Excel bestanden.

====================
WERKWIJZE PROGRAMMA:
====================
--------------------
     HOME TAB
--------------------
Met de knop "Open bestand" kiest de gebruiker een Excel bestand in xlsx formaat aan met de juiste bestandsindeling.
Hierna kiest de gebruiker voor het verwerken van een enkele of meerdere leveranciers.
Hiervoor kunnen de knoppen aan de rechterkant worden gebruikt of voor het kiezen van enkele leveranciers kan de control toets 
worden gebruikt in combinatie met muisklikken.

Met de knop "Include crawler" kiest de gebruiker om het programma wel of niet een aanwezige crawler te gebruiken.

Om gegevens van 2BA op te halen moet de gebruiker verbinding maken met het platform.
Dit gebeurd door middel van de knop "2BA connectie"

Nadat de zijn gemaakt drukt de gebruiker op "Start" om de verwerking te starten.
Tijdens de verwerking verschijnen er status updates in de balk onderaan het venster.

Na afloop van de verwerking verschijnt er op het scherm een laatste update waarmee word weergegeven dat de verwerking ten einde is.
Het output bestand staat in de Logo\Leverancier foto's preload\ map op de bestandsbeheer schijf.

-------------------
    IMAGES TAB
-------------------
Tijdens het download proces van de downloadtool kan de gebruiker reeds gegenereerde outputbestanden openen en de afbeeldingen verwerken.

Met de knop "Open bestand" kiest de gebruiker een Output bestand in xlsx formaat dat is gegenereerd door de download tool.
Hierna kiest de gebruiker een artikelnummer om de beschikbare afbeelding te bekijken.
Achter elk artikelnummer staat tussen vierkante haakjes hoeveel afbeeldingen er beschikbaar zijn.
Wanneer dit 1 is zal de afbeelding automatisch worden geladen. Bij meer afbeeldingen dient de gebruiker een keuze te maken via het 
dropdown kader.
Boven het kader word het merk en het artikelnummer leverancier weergegeven.
In de statusbalk word de afbeeldingnaam en het formaat weergegeven.
(Het formaat is het reeds bijgewerkte formaat, zie uitleg hieronder)

De gebruiker kan er voor kiezen om overbodige witruimte meteen te verwijderen door middel van de Whitespace optie.
Standaard staat deze aan. (Als deze word uitgeschakeld dient de gebruiker de afbeelding opnieuw op te vragen!)

Het bekijkvenster past zich automatisch aan de opgevraagde afbeelding aan.
Afbeeldingen die worden geopend ondergaan automatisch de volgende handelingen:
- Indeling word geconverteerd naar RGB
- Indien de afbeelding groter is dan 600x600 pixels word deze teruggeschaald totdat de langste zijde 600 pixels is. (verhouding word behouden)
- De afbeelding word opgeslagen als jpg bestand met als naam het artikelnummer van de leverancier.
De meest gangbare bestandsformaten kunnen worden geopend en bewerkt.
Er zijn uitzonderingen, zie het hoofdstuk problemen.

Met de knop "Save image" word de weergegeven afbeelding opgeslagen op de gebruikelijke XXX locatie.
Met de knop "Save all" word voor elk artikel in de lijst die slechts 1 afbeelding heeft de afbeelding opgeslagen.
Artikelen met meerdere afbeeldingen worden overgeslagen.

=============
AANPASSINGEN:
=============
-------------
SETTINGS TAB
-------------
Krediteurinformatie:
Ontbrekende informatie kan worden toegevoegd aan het bestand "Kred lijst.xlsx" in de programma map.
Tevens kan er informatie worden gewijzigd of verwijderd wanneer noodzakelijk.
Per krediteur dient er per merk een regel te worden geschreven in het overzichtsbestand.
Voorbeeld:
Kred	|Rol		|Kred		|Merk		|GLN			|GLN
nummer	|			|Naam		|Naam		|Data			|Prijzen
--------|-----------|-----------|-----------|---------------|-------------|
22600	|Leverancier|HEGEMA		|HEGEMA		|8712423006195	|			  |
22610	|Leverancier|HELAF		|HELAF		|				|			  |
22612	|Leverancier|HELLERMANN	|HELLERMANN	|4031026000008	|			  |
22640	|Agent		|HEMMINK	|TRIAX		|8712251990000	|8712251990000|
22640	|Agent		|HEMMINK	|HEMPRO		|8712251990000	|8712251990000|
22640	|Agent		|HEMMINK	|ADELS		|2220000035989	|8712251990000|
22640	|Agent		|HEMMINK	|ASTRO		|2220000020022	|8712251990000|
22640	|Agent		|HEMMINK	|BEDEA		|2220000020046	|8712251990000|
22640	|Agent		|HEMMINK	|CABELCON	|2220000020152	|8712251990000|
--------|-----------|-----------|-----------|---------------|-------------|

Extensielijst:
Hier kan de gebruiker uitzonderingen invoeren wanneer een download tot vreemde resultaten leidt.
Als voorbeeeld een downloadlink van Eaton artikel EB002:
- https://pl.eaton.com/image?doc_name=EasyBatteryPlus&type=TIFDetail
Omdat het proces de laatste 4 tekens van een link als extensie gebruikt resulteert dit in bestand EB002tail.

Met de Extensielijst kan de gebruiker deze anomalies afvangen.
Input hiervoor is {tekst},.{extensie} wat in bovenstaand voorbeeld tail,.tif word.

API Oauth2 gegevens:
Met deze gegevens word verbinding gelegd met de API van 2BA.
Een aantal van deze gegevens kunnen alleen worden verstrekt door 2BA!
Zie BACKUP API INLOG GEGEVENS.
=============
PROBLEMEN(?):
=============

-Bestand word niet geladen!
 Controleer bestand op juiste extensie (xlsx) en indeling.

-Geen resultaten in output bestand!
 Controleer verbinding met 2ba. Zie "2BA koppeling" rechtsboven in scherm.

-Nonetype errors
 Er zitten lege velden in het bestand waardoor het programma stopt.
 Vaak gaat het om het ontbreken van een merknaam in de kolom Typenummer en merk.

-Afbeelding word niet geladen
 Afbeelding heeft een indeling waar het programma niet mee overweg kan.
 In de meeste gevallen heeft dit betrekking tot het aantal kleurkanalen in een afbeelding.
 
-Waarom werkt die shit niet!
 Je hebt 'm kapot gemaakt of het is te moeilijk voor je, kies maar.
 
-Andere problemen!
 Contact opnemen met iemand, bij voorkeur niet de auteur!

==========================
BACKUP API INLOG GEGEVENS:
==========================
In het geval van een verkeerde aanpassing van de API gegevens de lokale readme raadplegen voor de backup gegevens.
