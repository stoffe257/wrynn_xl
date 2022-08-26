# wrynn_xl
An excel booklet for making seating charts for big dinners. It is made in Excel, with macros done in VBA to enable the needed functionality. The spreadsheet is in Swedish together with the description and user manual. 

I wrote the code my first year of university, and figured I would place it here to enable it to be used by others. Since this is an old project I probably won't keep adding features, but feel free to add feature requests and post issues.

The rest of the README will be in Swedish.

## Inledning
Wrynn-Execls är ett bordplaceringsark som underlättar jobbet av bordsplacering. Du kan antingen använda det som ett verktyg i ditt manuella plannerande eller be AI:n ta över jobbet och placera ut gästerna. Arkets syfte är helt enkelt att underlätta för dig som plannerar en sittning,
liten som stor.  

En bra sak att nämna är att den älskade ctrl z kommandot är mer eller mindre oanvändbart i
arket. Du kan läsa mer om det under begränsningar.

## De olika flikerna
Här beskrivs de olika flikarna kort. Vad de används till och vilka du som användare ska ändra i.

### Start
Startfliken är din arbetsplats. Här skapar du bord och placerar ut gäster. Den innehåller huvud-
menyn med lite olika funktioner som kommer visa sig användbara. Den håller även koll på antalet
gäster, bord, utplacerade gäster osv.

### Gästlista
Här finns gästlistan för sittningen. Den kommer du klista in själv här, men sedan behöver du inte ändra mer i den här fliken. Vill du lägga till, ta bort eller ändra information för gäster kan detta göras i denna flik, under hela arbetets gång.

### Till köket
I den här fliken finns en alternativ vy som är anpassad för att skickas till köket. Den visar bordsplaceringen med de drycker som ska finnas vart, samt vart personerna med allergier kommer sitta. Den här även lite användbar statistik på antalet av varje dryck mm. Den här fliken behöver du aldrig ändra på själv.

### Bordsplacering
Denna flik innehåller det sittningsschema ni sätter upp utanför sittningssalarna, så alla kan se vart de ska sitta. Den här fliken visar därför bara namn på gästerna där de ska sitta, samt sittningens namn i ena hörnet. Den här fliken behöver du aldrig ändra på själv.

### Till tryck
Här finns en gästlista med namn för att skicka till tryck för bordsplaceringslapparna. Den är i ordning efter vart personerna sitter, bord för bord. Detta för att underlätta vid den fysiska utplaceringen av lapparna. Hur väl detta funkar beror på hur de som trycker hanterar listan, men i värsta fall har du en klar lista att skicka.

### Specialkost
Denna flik är lite speciell, och syns helt enkelt inte. Den används för att generera en komplett lista på all specialkost, som sedan kommer exporteras som PDF. Den här fliken behöver du inte se, och inte heller ändra på själv.

## Så används arket
Nedan följer en genomgång om hur arket är tänkt att användas, samt kort om hur funktionerna
fungerar.

### Hämta gästlista
Först och främst måste arket få tillgång till gästlistan. Denna kopierar du enklast ifrån google-kalkylarken som skapas från googleformulären och klistrar in i arket under fliken Gästlista, under de rubriker som står. Här är det viktigt att ordningen på informationen är rätt, för att arket ska förstå vad som är vad. Även E columnen kommer att användas av programmet till att hålla koll på vilka gäster som är utplacerade, så där bör ingen info lagras (programmet kommer bara ta bort den ändå isåfall). När gästlistan är inklistrad är du klar med den fliken.

### Skapa bord
För att skapa ett bord markerar du cellerna där du vill placera bordet, för att sedan trycka på "Skapa bord" (ctrl k). Ett bord skapas alltid på vertikalen, och är alltid två celler brett. Ett bord kan inte skapas för nära andra bord, huvudmenyn eller kanter på fliken, då ett feleddelande visas istället.

### Hantering av bord
När ett bord har skapats markeras det med feta svarta kanter. Det står även bordets nummer strax ovan. Detta sker förutom i "Start" fliken även i "Till köket" och "Sittningsschema", men det återkommer vi till senare. Om du gjort fel med borden kan du klicka på "Rensa" så får du börja om. OBS! "Rensa" raderar även utplacerade namn.

### Utplacering av gäster
Om du önskar att placera ut gäster manuellt (t.ex TMs) så gör detta genom att skriva namnet på gästen i en av cellerna innanför bordets kanter. Om du inte orkar skriva hela namnet så kan du skriva en bit av det, och sen trycka på ctrl j, så fylls hela namnet i automatiskt. Om du placerat ut en gäst på flera ställen kommer ett felmeddande upp om detta. När en gäst placerats ut kommer informationen om vilken dryck gästen önskar samt gruppering om sådan angets finnas på sidan av bordet. På de andra flikarna med grafiska vyer står annan information, relevanta för de vyerna.

### Byta plats på namn (ctrl l)
Om du önskar att byta plats på två gäster markerar du först den ena, sen håller inne ctrl och trycker på den andre. Efter detta trycker du på knappen "Byt plats" i huvudmenyn, eller ctrl l. Önskar du däremot att ta bort ett namn klickar du bara på namnet och sen på backspace som vanligt.

### En gäst på flera platser
Om en gäst placerats ut på mer än en plats kommer programmet att påpeka detta, och markera
det med ett felmeddelande tills du har löst detta.

### "AI"
Programmet har även en "AI" som kan hjälpa till med placeringen. Den går igenom varje bord, plats för plats och placerar en lämplig gäst där. Första prioritering är att matcha både dryck (vin mitt emot vin) samt gruppering. Sen går AI:n på gruppering, och sen bara på dryck. Finns ingen som passar in på föregående kriterier, väljer programmet första bästa gäst och placerar.

### Skapa PDF:er
När du klickar på denna knapp kommer fem stycken olika PDF:er att genereras, och du kan välja var för sig vilka du vill spara. Första är vy från Startfliken, på all dess info helt enkelt. Nästa är gästlistan. Tredje är från vyn "Till köket". Sen kommer PDF:en för sittningsschemat som ska sättas upp. Sist är en komplett lista över specialkost att skicka till köket. Värt att tänka på är att alla PDF:er autoresizar för att få plats med allt på ett A4 i landscape. Därför kan det bli plottrit om man har större sittningar med stort antal bord. Då rekommenderar jag att själv skapa en PDF som vanligt i Excel.

## Begränsningar/utvecklingsmöjligheter
Även om arket är smart i många avseenden har det även sina begränsningar. Generellt är det mesta dynamiskt (uppdateras automatiskt med ändringar). Om du gör ändringar i gästlistan eller på startfliken görs dessa ändringar i samtliga andra flikar, utan att du behöver klicka på någon form av uppdateringsknapp. I huvudmenyn håller arket även koll på hur många gäster som är utplacerade, och detta uppdateras även det i realtid. Den håller på samma sätt koll på antal gäster, antal bord samt antal platser tillgängliga totalt baserat på utsatta bord. Dock finns det som sagt några begränsningar. Dessa följer nedan.

### Stänga och öppna arket
Om du sparar och stänger ner arket, för att sedan öppna det igen kommer du inte kunna lägga till fler bord eller använda AI:n direkt. Antingen rensar du och börjar om, eller så visar du manuellt programmet vart borden finns. Efter detta kan programmet funka som vanligt.

### Bordsformen
Arket kan bara hantera bord som är avlånga samt vertikalt placerade. De måsta vara just två celler breda, och man kan inte placera personer vid kanterna. Det har mycket med Excel att göra, men även att arket i sig skulle bli mycket mer komplext, samt mer svåranvänt för användaren. Ofta är borden på sittningen avlånga så i de flesta fallen bör det inte vara några problem. Annars rekommenderar jag att du använder din fantasi för att lösa problemet.

### Rensa borden
Som du kanske funderat på så måste man börja om från början om man gjort fel med borden. Var därför noga med placeringen av borden i början. I nästan alla fall har en sittning ett maxantal, och hur borden ska stå uppställda bör vara klart innan bordplaceringen görs. Trotts detta kikas det på en lösning för att ta bort endast ett bord, och kanske kan finnas tillgänlig i en nästa version av arket det med.

### AI:n slår inte människan
En annan sak du bör ha i åtanke är att AI:n inte kan ta hänsyn till mer information är vad du givit den. Den vet inte vilka som passar bra och mindre bra att sitta bredvid varanda, utan har i den aspekten endast grupperingen att gå på. Jag rekommenderar att du kollar igenom placeringen efter AI:n körts och använder din kännedom av gästerna för att göra små mindre förbättringar.

### Ctrl z
Då kod i programmet körs kan denna inte backas tillbaks från, vilket gör att ctrl z blir mer eller mindre oanvändbart i arket. Du kan alltid testa, men oftast kommer inget hända. Detta har jag försökt jobba mig runt genom att göra det enkelt att ändra saker, långt in i det sista, men det kan vara bra att ha i åtanke för er som har tendens att använda det kortkommandot mycket. Detta är tyvärr inget som går att göra något åt utan det beror på hur Excel VBA är designat.

## När saker inte går som det ska
De flesta sakerna som kan gå fel är du skyddad för genom felmeddelanden som avbryter processen. I felmeddelandet står det tydligt förklarat varför det inte går att göra som du vill. Om det dock mot formodan skulle hända att programmer krashar så väljer du att avsluta processen i pop-up- rutan som visas. Om detta händer kommer arket fungera som om du stängt ner och öppnat det igen. Annars är alltid den förslagna lösningen att rensa arket, så kommer allt funka som det ska igen.

## Vidareutveckling
Som nämnt ovan finns det en del saker som skulle kunna lösas, för att göra programmet aningen vassare. Beroende på hur mycket programmet används kan det kanske komma en ny version. Känner du att du själv vill förbättra något är det fritt fram att göra så. Forka projektet och kör hårt!