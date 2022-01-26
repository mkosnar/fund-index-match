# fund-index-match

Zadání

1. Máme excel s 4 listy. První tøi listy jsou vždy o vztahu podílového fondu (v èase a v CZK) k burzovnímu indexu (v èase a USD) a poslední list je kurz USD/CZK v èase
2. Cílem je spárovat ceny podle datumù a to tak, aby se do sloupce C zapisovaly hodnoty indexù (sloupec G) podle odpovídajícího datumu
	1. Situace, že neexistuje odpovídající hodnota indexu (protože se ten den v USA neobchodovalo) - pak se musí žlutì podbarvit pole datumu ve sloupci A
	 do sloupce C se nezapíše žádná hodnota ale naopak do sloupce D odpovídajícího øádku se zapíše hodnota z pøedchozího dne hodnoty indexu
		1. Pakliže nastane situace 2.1, ale jedná se o delší období neobchodování resp. chybìjících dat indexu, je nutné po tuto èasovou mezeru
		 do sloupce D zapisovat poslední dostupnou hodnotu indexu, do doby, než budou k dispozici opìt odpovídající hodnoty indexu podle datumu
	2. Situace, že hodnota indexu je k danému datu k dispozici, ale data podílového fondu nikoliv. V takovém pøípadì dojde k pøeskoèení zapisování hodnoty - hodnota indexu se zahodí.
3. Na ètvrtém listì jsou hodnoty kurzu CZK/USD v èase. Na všech tøech listech  je potøeba do sloupce E pøepoèítat dolarovou hodnotu indexu ze sloupcù C nebo D do CZK a to podle kurzu,
 který byl v ten konkrétní den platný.
4. Pakliže se nachází nìkde neèíselná hodnota a je oèekáváná èíselná, zapíše program tuto skuteènost na nový list (list mùže být vytvoøen už na zaèátku zpracování dat a mùže zùstat prázdný).
 Na tento list zapíše na jakých listech a jakých bunìk se problém týka.
5. Pokud nastane jiná zde nepopsaná chyba, bude na listu s ostatnímy nálezy, ale nebude urèeno, o jakou chybu se jedná
 (jen jakých listù a bunìk se týká a pole urèující porblematické bunky a listy bude podbarvené èervenì)

Poznámka - klíèové je kdy se obchodoval fond, ne index.