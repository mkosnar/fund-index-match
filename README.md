# fund-index-match

Zad�n�

1. M�me excel s 4 listy. Prvn� t�i listy jsou v�dy o vztahu pod�lov�ho fondu (v �ase a v CZK) k burzovn�mu indexu (v �ase a USD) a posledn� list je kurz USD/CZK v �ase
2. C�lem je sp�rovat ceny podle datum� a to tak, aby se do sloupce C zapisovaly hodnoty index� (sloupec G) podle odpov�daj�c�ho datumu
	1. Situace, �e neexistuje odpov�daj�c� hodnota indexu (proto�e se ten den v USA neobchodovalo) - pak se mus� �lut� podbarvit pole datumu ve sloupci A
	 do sloupce C se nezap�e ��dn� hodnota ale naopak do sloupce D odpov�daj�c�ho ��dku se zap�e hodnota z p�edchoz�ho dne hodnoty indexu
		1. Pakli�e nastane situace 2.1, ale jedn� se o del�� obdob� neobchodov�n� resp. chyb�j�c�ch dat indexu, je nutn� po tuto �asovou mezeru
		 do sloupce D zapisovat posledn� dostupnou hodnotu indexu, do doby, ne� budou k dispozici op�t odpov�daj�c� hodnoty indexu podle datumu
	2. Situace, �e hodnota indexu je k dan�mu datu k dispozici, ale data pod�lov�ho fondu nikoliv. V takov�m p��pad� dojde k p�esko�en� zapisov�n� hodnoty - hodnota indexu se zahod�.
3. Na �tvrt�m list� jsou hodnoty kurzu CZK/USD v �ase. Na v�ech t�ech listech  je pot�eba do sloupce E p�epo��tat dolarovou hodnotu indexu ze sloupc� C nebo D do CZK a to podle kurzu,
 kter� byl v ten konkr�tn� den platn�.
4. Pakli�e se nach�z� n�kde ne��seln� hodnota a je o�ek�v�n� ��seln�, zap�e program tuto skute�nost na nov� list (list m��e b�t vytvo�en u� na za��tku zpracov�n� dat a m��e z�stat pr�zdn�).
 Na tento list zap�e na jak�ch listech a jak�ch bun�k se probl�m t�ka.
5. Pokud nastane jin� zde nepopsan� chyba, bude na listu s ostatn�my n�lezy, ale nebude ur�eno, o jakou chybu se jedn�
 (jen jak�ch list� a bun�k se t�k� a pole ur�uj�c� porblematick� bunky a listy bude podbarven� �erven�)

Pozn�mka - kl��ov� je kdy se obchodoval fond, ne index.