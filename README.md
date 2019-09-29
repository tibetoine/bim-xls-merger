# bim-xls-merger
Tool to merge Xls files into one

****** PreRequis - Matrice

* Renommage des fichiers :
- matrice.xlsx
- abyl.xlsx


* Renommage des onglets
- Dans matrice : 'Composants (Attribut)' devient 'C'
- Dans matrice : 'Expaces (finitions)' devient 'E'

(Abyl : Le fichier doit avoir été simplifié)
- Dans abyl : L'onglet 1 s'appelle 'A'


** Faire les tables de correspondances
- On utilise Excel
- On récupère la colonne Abyl.Piece , Puis on enlève les doublons, puis on trie de A à Z
- On récupère la colonne Matrice.Espace , Puis on enlève les doublons, puis on trie de A à Z
- On fait un petit tableau 
- Puis génération JSON