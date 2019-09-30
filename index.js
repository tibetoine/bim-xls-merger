const Excel = require('exceljs');

var workbookMatrice = new Excel.Workbook();
workbookMatrice.xlsx.readFile('./resources/matrice.xlsx')
  .then(function() {
    var workbookAbyl = new Excel.Workbook();
    workbookAbyl.xlsx.readFile('./resources/abyl.xlsx')
    .then(function() {
        var worksheetMatrice = workbookMatrice.getWorksheet('C');
        worksheetMatrice.eachRow(function(rowMatrice, rowNumber) {
            var codeLogement = getCodeLogement(rowMatrice.getCell(5).value)
            // console.log (codeLogement + ' : ' + isCodeLogement(codeLogement))
            if (isCodeLogement(codeLogement)) {
                var isContinue = false; // Passe à true si on doit passer à la ligne suivante.
                // 1/ Je récupère le code pièce Abyl (ou les codes pièces Abyl) correspondant au code pièce Matrice
                var pieceMatrice = rowMatrice.getCell(4).value
                if (!pieceMatrice) {
                  // Je log l'erreur et je passe à la ligne suivante  
                  console.error("Pas de pièce dans le fichier matrice pour la ligne : " + rowNumber)
                  isContinue = true
                }
                var pieceAbylArray = getPieceAbylAPartirDePieceMatrice(pieceMatrice)
                
                if (!pieceAbylArray) {
                  // Je log l'erreur et je passe à la ligne suivante  
                  console.error("Pas de pièce Abyl correspondant à " + pieceMatrice + " voir la ligne dans fichier matrice: " + rowNumber)
                  isContinue = true
                }
                // 2/ Je récupère le code composant Abyl (ou les codes composants Abyl) correspondant au code composant Matrice
                var composantMatrice = rowMatrice.getCell(3).value
                if (!composantMatrice) {
                  // Je log l'erreur et je passe à la ligne suivante  
                  console.error("Pas de composant dans le fichier matrice pour la ligne : " + rowNumber)
                  isContinue = true
                }
                var composantAbyl = getComposantAbylAPartirDeComposantMatrice(composantMatrice)
                if (!composantAbyl) {
                  // Je log l'erreur et je passe à la ligne suivante  
                  console.error("Pas de composant Abyl correspondant à " + composantMatrice + " voir la ligne dans fichier matrice: " + rowNumber)
                  isContinue = true
              }
              console.log("Ligne %s : Code logement %s, piece Matrice : %s , pièce Abyl : %s || Composant matrice : %s , composant Abyl : %s ", rowNumber, codeLogement, pieceMatrice, pieceAbylArray, composantMatrice, composantAbyl)
              if (!isContinue) {
                // Ici je fais la recherche des données dans Abyl.
                var worksheetAbyl = workbookAbyl.getWorksheet('A');
                worksheetAbyl.eachRow(function (rowAbyl, rowNumberAbyl) {
                  if (rowAbyl.getCell(3).value === codeLogement) {                    
                    console.log(pieceAbylArray.includes(rowAbyl.getCell(5).value))
                    if (pieceAbylArray.includes(rowAbyl.getCell(5).value) && rowAbyl.getCell(7).value) {
                      console.log("Yata ! ")
                    }
                  }
                })
              }
            } 
            // console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
        });
        
        // console.log(row.getCell(1).value)

    });
  });

  /**
   * Transforme une chaine en code logement.
   * 
   */
  function getCodeLogement(codeFromMatrice){
    // 0112B 103 101 1023
    //0112B1031011023
    
    if (codeFromMatrice){
        // J'enleve les blancs
        codeFromMatrice = codeFromMatrice.replace(/ /g,'')
        // Je crée le code logement
        codeFromMatrice =codeFromMatrice.substring(0, 4) + codeFromMatrice.substring(11, 15) ;     
        return codeFromMatrice
    } else {
        return null
    }    
  }

  /**
   * Return true si c'est un code logement
   * @param {} code 
   */
  function isCodeLogement(code){
    if (!code) return false    
    if (code.length === 8) {
        return true
    } 
    return false
  }


  /**
   * Retourne un code Piece Abyl à partir d'un code piece Matrice
   * Se base sur une table de correspondance (Voir doc pour savoir comment la faire : README.md)
   * @param {*} pieceMatrice 
   */
  function getPieceAbylAPartirDePieceMatrice(pieceMatrice){
    var jsonMap = {
      "CEL":["Cellier"],
      "CH1":["Chambre 1"],
      "CH2":["Chambre 2"],
      "CH3":["Chambre 3"],
      "CH4":["Chambre 4"],
      "CUI":["Cuisine_901"],
      "DGT":["Degagement","Degagement 2","Dégagement 2"],
      "ENT":["Entrée_942"],
      "ESC":["Escalier"],
      "PLA":["Placard 1 CH1","Placard 1 CH2","Placard 1 CH3","Placard 1 Dégag. 1","Placard 1 Dégag. 2","Placard 1 Entrée","Placard 1 Entrée Dégagement","Placard 1 Pièce Principale","Placard 1 SBD","Placard 1 Séjour","Placard 2 CH2","Placard 2 Entrée"],
      "SDB":["Bain et WC","Salle de bain","Salle d'eau et WC"],
      "SEJ":["Sejour"],
      "WC":["WC"]
      
    }  
    var pieceAbylArray = jsonMap[pieceMatrice]
    
    return pieceAbylArray
  }

   /**
   * Retourne un code Composant Abyl à partir d'un code composant Matrice
   * Se base sur une table de correspondance (Voir doc pour savoir comment la faire : README.md)
   * @param {*} composantMatrice 
   */
  function getComposantAbylAPartirDeComposantMatrice(composantMatrice){
    var jsonMap = {
      "Bac à douche (0.80 x 0.80 m)":"Receveur de douche",
      "Baignoire (1.61 x 0.70 m)":"Baignoire",
      "Evier":"Evier",
      "Fenetre_1V (0.90 x 1.50m) VR":"Fenêtre",
      "Lavabo":"Lavabo",
      "Meuble sous evier":"Meuble sous évier",
      "Meuble sous lavabo":"Meuble sous évier",
      "Placard 2 panneaux (largeur totale 0.65m)":"",
      "Placard 2 panneaux (largeur totale 0.80m)":"",
      "Placard 2 panneaux (largeur totale 1.00m)":"",
      "Porte d'entree logement (0.90m x 2.15m)":"Porte d'entrée",
      "Porte Fenetre_1V (0.95 x 2.10m) VR":"Porte-fenêtre",
      "Porte Fenetre_3V (2.85 x 2.10m) VR":"Porte-fenêtre",
      "Porte Fenetre_4V (3.80 x 2.10m) VR":"Porte-fenêtre",
      "Portes interieures (0.40m x 2.10m)":"Porte",
      "Portes interieures (0.73m x 2.04m)":"Porte",
      "Portes interieures (0.83m x 2.10m)":"Porte",
      "Portes interieures doubles (0.70/0.32m x 2.04m)":"Porte",
      "Portes interieures doubles (0.83/0.32m x 2.10m)":"Porte",
      "Programateur thermostat":"Thermostat d'ambiance",
      "Radiateur":"Radiateur",
      "Robinet":"Robinetterie",
      "Robinet douche":"Robinetterie",
      "Robinet thermostatique":"Robinetterie",
      "Robinetterie sanitaire":"Robinetterie",
      "Tableau electrique logement":"Tableau électrique",
      "Visiophone/interphone logement":"Interphone",
      "WC standard":"Cuvette wc"
      

    }  
    var composantAbyl = jsonMap[composantMatrice]
    
    return composantAbyl
  }