const Excel = require('exceljs');

var workbookMatrice = new Excel.Workbook();
workbookMatrice.xlsx.readFile('./resources/matrice.xlsx')
  .then(function() {
    var workbookAbyl = new Excel.Workbook();
    workbookAbyl.xlsx.readFile('./resources/abyl.xlsx')
    .then(function() {
        var worksheetMatrice = workbookMatrice.getWorksheet('C');
        worksheetMatrice.eachRow(function(row, rowNumber) {
            var codeLogement = getCodeLogement(row.getCell(5).value)
            // console.log (codeLogement + ' : ' + isCodeLogement(codeLogement))
            if (isCodeLogement(codeLogement)) {
                // 1/ Je récupère le code pièce Abyl (ou les codes pièces Abyl) correspondant au code pièce Matrice
                var pieceAbylArray = getPieceAbylAPartirDePieceMatrice('SDB')

                // 2/ Je récupère le code composant Abyl (ou les codes composants Abyl) correspondant au code composant Matrice

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
                    "SDB":["Bain et WC","Salle de bain","Salle d'eau et WC"],
                    "CUI":["Cuisine_901"]
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
                    "SDB":["Bain et WC","Salle de bain","Salle d'eau et WC"],
                    "CUI":["Cuisine_901"]
    }  
    var composantsAbylArray = jsonMap[composantMatrice]
    
    return composantsAbylArray
  }