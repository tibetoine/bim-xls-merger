const Excel = require('exceljs');

var workbookMatrice = new Excel.Workbook();
workbookMatrice.xlsx.readFile('c:/dev/resources/matrice.xlsx')
  .then(function() {
    var workbookAbyl = new Excel.Workbook();
    workbookAbyl.xlsx.readFile('c:/dev/resources/abyl.xlsx')
    .then(function() {
        var worksheetMatrice = workbookMatrice.getWorksheet('C');
        worksheetMatrice.eachRow(function(row, rowNumber) {
            var codeLogement = getCodeLogement(row.getCell(5).value)
            // console.log (codeLogement + ' : ' + isCodeLogement(codeLogement))
            if (isCodeLogement(codeLogement)) {
                // 1/ Je cherche 
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
      return null
  }