Attribute VB_Name = "modConstant"
' R�f�rence au classeur principal
Public wb As Workbook

' R�f�rence � la feuille contenant les donn�es de stock
Public wsStock As Worksheet

' Tableau structur� nomm� (ListObject) dans la feuille de stock
Public tabStock As ListObject

' Ligne nouvellement ajout�e dans le tableau de stock
Public newStockRow As ListRow

' Plage repr�sentant le tableau de stock
Public rangeStock As Range

' Plage repr�sentant l'adresse du tableau de stock
Public rangeStockAddress As String

' R�f�rence � la feuille contenant les mouvements de stock
Public wsMovement As Worksheet

' Tableau structur� nomm� (ListObject) dans la feuille de mouvement
Public tabMovement As ListObject

' Ligne nouvellement ajout�e dans le tableau des mouvement
Public newMovementRow As ListRow

' Plage repr�sentant le tableau des mouvement
Public rangeMovement As Range

' Plage repr�sentant l'adresse du tableau de mouvement
Public rangeMovementAdress As String

