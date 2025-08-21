Attribute VB_Name = "modStockModificationSub"
' ==============================================================================================
' Proc�dure : addItem
' Objectif  : Ajouter un nouvel mat�riel dans le tableau "stock"
' Param�tre : itemLabel libell� du mat�riel � ajouter
' ==============================================================================================
Public Sub addItem(itemLabel As String)
    ' R�f�rences aux objets principaux
     ' Classeur contenant la macro
    Set wb = ThisWorkbook
    ' Feuille de calcul "stock"
    Set wsStock = wb.Worksheets("stock")
    ' Tableau structur� nomm� "stock"
    Set tabStock = wsStock.ListObjects("stock")
    ' Plage physique correspondant au tableau
    Set rangeStock = wsStock.Range("stock")

    ' Ajout d'une nouvelle ligne � la fin du tableau
    Set newStockRow = tabStock.ListRows.Add

    ' Remplissage des cellules de la nouvelle ligne
    ' Colonne 1 : libell� du mat�riel
    newStockRow.Range.Cells(1).Value = itemLabel

    ' Colonne 8 : num�ro de ligne relatif dans le tableau (ROW() - 2 car en-t�tes + ligne d�part)
    newStockRow.Range.Cells(8).Formula = "=ROW()-2"

    ' Colonne 9 : num�ro de ligne absolu dans la feuille
    newStockRow.Range.Cells(9).Formula = "=ROW()"
End Sub

' ==============================================================================================
' Proc�dure : addMovement
' Objectif  : Ajouter un nouvel enregistrement dans le tableau "movement"
' Param�tres:
'   moveDate            Date du mouvement
'   moveType            Type de mouvement ("Entr�e" ou "Sortie")
'   moveValue           Quantit� du mouvement
'   moveDescription     Description
'   moveItem            Libell� du mat�riel concern�
' ==============================================================================================
Public Sub addMovement(moveDate As Date, moveType As String, moveValue As Integer, moveDescription As String, moveItem As String)
    ' R�f�rences aux objets principaux
    ' Classeur contenant la macro
    Set wb = ThisWorkbook
    ' Feuille "mouvement"
    Set wsMovement = wb.Worksheets("mouvement")
    ' Tableau structur� "movement"
    Set tabMovement = wsMovement.ListObjects("movement")
    ' Plage physique du tableau
    Set rangeMovement = wsMovement.Range("movement")

    ' Ajout d'une nouvelle ligne
    Set newMovementRow = tabMovement.ListRows.Add

    ' Remplissage des cellules de la nouvelle ligne
    ' Colonne 1 : date du mouvement
    newMovementRow.Range.Cells(1).Value = moveDate

    ' Colonne 2 : type de mouvement
    newMovementRow.Range.Cells(2).Value = moveType

    ' Colonne 3 : quantit� associ�e au mouvement
    newMovementRow.Range.Cells(3).Value = moveValue

    ' Colonne 4 : description
    newMovementRow.Range.Cells(4).Value = moveDescription

    ' Colonne 5 : libell� du mat�riel concern�
    newMovementRow.Range.Cells(5).Value = moveItem
End Sub

