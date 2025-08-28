VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMovement 
   Caption         =   "UserForm1"
   ClientHeight    =   8436
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11100
   OleObjectBlob   =   "frmMovement.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================================
' Proc�dure : UserForm_Initialize
' Objectif  : Initialiser le formulaire "Ajouter un mouvement" en d�finissant les r�f�rences
'             aux donn�es et en configurant l'interface graphique (front-end)
' ==============================================================================================
Private Sub UserForm_Initialize()
' ----------------------------------------------------------------------------------------------
' Section pour d�claration des variables et initialisation des r�f�rences
' ----------------------------------------------------------------------------------------------

' R�f�rence au classeur qui contient la macro
Set wb = ThisWorkbook

' R�f�rence � la feuille "stock" contenant la liste des mat�riels
Set wsStock = wb.Worksheets("stock")

' R�f�rence au tableau structur� nomm� "stock"
Set tabStock = wsStock.ListObjects("stock")

' R�f�rence � la feuille "mouvement" contenant l�historique des mouvements
Set wsMovement = wb.Worksheets("mouvement")

' R�f�rence au tableau structur� nomm� "movement" (historique des entr�es/sorties)
Set tabMovement = wsMovement.ListObjects("movement")

' R�f�rence � la plage physique correspondant au tableau "movement"
Set rangeMovement = wsMovement.Range("movement")

' Variables pour d�couper les adresses de plages et d�terminer le nombre de lignes
Dim rangeStockAddressPart() As String
Dim rangeStockLastLine As Long
Dim rangeMovementAddressPart() As String
Dim rangeMovementLastLine As Long
Dim startDate As Date
Dim endDate As Date

' ----------------------------------------------------------------------------------------------
' Section pour d�finir le front-end du formulaire (dimensions, titre, couleurs)
' ----------------------------------------------------------------------------------------------

' Propri�t�s g�n�rales du formulaire
With Me
    .Width = 420
    .Height = 300
    .Caption = "Ajouter un mouvement"
    .BackColor = COLOR_GRAY_DARK
End With

' Label "Mat�riel"
With lblMovementItem
    .Left = 30
    .Top = 20
    .Width = 100
    .Height = 20
    .Caption = "Mat�riel"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Liste d�roulante pour choisir le mat�riel
With cmbMovementItem
    .Left = 140
    .Top = 20
    .Width = 250
    .Height = 20
    .Style = fmStyleDropDownList
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .RowSource = tabStock
End With

' Label "Date"
With lblMovementDate
    .Left = 30
    .Top = 60
    .Width = 100
    .Height = 20
    .Caption = "Date"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Liste d�roulante pour la date
With cmbMovementDate
    .Left = 140
    .Top = 60
    .Width = 250
    .Height = 20
    .MaxLength = 10
    .Style = fmStyleDropDownList
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Alimentation de la liste d�roulante pour la date � partir d'une p�riode donn�e
startDate = CDate("01/01/2025")
endDate = CDate("31/12/2025")

While endDate <> startDate
    cmbMovementDate.addItem (endDate)
    endDate = endDate - 1
Wend

' Label "Type"
With lblMovementType
    .Left = 30
    .Top = 100
    .Width = 100
    .Height = 20
    .Caption = "Type"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Bouton radio "Entr�e"
With rdbEntry
    .Left = 140
    .Top = 100
    .Width = 50
    .Height = 20
    .Caption = "Entr�e"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Bouton radio "Sortie"
With rdbExit
    .Left = 210
    .Top = 100
    .Width = 50
    .Height = 20
    .Caption = "Sortie"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Label "Valeur"
With lblMovementValue
    .Left = 30
    .Top = 140
    .Width = 115
    .Height = 20
    .Caption = "Valeur"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Zone de saisie pour la valeur
With txtMovementValue
    .Left = 140
    .Top = 140
    .Width = 90
    .Height = 20
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Label "Description"
With lblMovementDescription
    .Left = 30
    .Top = 180
    .Width = 100
    .Height = 20
    .Caption = "Description"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Zone de saisie pour la description
With txtMovementDescription
    .Left = 140
    .Top = 180
    .Width = 250
    .Height = 20
    .MaxLength = 30
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Bouton "Ajouter" ? Valide et enregistre le mouvement
With btnAddMovement
    .Left = 100
    .Top = 225
    .Width = 100
    .Height = 25
    .Caption = "Ajouter"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_FOREST_GREEN
    .ForeColor = COLOR_WHITE
End With

' Bouton "Annuler" ? Ferme le formulaire sans action
With btnCancelAddMovement
    .Left = 220
    .Top = 225
    .Width = 100
    .Height = 25
    .Caption = "Annuler"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_CRIMSON_DARK
    .ForeColor = COLOR_WHITE
End With
End Sub

' ----------------------------------------------------------------------------------------------
' Bouton "Ajouter" : enregistre un nouveau mouvement et met � jour la liste
' ----------------------------------------------------------------------------------------------
Private Sub btnAddMovement_Click()
 Dim wb As Workbook
 Dim wsStock As Worksheet, wsMovement As Worksheet
 Dim tabStock As ListObject, tabMovement As ListObject
 Dim rangeStock As Range
 Dim moveItemLabel As String, moveType As String, moveDescription As String
 Dim moveDate As Date
 Dim moveValue As Variant
 Dim activeItemRowTab As Variant
 Dim activeItemCurrentQuantity As Variant
 Dim i As Long

' R�f�rences de base
' Classeur actif
Set wb = ThisWorkbook
' Feuille stock
Set wsStock = wb.Worksheets("stock")
' Feuille mouvements
Set wsMovement = wb.Worksheets("mouvement")
' Tableau structur� des stocks
Set tabStock = wsStock.ListObjects("stock")
' Tableau structur� mouvements
Set tabMovement = wsMovement.ListObjects("movement")

' Lecture des donn�es saisies par l'utilisateur
' Nom du mat�riel
 moveItemLabel = Trim(cmbMovementItem.Value)
  
' Si vide ? on sort
If Len(moveItemLabel) = 0 Then Exit Sub

' Conversion en date
moveDate = CDate(cmbMovementDate.Value)

' D�termination du type de mouvement
If rdbEntry.Value Then
    moveType = "entr�e"
ElseIf rdbExit.Value Then
    moveType = "sortie"
Else
    Exit Sub
End If

' R�cup�ration de la valeur du mouvement
moveValue = txtMovementValue.Value

' V�rification que la valeur est bien num�rique sinon elle sera �gale � 0
If Not IsNumeric(moveValue) Then
    moveValue = 0
End If

' V�rification que la valeur est bien positive sinon elle sera convertie en nombre positif : -5 deviendra 5
If moveValue < 0 Then
    moveValue = moveValue - (moveValue * 2)
End If
    
' Description libre
moveDescription = CStr(LCase(Trim(txtMovementDescription.Value)))
 
' Localisation du mat�riel dans le stock
Set rangeStock = tabStock.ListColumns(1).DataBodyRange
activeItemRowTab = Application.Match(moveItemLabel, rangeStock, 0)
If IsError(activeItemRowTab) Then Exit Sub    ' Si introuvable ? sortie

' Quantit� actuelle avant mouvement
activeItemCurrentQuantity = tabStock.DataBodyRange.Cells(activeItemRowTab, 2).Value

' Mise � jour de la quantit� en stock
If moveType = "entr�e" Then
    tabStock.DataBodyRange.Cells(activeItemRowTab, 2).Value = CLng(activeItemCurrentQuantity) + moveValue
Else
    tabStock.DataBodyRange.Cells(activeItemRowTab, 2).Value = CLng(activeItemCurrentQuantity) - moveValue
End If

' Mise � jour de la date de derni�re MAJ
tabStock.DataBodyRange.Cells(activeItemRowTab, 4).Value = moveDate

' Ajout de l'enregistement dans le tableau "movement"
addMovement moveDate, moveType, moveValue, moveDescription, moveItemLabel

' Rafra�chissement de la liste lstITems (vue d'ensemble des stocks)
frmStock.lstItems.Clear
For i = 1 To tabStock.DataBodyRange.Rows.Count
    frmStock.lstItems.addItem tabStock.DataBodyRange.Cells(i, 1).Value
    frmStock.lstItems.List(frmStock.lstItems.ListCount - 1, 1) = tabStock.DataBodyRange.Cells(i, 2).Value
    frmStock.lstItems.List(frmStock.lstItems.ListCount - 1, 2) = tabStock.DataBodyRange.Cells(i, 3).Value
    frmStock.lstItems.List(frmStock.lstItems.ListCount - 1, 3) = tabStock.DataBodyRange.Cells(i, 4).Value
Next i

' Rafra�chissement de la liste lstItemHistorical (vue individualis�e des mouvements)
With frmStock.lstItemHistorical
    ' Supprime le lien direct au tableau
    .RowSource = ""
    ' Nombre de colonnes dans la liste
    .ColumnCount = 4
    .Clear
End With

 If Not tabMovement.DataBodyRange Is Nothing Then
        Dim idx As Long, j As Long
        For j = 1 To tabMovement.DataBodyRange.Rows.Count
            ' Filtre sur colonne 5 ? Mat�riel
            If CStr(tabMovement.DataBodyRange.Cells(j, 5).Value) = CStr(moveItemLabel) Then
                ' Ajout d�une ligne dans la listbox historique
                frmStock.lstItemHistorical.addItem tabMovement.DataBodyRange.Cells(j, 1).Value
                idx = frmStock.lstItemHistorical.ListCount - 1
                ' Colonnes suivantes : Type, Valeur, Description
                frmStock.lstItemHistorical.List(idx, 1) = tabMovement.DataBodyRange.Cells(j, 2).Value
                frmStock.lstItemHistorical.List(idx, 2) = tabMovement.DataBodyRange.Cells(j, 3).Value
                frmStock.lstItemHistorical.List(idx, 3) = tabMovement.DataBodyRange.Cells(j, 4).Value
            End If
        Next j
    End If
    
' Fermeture du formulaire
Unload Me
End Sub

' ----------------------------------------------------------------------------------------------
' Bouton "Annuler" : ferme simplement le formulaire sans rien modifier
' ----------------------------------------------------------------------------------------------
Private Sub btnCancelAddMovement_Click()
    Unload Me
End Sub

