VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmItem 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmItem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
' ==============================================================================================
' Proc�dure : UserForm_Initialize
' Objectif  : Initialiser le formulaire "Ajouter un mat�riel" en d�finissant les r�f�rences
'             aux donn�es et en configurant l'interface graphique (front-end)
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' Section pour d�finir le front-end du formulaire (dimensions, titre, couleurs)
' ----------------------------------------------------------------------------------------------

' Param�tres de base du formulaire
With Me
    .Width = 280
    .Height = 180
    .Caption = "Ajouter un mat�riel"
    .BackColor = COLOR_GRAY_DARK
End With

' Label descriptif "Libell� du nouveau mat�riel"
With lblAddItem
    .Left = 40
    .Top = 20
    .Width = 200
    .Height = 20
    .Caption = "Libell� du nouveau mat�riel"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .Font.Bold = True
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Zone de saisie pour le libell� du mat�riel
With txtAddItem
    .Left = 40
    .Top = 55
    .Width = 200
    .Height = 20
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_IRON
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
    .MaxLength = 50
End With

' Bouton "Ajouter"
With btnAddItem
    .Left = 40
    .Top = 100
    .Width = 95
    .Height = 25
    .Caption = "Ajouter"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .Font.Bold = True
    .BackColor = COLOR_FOREST_GREEN
    .ForeColor = COLOR_WHITE
End With

' Bouton "Annuler"
With btnCancelAddItem
    .Left = 145
    .Top = 100
    .Width = 95
    .Height = 25
    .Caption = "Annuler"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .Font.Bold = True
    .BackColor = COLOR_CRIMSON_DARK
    .ForeColor = COLOR_WHITE
End With
End Sub

' ----------------------------------------------------------------------------------------------
' Bouton "Ajouter" : enregistre un nouveau mat�riel et met � jour la liste
' ----------------------------------------------------------------------------------------------
Private Sub btnAddItem_Click()
Dim itemLabel As String
Dim rangeStockAddressPart() As String
Dim rangeStockLastLine As Long

' R�cup�re le texte saisi
itemLabel = txtAddItem.Value

' Met en minuscule et retire espaces inutiles
itemLabel = LCase(Trim(itemLabel))

' Appelle la proc�dure d'ajout dans la base
addItem (itemLabel)

' Vide le champ de saisie
txtAddItem.Value = ""

On Error Resume Next
' R�cup�re la derni�re ligne du tableau "stock"
rangeStockAddress = rangeStock.Address
rangeStockAddressPart = Split(rangeStockAddress, "$")
rangeStockLastLine = CLng(rangeStockAddressPart(4))

' Rafra�chit la liste principale (frmStock.lstItems)
frmStock.lstItems.Clear
For i = 3 To rangeStockLastLine + 1
    frmStock.lstItems.addItem tabStock.Range.Cells(i - 1, 1)
    frmStock.lstItems.List(frmStock.lstItems.ListCount - 1, 1) = tabStock.Range.Cells(i - 1, 2)
    frmStock.lstItems.List(frmStock.lstItems.ListCount - 1, 2) = tabStock.Range.Cells(i - 1, 3)
    frmStock.lstItems.List(frmStock.lstItems.ListCount - 1, 3) = tabStock.Range.Cells(i - 1, 4)
Next i

' Ferme le formulaire apr�s ajout
Unload Me
End Sub

' ----------------------------------------------------------------------------------------------
' Bouton "Annuler" : ferme simplement le formulaire sans rien modifier
' ----------------------------------------------------------------------------------------------
Private Sub btnCancelAddItem_Click()
    Unload Me
End Sub

