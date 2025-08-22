VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStock 
   Caption         =   "UserForm1"
   ClientHeight    =   8532
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15648
   OleObjectBlob   =   "frmStock.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
' ==============================================================================================
' Proc�dure : UserForm_Initialize
' Objectif  : Initialiser l�interface et les variables du formulaire de gestion du stock
' D�clenchement : Automatique � l'ouverture du UserForm
' ==============================================================================================

' ----------------------------------------------------------------------------------------------
' Section pour d�claration des variables et initialisation des r�f�rences
' ----------------------------------------------------------------------------------------------

' R�f�rence au classeur actif (celui qui contient le code)
Set wb = ThisWorkbook

' R�f�rence � la feuille "stock" (donn�es des mat�riels en stock)
Set wsStock = wb.Worksheets("stock")

' R�f�rence � la feuille "mouvement" (historique des entr�es/sorties de stock)
Set wsMovement = wb.Worksheets("mouvement")

' R�f�rence au tableau structur� nomm� "stock" pr�sent dans wsStock
Set tabStock = wsStock.ListObjects("stock")

' R�f�rence � la plage de cellules couvrant le tableau "stock"
Set rangeStock = wsStock.Range("stock")

' R�f�rence au tableau structur� nomm� "movement" pr�sent dans wsMovement
Set tabMovement = wsMovement.ListObjects("movement")

' R�f�rence � la plage de cellules couvrant le tableau "movement"
Set rangeMovement = wsMovement.Range("movement")

' Variables servant au d�coupage d'adresse de plage et calcul de lignes pour "stock"
Dim rangeStockAddressPart() As String
Dim rangeStockLastLine As Long

' Variables servant au d�coupage d'adresse de plage et calcul de lignes pour "movement"
Dim rangeMovementAddressPart() As String
Dim rangeMovementLastLine As Long

' ----------------------------------------------------------------------------------------------
' Section pour d�finir le front-end du formulaire (dimensions, titre, couleurs)
' ----------------------------------------------------------------------------------------------

With Me
    ' Largeur totale du formulaire
    .Width = 900
    ' Hauteur totale du formulaire
    .Height = 520
    ' Titre affich� dans la barre du formulaire
    .Caption = "Stock du service informatique"
    ' Couleur d�arri�re-plan g�n�rale
    .BackColor = COLOR_GRAY_DARK
End With

' ----------------------------------------------
' Section : Contr�les de recherche et filtres
' ----------------------------------------------

' Zone de texte de recherche mat�riel
With txtSearchItem
    .Left = 20
    .Top = 20
    .Width = 265
    .Height = 22
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_IRON
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Bouton de lancement de la recherche
With btnSearchItem
    .Left = 295
    .Top = 20
    .Width = 125
    .Height = 22
    .Caption = "Recherche"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_SLATE
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Bouton filtrant les mat�riels en faible quantit�
With btnFilterLowQuantity
    .Left = 430
    .Top = 20
    .Width = 120
    .Height = 22
    .Caption = "Quantit�s faibles"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_WHITE
End With

' ----------------------------------------------
' Section : Boutons de tri des donn�es
' ----------------------------------------------

' Tri par libell� (nom mat�riel)
With btnSortItemLabel
    .Left = 20
    .Top = 50
    .Width = 125
    .Height = 22
    .Caption = "Trier par nom"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_SLATE
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Tri par quantit� en stock
With btnSortItemQuantity
    .Left = 155
    .Top = 50
    .Width = 130
    .Height = 22
    .Caption = "Trier par quantit�"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_SLATE
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Tri par cat�gorie
With btnSortItemCategory
    .Left = 295
    .Top = 50
    .Width = 125
    .Height = 22
    .Caption = "Trier par cat�gorie"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_SLATE
    .ForeColor = COLOR_GRAY_LIGHT
End With

' Tri par date de mise � jour
With btnSortItemUpdateDate
    .Left = 430
    .Top = 50
    .Width = 120
    .Height = 22
    .Caption = "Trier par date"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_SLATE
    .ForeColor = COLOR_GRAY_LIGHT
End With

' ----------------------------------------------
' Section : Liste des items (ListBox principale)
' ----------------------------------------------

With lstItems
    .Left = 20
    .Top = 80
    .Width = 531
    .Height = 380
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_IRON
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
    .SpecialEffect = fmSpecialEffectFlat
    
    ' Structure de la ListBox : nombre et largeur des colonnes
    .ColumnCount = 4
    .ColumnWidths = "190;50;170;115"
   
    ' R�cup�ration de l'adresse du tableau "stock" pour d�terminer la derni�re ligne
    rangeStockAddress = rangeStock.Address
    ' D�coupe l'adresse en parties ($A$1 ? {"","A","1"})
    rangeStockAddressPart = Split(rangeStockAddress, "$")
    ' R�cup�re le num�ro de ligne de fin (le 4�me �l�ment de la chaine d�coup�e)
    rangeStockLastLine = CLng(rangeStockAddressPart(4))

    ' Remplissage de la ListBox avec les donn�es du tableau
    ' On commence � la ligne 3 pour ignorer l�en-t�te du tableau
    For i = 3 To rangeStockLastLine
        ' Colonne 0 : Libell�
        .addItem tabStock.Range.Cells(i - 1, 1)
        ' Colonne 1 : Quantit�
        .List(.ListCount - 1, 1) = tabStock.Range.Cells(i - 1, 2)
        ' Colonne 2 : Cat�gorie
        .List(.ListCount - 1, 2) = tabStock.Range.Cells(i - 1, 3)
        ' Colonne 3 : Date / autre info
        .List(.ListCount - 1, 3) = tabStock.Range.Cells(i - 1, 4)
    Next i
End With

' ----------------------------------------------
' Section : D�tails du mat�riel (panneau de droite)
' ----------------------------------------------

' Bouton pour sauvegarder les modifications sur un item
With btnSaveItemUpdate
    .Left = 780
    .Top = 45
    .Width = 75
    .Height = 25
    .Caption = "Sauvegarder"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_SILVER_GLINT
    .ForeColor = COLOR_GRAY_DARK
    ' D�sactiv� par d�faut tant qu'aucune modification n�est en cours
    .Enabled = False
End With

' Titre du panneau de d�tails
With lblItemDetail
    .Left = 580
    .Top = 50
    .Width = 200
    .Height = 25
    .Caption = "D�tail du mat�riel"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Label du champ "Libell�"
With lblItemLabel
    .Left = 580
    .Top = 80
    .Width = 80
    .Height = 20
    .Caption = "Libell�"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propri�t�s de la zone de texte pour le libell� du mat�riel
With txtItemLabel
    .Left = 675
    .Top = 80
    .Width = 180
    .Height = 20
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propri�t�s du label pour indiquer la cat�gorie
With lblItemCategory
    .Left = 580
    .Top = 118
    .Width = 80
    .Height = 20
    .Caption = "Cat�gorie"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propri�t�s de la liste d�roulante pour choisir la cat�gorie
With cmbItemCategory
    .Left = 675
    .Top = 118
    .Width = 180
    .Height = 20
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
    .RowSource = "category"
End With

' Propri�t�s du label pour indiquer la sous-cat�gorie
With lblItemSubcategory
    .Left = 580
    .Top = 156
    .Width = 80
    .Height = 20
    .Caption = "Sous-cat�gorie"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propri�t�s de la liste d�roulante pour s�lectionner la sous-cat�gorie
With cmbItemSubcategory
    .Left = 675
    .Top = 156
    .Width = 180
    .Height = 20
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propri�t�s du label pour afficher le texte "En stock"
With lblItemCurrentQuantity
    .Left = 580
    .Top = 194
    .Width = 80
    .Height = 20
    .Caption = "En stock"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propri�t�s de la zone de texte pour saisir la quantit� actuelle
With txtItemCurrentQuantity
    .Left = 675
    .Top = 194
    .Width = 50
    .Height = 20
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propri�t�s du label pour symboliser "quantit� minimale" (>=)
With lblItemMinQuantity
    .Left = 750
    .Top = 194
    .Width = 60
    .Height = 20
    .Caption = ">="
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propri�t�s de la zone de texte pour saisir la quantit� minimale
With txtItemMinQuantity
    .Left = 804
    .Top = 194
    .Width = 50
    .Height = 20
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propri�t�s du label pour la date de mise � jour
With lblItemUpdateDate
    .Left = 580
    .Top = 232
    .Width = 80
    .Height = 20
    .Caption = "Date de MAJ"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propri�t�s de la zone de texte pour saisir ou afficher la date de mise � jour
With txtItemUpdateDate
    .Left = 675
    .Top = 232
    .Width = 180
    .Height = 22
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propri�t�s du label pour le champ commentaire
With lblItemComment
    .Left = 580
    .Top = 270
    .Width = 80
    .Height = 20
    .Caption = "Commentaire"
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' Propri�t�s de la zone de texte pour saisir un commentaire libre
With txtItemComment
    .Left = 675
    .Top = 270
    .Width = 180
    .Height = 22
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_SMALL
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' ----------------------------------------------
' Section : Historique des mouvements
' ----------------------------------------------

' Label titre de la section historique
With lblItemHistorical
    .Left = 580
    .Top = 310
    .Width = 200
    .Height = 25
    .Caption = "Historique des mouvements"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .BackColor = COLOR_GRAY_DARK
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
End With

' ListBox affichant l'historique
With lstItemHistorical
    .Left = 580
    .Top = 335
    .Width = 275
    .Height = 155
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_EXTRA_SMALL
    .BackColor = COLOR_GRAY_IRON
    .ForeColor = COLOR_GRAY_LIGHT
    .BorderColor = COLOR_GRAY_LIGHT
    .SpecialEffect = fmSpecialEffectFlat
    .ColumnCount = 4
    .ColumnWidths = "70;45;30;125"
End With

' ----------------------------------------------
' Section : Boutons d'action
' ----------------------------------------------

' Bouton pour ajouter un nouveau mat�riel
With btnAddItem
    .Left = 20
    .Top = 450
    .Width = 170
    .Height = 35
    .Caption = "Nouveau"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .BackColor = COLOR_FOREST_GREEN
    .ForeColor = COLOR_WHITE
End With

' Bouton pour supprimer un mat�riel s�lectionn�
With btnDeleteItem
    .Left = 200
    .Top = 450
    .Width = 170
    .Height = 35
    .Caption = "Supprimer"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .BackColor = COLOR_CRIMSON_DARK
    .ForeColor = COLOR_WHITE
End With

' Bouton pour enregistrer un mouvement (entr�e ou sortie) sur un mat�riel
With btnAddMovement
    .Left = 380
    .Top = 450
    .Width = 170
    .Height = 35
    .Caption = "Mouvement"
    .Font.Bold = True
    .Font.Name = FONT_NAME
    .Font.Size = FONT_SIZE_LARGE
    .BackColor = COLOR_NAVY_SLATE
    .ForeColor = COLOR_WHITE
End With
End Sub

' ----------------------------------------------------------------------------------------------
' �v�nement : Clic sur le bouton "Ajouter un mat�riel"
' Action : Ouvre le formulaire de gestion d'un nouveau mat�riel
' ----------------------------------------------------------------------------------------------
Private Sub btnAddItem_Click()
' Affiche le formulaire frmItem en mode modal
 frmItem.Show
End Sub

' ----------------------------------------------------------------------------------------------
' �v�nement : Clic sur le bouton "Ajouter un mouvement"
' Action : Ouvre le formulaire d'enregistrement d'un nouveau mouvement (entr�e ou sortie)
' ----------------------------------------------------------------------------------------------
Private Sub btnAddMovement_Click()
' Affiche le formulaire frmMovement en mode modal
frmMovement.Show
End Sub

' ----------------------------------------------------------------------------------------------
' Proc�dure : displayItems
' Objectif : Vider et recharger la liste principale (lstItems) avec les donn�es du tableau "stock"
' ----------------------------------------------------------------------------------------------
Private Sub displayItems()

' D�claration des variables pour g�rer les lignes et d�couper l'adresse du tableau
Dim rangeStockAddressPart() As String
Dim rangeStockLastLine As Long
Dim i As Long

' R�cup�re l'adresse de la plage "stock"
rangeStockAddress = rangeStock.Address

' S�pare l'adresse en parties (ex: "$A$1:$D$20" -> {"","A","1","","D","20"})
rangeStockAddressPart = Split(rangeStockAddress, "$")

' Convertit en nombre la derni�re ligne du tableau (�l�ment n�4 du tableau apr�s split)
rangeStockLastLine = CLng(rangeStockAddressPart(4))

' Vide la liste avant de la recharger
lstItems.Clear

' Boucle � partir de la 3e ligne (pour sauter l�en-t�te du tableau)
For i = 3 To rangeStockLastLine
    ' Ajoute un nouvel �l�ment dans la premi�re colonne
    lstItems.addItem tabStock.Range.Cells(i - 1, 1)
    ' Remplit les colonnes 2 � 4 avec les donn�es correspondantes
    lstItems.List(lstItems.ListCount - 1, 1) = tabStock.Range.Cells(i - 1, 2) ' Quantit�
    lstItems.List(lstItems.ListCount - 1, 2) = tabStock.Range.Cells(i - 1, 3) ' Cat�gorie
    lstItems.List(lstItems.ListCount - 1, 3) = tabStock.Range.Cells(i - 1, 4) ' Date ou autre info
Next i

End Sub

' ----------------------------------------------------------------------------------------------
' �v�nement : lstItems_Change
' Objectif : Afficher les d�tails du mat�riel s�lectionn� + son historique
' ----------------------------------------------------------------------------------------------
Private Sub lstItems_Change()

 ' Variables pour stocker les informations de l�mat�riel actif
 Dim activeItemLabel As String
 Dim activeItemCategory As String
 Dim activeItemSubcategory As String
 Dim activeItemCurrentQuantity As Integer
 Dim activeItemMinQuantity As Integer
 Dim activeItemUpdateDate As Date
 Dim activeItemComment As String
 Dim lastRow As Long

 ' Active le bouton de sauvegarde (un �l�ment est s�lectionn�)
 btnSaveItemUpdate.Enabled = True

 ' ----------------------------
 ' Pr�paration de la liste historique
 ' ----------------------------

 ' R�cup�re l'adresse du tableau "movement"
 rangeMovementAdress = rangeMovement.Address

 ' S�pare l'adresse en parties pour identifier la derni�re ligne
 rangeMovementAddressPart = Split(rangeMovementAdress, "$")
 rangeMovementLastLine = CLng(rangeMovementAddressPart(4)) + 1

 ' ----------------------------
 ' R�cup�ration des infos du mat�riel s�lectionn�
 ' ----------------------------
 ' �vite les erreurs si mat�riel est introuvable
 On Error Resume Next

 ' Libell� de l�mat�riel s�lectionn�
 activeItemLabel = lstItems.Value


 ' Permet d'actualiser le formulaire dynamique apr�s enregistrement
 lastRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row
 Set rangeStock = wsStock.Range("A3:G" & lastRow)

 If activeItemLabel = "" Then
     ' Aucun �l�ment s�lectionn� ? r�initialise les champs
     txtItemLabel.Value = ""
     cmbItemCategory.Value = ""
     cmbItemSubcategory.Value = ""
     txtItemCurrentQuantity.Value = ""
     txtItemMinQuantity = ""
     txtItemUpdateDate = ""
     txtItemComment = ""
 Else
     ' Recherche des infos dans le tableau "stock" avec VLOOKUP
     activeItemCategory = WorksheetFunction.VLookup(activeItemLabel, rangeStock, 3, False)
     activeItemSubcategory = WorksheetFunction.VLookup(activeItemLabel, rangeStock, 6, False)
     activeItemCurrentQuantity = WorksheetFunction.VLookup(activeItemLabel, rangeStock, 2, False)
     activeItemMinQuantity = WorksheetFunction.VLookup(activeItemLabel, rangeStock, 5, False)
     activeItemUpdateDate = WorksheetFunction.VLookup(activeItemLabel, rangeStock, 4, False)
     activeItemComment = WorksheetFunction.VLookup(activeItemLabel, rangeStock, 7, False)
     
     ' Affichage dans les contr�les du formulaire
     txtItemLabel.Value = activeItemLabel
     cmbItemCategory.Value = activeItemCategory
     cmbItemSubcategory.Value = activeItemSubcategory
     txtItemCurrentQuantity.Value = activeItemCurrentQuantity
     txtItemCurrentQuantity.Value = activeItemCurrentQuantity
     txtItemMinQuantity.Value = activeItemMinQuantity
     txtItemUpdateDate.Value = activeItemUpdateDate
     txtItemComment = activeItemComment
     
     ' ----------------------------
     ' Remplissage de l�historique
     ' ----------------------------
     lstItemHistorical.Clear
     For i = 2 To rangeMovementLastLine + 2
         ' V�rifie si la colonne 5 du mouvement correspond � mat�riel s�lectionn�
         If tabMovement.Range.Cells(i - 1, 5) = activeItemLabel Then
            ' Date mouvement
             lstItemHistorical.addItem tabMovement.Range.Cells(i - 1, 1)
             ' Type (entr�e/sortie)
             lstItemHistorical.List(lstItemHistorical.ListCount - 1, 1) = tabMovement.Range.Cells(i - 1, 2)
             ' Quantit�
             lstItemHistorical.List(lstItemHistorical.ListCount - 1, 2) = tabMovement.Range.Cells(i - 1, 3)
             ' Commentaire
             lstItemHistorical.List(lstItemHistorical.ListCount - 1, 3) = tabMovement.Range.Cells(i - 1, 4)
         End If
     Next i
 End If
 
 Exit Sub
 ' R�initialise la gestion des erreurs
 On Error GoTo 0

End Sub

' ----------------------------------------------------------------------------------------------
' �v�nement : cmbItemCategory_Change
' Objectif : Modifie les options de sous-cat�gorie selon la cat�gorie choisie
' ----------------------------------------------------------------------------------------------
Private Sub cmbItemCategory_Change()

Dim categoryChoiced As String

categoryChoiced = cmbItemCategory.Value

If categoryChoiced = "Accessoire" Then
    cmbItemSubcategory.RowSource = "accessorie"
ElseIf categoryChoiced = "Composant Interne" Then
    cmbItemSubcategory.RowSource = "internal_component"
ElseIf categoryChoiced = "Connectique/C�blage" Then
    cmbItemSubcategory.RowSource = "connector_cabling"
ElseIf categoryChoiced = "Consommable" Then
    cmbItemSubcategory.RowSource = "consumable"
ElseIf categoryChoiced = "Imprimante/Scanner" Then
    cmbItemSubcategory.RowSource = "printer_scanner"
ElseIf categoryChoiced = "Logiciel/Licence" Then
    cmbItemSubcategory.RowSource = "software_licence"
ElseIf categoryChoiced = "Mat�riel de Bureau" Then
    cmbItemSubcategory.RowSource = "office_equipment"
ElseIf categoryChoiced = "Mat�riel Mobile" Then
    cmbItemSubcategory.RowSource = "mobile_hardware"
ElseIf categoryChoiced = "Mat�riel R�seau" Then
    cmbItemSubcategory.RowSource = "network_hardware"
ElseIf categoryChoiced = "P�riph�rique" Then
    cmbItemSubcategory.RowSource = "peripheral"
ElseIf categoryChoiced = "Stockage" Then
    cmbItemSubcategory.RowSource = "storage"
End If
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur le bouton "Trier par nom" ? tri ascendant
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemLabel_Click()
SortStockByLabelAscending
' Rafra�chit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Double-clic sur le bouton "Trier par nom" ? tri descendant
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
SortStockByLabelDescending
' Rafra�chit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur le bouton "Trier par quantit�" ? croissant
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemQuantity_Click()
SortStockByCurrentQuantityAscending
' Rafra�chit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Double-clic sur le bouton "Trier par quantit�" ? d�croissant
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemQuantity_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
SortStockByCurrentQuantityDescending
' Rafra�chit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur le bouton "Trier par cat�gorie" ? A ? Z
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemCategory_Click()
SortStockByCategoryAscending
' Rafra�chit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Double-clic sur le bouton "Trier par cat�gorie" ? Z ? A
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemCategory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 SortStockByCategoryDescending
' Rafra�chit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur le bouton "Trier par date" ? plus ancien au plus r�cent
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemUpdateDate_Click()
 SortStockByUpdateDateAscending
' Rafra�chit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Double-clic sur le bouton "Trier par date" ? plus r�cent au plus ancien
' ----------------------------------------------------------------------------------------------
Private Sub btnSortItemUpdateDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
SortStockByUpdateDateDescending
' Rafra�chit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur "Quantit�s faibles" : affiche uniquement les mat�riels dont la quantit� = quantit� mini
' ----------------------------------------------------------------------------------------------
Private Sub btnFilterLowQuantity_Click()
Dim rangeStockAddressPart() As String
Dim rangeStockLastLine As Long

' R�cup�re l�adresse et le num�ro de ligne max de la table "stock"
rangeStockAddress = rangeStock.Address
rangeStockAddressPart = Split(rangeStockAddress, "$")
rangeStockLastLine = CLng(rangeStockAddressPart(4))

' Vide la liste avant de la remplir
lstItems.Clear

' Parcours des mat�riels
For i = 3 To rangeStockLastLine
    ' Col2 = Qt� actuelle / Col5 = Qt� minimum ? filtre
    If tabStock.Range.Cells(i - 1, 2) <= tabStock.Range(i - 1, 5) Then
        lstItems.addItem tabStock.Range.Cells(i - 1, 1)
        lstItems.List(lstItems.ListCount - 1, 1) = tabStock.Range.Cells(i - 1, 2)
        lstItems.List(lstItems.ListCount - 1, 2) = tabStock.Range.Cells(i - 1, 3)
        lstItems.List(lstItems.ListCount - 1, 3) = tabStock.Range.Cells(i - 1, 4)
    End If
Next i
End Sub

' ----------------------------------------------------------------------------------------------
' Double-clic sur "Quantit�s faibles" : r�initialise l�affichage
' ----------------------------------------------------------------------------------------------
Private Sub btnFilterLowQuantity_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' Rafra�chit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur recherche" : filtre la liste selon le texte saisi dans txtSearchItem
' ----------------------------------------------------------------------------------------------
Private Sub btnSearchItem_Click()
Dim rangeStockAddressPart() As String
Dim rangeStockLastLine As Long
Dim userSearch As String
Dim itemLabelValue As String

' D�termine la plage "stock"
rangeStockAddress = rangeStock.Address
rangeStockAddressPart = Split(rangeStockAddress, "$")
rangeStockLastLine = CLng(rangeStockAddressPart(4))

' R�cup�re et pr�pare le texte recherch�
userSearch = CStr(LCase(Trim(txtSearchItem.Value)))

' Vide la liste avant de la remplir avec les r�sultats
lstItems.Clear

' Boucle sur les mat�riels
For i = 3 To rangeStockLastLine
    itemLabelValue = CStr(LCase(Trim(tabStock.Range(i - 1, 1).Value)))
    ' InStr(..., ..., vbTextCompare) = 1 ? commence par le texte recherch�
    If InStr(1, itemLabelValue, userSearch, vbTextCompare) = 1 Then
        lstItems.addItem tabStock.Range.Cells(i - 1, 1)
        lstItems.List(lstItems.ListCount - 1, 1) = tabStock.Range.Cells(i - 1, 2)
        lstItems.List(lstItems.ListCount - 1, 2) = tabStock.Range.Cells(i - 1, 3)
        lstItems.List(lstItems.ListCount - 1, 3) = tabStock.Range.Cells(i - 1, 4)
    End If
Next i

    ' R�initialise la zone de recherche
    txtSearchItem.Value = ""
End Sub

' ----------------------------------------------------------------------------------------------
' Double-clic sur "Recherche" : r�affiche tous les mat�riels
' ----------------------------------------------------------------------------------------------
Private Sub btnSearchItem_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
' Rafra�chit la liste principale
displayItems
End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur "Supprimer" : efface l�mat�riel s�lectionn� apr�s confirmation
' ----------------------------------------------------------------------------------------------
Private Sub btnDeleteItem_Click()
Dim activeItemLabel As String
Dim rowToDelete As Variant
Dim lastRow As Long
Dim i As Long

On Error Resume Next
activeItemLabel = lstItems.Value
If activeItemLabel = "" Then Exit Sub

' Recalcul plage avant recherche
lastRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row
Set rangeStock = wsStock.Range("A3:A" & lastRow)

On Error Resume Next
rowToDelete = Application.Match(activeItemLabel, rangeStock, 0)
On Error GoTo 0

If IsError(rowToDelete) Then
    MsgBox "�l�ment introuvable."
    Exit Sub
End If

' Demande de confirmation avant suppression
If MsgBox("Confirmer la suppression de : " & activeItemLabel, vbYesNo) = vbYes Then
    ' +2 car plage commence � A3
    wsStock.Rows(rowToDelete + 2).Delete
End If

' Recalcul plage apr�s suppression
lastRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row
Set rangeStock = wsStock.Range("A3:A" & lastRow)

' Recharge la liste apr�s suppression
lstItems.Clear
For i = 3 To lastRow
    lstItems.addItem wsStock.Cells(i, 1).Value
    lstItems.List(lstItems.ListCount - 1, 1) = wsStock.Cells(i, 2).Value
    lstItems.List(lstItems.ListCount - 1, 2) = wsStock.Cells(i, 3).Value
    lstItems.List(lstItems.ListCount - 1, 3) = wsStock.Cells(i, 4).Value
Next i

' Sauvegarde le classeur
wb.Save

End Sub

' ----------------------------------------------------------------------------------------------
' Clic simple sur "Sauvegarder" : enregistre les modifications apport�es au mat�riel s�lectionn�
' ----------------------------------------------------------------------------------------------
Private Sub btnSaveItemUpdate_Click()
    Dim activeItemLabel As String
    Dim rowToUpdate As Variant
    Dim lastRow As Long
    Dim saveConfirmation As VbMsgBoxResult
    Dim i As Long

    ' V�rifier la s�lection de l'�l�ment dans la ListBox
    If lstItems.ListIndex = -1 Then
        MsgBox "Veuillez s�lectionner un mat�riel � modifier.", vbExclamation
        Exit Sub
    End If
    
    ' Conserver la valeur de l'�l�ment s�lectionn�
    activeItemLabel = lstItems.Value
    
    '  Re-calculer la plage de recherche
    lastRow = wsStock.Cells(wsStock.Rows.Count, "A").End(xlUp).Row
    Dim rangeStock As Range
    Set rangeStock = wsStock.Range("A3:A" & lastRow)
    
    ' Chercher la ligne correspondante avec Application.Match
    On Error Resume Next
    rowToUpdate = Application.Match(activeItemLabel, rangeStock, 0)
    On Error GoTo 0
    
    ' G�rer l'erreur si mat�riel n'est pas trouv�
    If IsError(rowToUpdate) Then
        MsgBox "Erreur : mat�riel s�lectionn� n'a pas �t� trouv� dans le tableau.", vbCritical
        Exit Sub
    End If
    
    ' Demander confirmation avant la sauvegarde
    saveConfirmation = MsgBox("Confirmer la sauvegarde des modifications", vbYesNo)
    
    If saveConfirmation = vbYes Then
        ' Mettre � jour les donn�es sur la feuille de calcul
        ' Le +2 est n�cessaire car la plage commence � la ligne 3 (rowToUpdate est un index bas� sur la plage)
        wsStock.Cells(rowToUpdate + 2, 1).Value = txtItemLabel.Value
        wsStock.Cells(rowToUpdate + 2, 3).Value = cmbItemCategory.Value
        wsStock.Cells(rowToUpdate + 2, 6).Value = cmbItemSubcategory.Value
        wsStock.Cells(rowToUpdate + 2, 5).Value = txtItemMinQuantity.Value
        wsStock.Cells(rowToUpdate + 2, 7).Value = txtItemComment.Value
        
        MsgBox "Modifications sauvegard�es avec succ�s !", vbInformation
    End If

    ' Recharger la ListBox
    lstItems.Clear
    For i = 3 To lastRow
        lstItems.addItem wsStock.Cells(i, 1).Value
        lstItems.List(lstItems.ListCount - 1, 1) = wsStock.Cells(i, 2).Value
        lstItems.List(lstItems.ListCount - 1, 2) = wsStock.Cells(i, 3).Value
        lstItems.List(lstItems.ListCount - 1, 3) = wsStock.Cells(i, 4).Value
    Next i
    
    ' Res�lectionner l'�l�ment dans la ListBox
    For i = 0 To lstItems.ListCount - 1
        If lstItems.List(i, 0) = activeItemLabel Then
            lstItems.Selected(i) = True
            Exit For
        End If
    Next i
    
' Sauvegarde le classeur
wb.Save
End Sub

