Attribute VB_Name = "modStockSortingSub"
' ==============================================================================================
' Proc�dure : SortStockByLabelAscending
' Objectif  : Trier le stock par libell� (A ? Z)
' ==============================================================================================
Public Sub SortStockByLabelAscending()
    ' On efface les crit�res de tri pr�c�dents
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    
    ' Ajout du crit�re : colonne [libell�] en ordre croissant
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libell�]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    ' Application du tri
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        ' La premi�re ligne contient les en-t�tes
        .Header = xlYes
        ' Ne pas tenir compte de la casse
        .MatchCase = False
        ' Tri vertical
        .Orientation = xlTopToBottom
        ' M�thode de tri compatible avec caract�res accentu�s
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Proc�dure : SortStockByLabelDescending
' Objectif  : Trier le stock par libell� (Z ? A)
' ==============================================================================================
Public Sub SortStockByLabelDescending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libell�]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Proc�dure : SortStockByCurrentQuantityAscending
' Objectif  : Trier par quantit� (croissante), puis libell� (A ? Z)
' ==============================================================================================
Public Sub SortStockByCurrentQuantityAscending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ' Premier crit�re : quantit�
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[stock]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ' Deuxi�me crit�re : libell�
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libell�]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Proc�dure : SortStockByCurrentQuantityDescending
' Objectif  : Trier par quantit� (d�croissante), puis libell� (A ? Z)
' ==============================================================================================
Public Sub SortStockByCurrentQuantityDescending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ' Premier crit�re : quantit� d�croissante
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[stock]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ' Deuxi�me crit�re : libell�
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libell�]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Proc�dure : SortStockByCategoryAscending
' Objectif  : Trier par cat�gorie (A ? Z), puis libell� (A ? Z)
' ==============================================================================================
Public Sub SortStockByCategoryAscending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ' Premier crit�re : cat�gorie
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[cat�gorie]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ' Deuxi�me crit�re : libell�
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libell�]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Proc�dure : SortStockByCategoryDescending
' Objectif  : Trier par cat�gorie (Z ? A), puis libell� (A ? Z)
' ==============================================================================================
Public Sub SortStockByCategoryDescending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ' Premier crit�re : cat�gorie d�croissante
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[cat�gorie]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ' Deuxi�me crit�re : libell�
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libell�]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Proc�dure : SortStockByUpdateDateAscending
' Objectif  : Trier par date de mise � jour (ancienne ? r�cente), puis libell� (A ? Z)
' ==============================================================================================
Public Sub SortStockByUpdateDateAscending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ' Premier crit�re : date croissante
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[maj]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ' Deuxi�me crit�re : libell�
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libell�]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' ==============================================================================================
' Proc�dure : SortStockByUpdateDateDescending
' Objectif  : Trier par date de mise � jour (r�cent ? ancien), puis libell� (A ? Z)
' ==============================================================================================
Public Sub SortStockByUpdateDateDescending()
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Clear
    ' Premier crit�re : date d�croissante
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[maj]"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ' Deuxi�me crit�re : libell�
    ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort.SortFields.Add2 _
        Key:=Range("stock[libell�]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("stock").ListObjects("stock").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


