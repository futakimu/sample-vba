Function ExistsInTarget(Name As String, ParentID As Variant, TargetSheet As Worksheet) As Long
    Dim LastRow As Long, i As Long
    
    LastRow = TargetSheet.Cells(TargetSheet.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        If TargetSheet.Cells(i, 2).Value = Name And TargetSheet.Cells(i, 3).Value = ParentID Then
            ExistsInTarget = TargetSheet.Cells(i, 1).Value 'IDを返す
            Exit Function
        End If
    Next i
    
    ExistsInTarget = 0 '見つからなかった場合は0を返す
End Function

Sub ConvertRelationshipStructureModified()
    Dim SourceSheet As Worksheet
    Dim TargetSheet As Worksheet
    Dim LastRow As Long, TargetLastRow As Long, i As Long, j As Long
    Dim Name As String, ParentID As Variant, ExistingID As Long
    
    Set SourceSheet = ThisWorkbook.Sheets("Sheet1")
    Set TargetSheet = ThisWorkbook.Sheets.Add
    TargetSheet.Name = "ConvertedData"
    TargetSheet.Cells(1, 1).Value = "ID"
    TargetSheet.Cells(1, 2).Value = "Name"
    TargetSheet.Cells(1, 3).Value = "ParentID"
    
    LastRow = SourceSheet.Cells(SourceSheet.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        ParentID = Empty
        
        For j = 1 To 4
            Name = SourceSheet.Cells(i, j).Value
            
            If Not IsEmpty(Name) And Not IsNull(Name) Then
                ExistingID = ExistsInTarget(Name, ParentID, TargetSheet)
                
                If ExistingID = 0 Then '存在しない場合
                    TargetLastRow = TargetSheet.Cells(TargetSheet.Rows.Count, 1).End(xlUp).Row + 1
                    TargetSheet.Cells(TargetLastRow, 1).Value = TargetLastRow - 1
                    TargetSheet.Cells(TargetLastRow, 2).Value = Name
                    TargetSheet.Cells(TargetLastRow, 3).Value = ParentID
                    ParentID = TargetLastRow - 1
                Else
                    ParentID = ExistingID
                End If
            End If
        Next j
    Next i
End Sub
