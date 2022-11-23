Attribute VB_Name = "Module3"
Sub HyperlinkConverter()
    Dim tbl As ListObject
    Dim rng As Range
    Dim path As String
    
    Set tbl = ActiveSheet.ListObjects("Table1")
    Set rng = tbl.ListColumns(1).DataBodyRange
    
    'Debug.Print rng.Address
    
    ' For Each xCell In Range("M:O")
    For Each xCell In rng
        'Debug.Print xCell.Value
        If xCell.Value <> "" And xCell.Row <> 1 Then
            path = "https://app.copia.io/DFS/" & xCell.Value
            Debug.Print path
            ActiveSheet.Hyperlinks.Add Anchor:=xCell, Address:=path
        End If
    Next xCell
End Sub
