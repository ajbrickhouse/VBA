Attribute VB_Name = "Module1"
Sub AddRowToTable()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim addedRow As ListRow
    Dim RepName As String
    Dim product_line As String
    
    Set ws = ActiveSheet    ' Select the current sheet
    
    RepName = Range("F3").Value
    
    ' Prompt user for Produc Line Name
    product_line = InputBox("Input a Product Line Name for (" & RepName & ").", "Product Line")
    
    ' Deal with results from Input box
    If StrPtr(product_line) = 0 Then ' If the InputBox is canceled
       Exit Sub
    ElseIf product_line = NullString Then ' If the InputBox string is blank
       Exit Sub
    Else ' If everything is
        Set tbl = ws.ListObjects("Table1")  ' Select table
        Set addedRow = tbl.ListRows.Add ' Add row to the end of the selected table
        With addedRow   ' add data to the addedRow
            .Range(1) = RepName   ' get the generated product name
            .Range(2) = product_line    ' string from the InputBox
        End With
    End If
    
    HyperlinkConverter
    
End Sub
