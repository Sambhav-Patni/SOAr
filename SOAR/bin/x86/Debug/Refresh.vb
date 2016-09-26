Public Sub Refresh()
' Macro to update my Power Query script(s)

Dim cn As WorkbookConnection

For Each cn In ThisWorkbook.Connections
If Left(cn, 13) = "Power Query -" Then cn.Refresh
Next cn
End Sub