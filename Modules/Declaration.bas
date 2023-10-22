Attribute VB_Name = "Declaration"
Public Function isExistinLV(ByRef srcLV As ListView, ByVal strFind As String, ByVal inFirst As Boolean, Optional numCol As Byte) As Boolean
If srcLV.ListItems.Count < 1 Then Exit Function
Dim i As Long
For i = 1 To srcLV.ListItems.Count
    srcLV.ListItems(i).Selected = True
    If inFirst = True Then
        If srcLV.SelectedItem = strFind Then isExistinLV = True: Exit For
    Else
        If srcLV.SelectedItem.ListSubItems(numCol) = strFind Then isExistinLV = True: Exit For
    End If
Next i
i = 0
End Function
Public Sub ClearText(ByRef SRC_FORM As Form)
On Error Resume Next
Dim Control As Control
    For Each Control In SRC_FORM.Controls
        If (TypeOf Control Is TextBox) Then Control = vbNullString
    Next Control
    Set Control = Nothing
End Sub
Public Sub INSERT_RECORD(ByVal SQL As String)
Set ADD_REC = New ADODB.Command
ADD_REC.ActiveConnection = cn
ADD_REC.CommandText = SQL
ADD_REC.Execute
End Sub
Public Sub Highlight(ByRef srcText)
On Error Resume Next
    With srcText
        .SelStart = 0
        .SelLength = Len(srcText.Text)
    End With
End Sub

Function Capitalize(c As Integer) As Integer
    If c >= 97 And c <= 122 Then
        c = c - 32
    End If
    Capitalize = c
End Function

