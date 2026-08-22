Public Function getSecurityID(rowNum As Long)

Dim result() As String
Dim findUsername() As String
result = Split(Cells(rowNum, 6).Value, vbLf)
findSecurityID = Split(result(3), ":")
getSecurityID = Trim(findSecurityID(1))

End Function



Public Function getAccountName(rowNum As Long)
  
Dim result() As String
Dim findAccountName() As String
result = Split(Cells(rowNum, 6).Value, vbLf)
findAccountName = Split(result(4), ":")
getAccountName = Trim(findAccountName(1))

End Function



Public Function getAccountDomain(rowNum As Long)

Dim result() As String
Dim findAccountDomain() As String
result = Split(Cells(rowNum, 6).Value, vbLf)
findAccountDomain = Split(result(5), ":")
getAccountDomain = Trim(findAccountDomain(1))
End Function



Public Function getShareName(rowNum As Long)

Dim result() As String
Dim findShareName() As String
result = Split(Cells(rowNum, 6).Value, vbLf)
findShareName = Split(result(14), ":")
getShareName = Trim(findShareName(1))

End Function



Public Function getSharePath(rowNum As Long)

Dim result() As String
Dim findSharePath As String
Dim colonPos As Long
Dim result2 As String

result = Split(Cells(rowNum, 6).Value, vbLf)
colonPos = InStr(result(15), ":")
result2 = Mid(result(15), colonPos + 1)
getSharePath = Trim(result2)

End Function



Public Function getSharePathName(rowNum As Long)

Dim result() As String
Dim sharePathName As String
Dim trimmed_findRelativeTargetName() As String

result = Split(Cells(rowNum, 6).Value, vbLf)
findRelativeTargetName = Split(result(16), ":")
trimmed_findRelativeTargetName = Split(findRelativeTargetName(1), "\")
getSharePathName = Trim(trimmed_findRelativeTargetName(0))

End Function

Public Function getRelativeTargetName(rowNum As Long)

Dim result() As String
Dim findRelativeTargetName() As String

result = Split(Cells(rowNum, 6).Value, vbLf)
findRelativeTargetName = Split(result(16), ":")
getRelativeTargetName = Trim(findRelativeTargetName(1))

End Function


Public Function getSourceAddress(rowNum As Long)

Dim result() As String
Dim findSourceAddress() As String

result = Split(Cells(rowNum, 6).Value, vbLf)
findSourceAddress = Split(result(10), ":")
getSourceAddress = Trim(findSourceAddress(1))

End Function



Public Sub importE3users()

Dim FilePath As String

FilePath = "C:\Users\snarendr\Documents\e3Users.txt"

Cells(1, 14).Value = "E3 Users"

With ActiveSheet.QueryTables.Add( _
    Connection:="TEXT;" & FilePath, _
    Destination:=Range("N2"))
    
    .TextFileParseType = xlDelimited
    .TextFileCommaDelimiter = True
    .Refresh
    
End With

End Sub



Public Sub HightlightDups(co1 As Range, co2 As Range)

co1.FormatConditions.Delete

'Applying the conditional formating
co1.FormatConditions.Add _
Type:=xlExpression, _
Formula1:="=COUNTIF(" & co2.Address & "," & _
co1.Cells(1, 1).Address(False, False) & ")>0"
    
co1.FormatConditions(1).Interior.Color = RGB(255, 255, 0)
    
End Sub



Sub Testing()

HightlightDups Columns("G"), Columns("N")
HightlightDups Columns("N"), Columns("G")

End Sub



Public Function DuplicateandCut()

Dim Cell3 As Range

ActiveSheet.Name = "F3 only"

Dim wsNew As Worksheet

ActiveSheet.Copy After:=ActiveSheet

Set wsNew = ActiveSheet

wsNew.Name = "E# and F3"

Dim LastRow As Long
Dim i As Long

LastRow = Cells(Rows.Count, "G").End(xlUp).Row

For i = LastRow To 2 Step -1

'For Each Cell3 In Range("G2:G" & Cells(Rows.Count, "G").End(xlUp).Row)

If Cells(i, "G").DisplayFormat.Interior.Color = RGB(255, 255, 0) Then
    Range("A" & i & ":M" & i).Delete Shift:=xlUp
End If

Next i
