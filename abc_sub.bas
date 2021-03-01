Attribute VB_Name = "abc_sub"
Const mo As String = "abc_key"
Const com As String = ",": Const spe As String = " "

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX    debug_main: run debug                                                   XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub debug_main()
'setkey(val) => val = "variable,type,value"
Call setkey
'val = "var1,string,value1" => abc_key : Const var1 As String = "value1"
Call setkey("var1,string,value1")
'val = "int1,integer,1" => abc_key : Const int1 As Integer = 1
Call setkey("int1,integer,1")
'debug.print abc_key list
Call outputkey

'getkey(var) => var = "variable"
Debug.Print "return value of getkey" & vbLf & " => " & getkey
'var = "var1" => debug.print : Const var1 As String = "value1"
Debug.Print "return value of getkey(""var1"")" & vbLf & " => " & getkey("var1")
'var = "var1",n = 1 => debug.print : "value1"
Debug.Print "return value of getkey(""var1"", 1)" & vbLf & " => " & getkey("var1", 1)
'debug.print abc_key list

'change variable value
Call setkey("var1,string,value2")
Call outputkey(2)

'delete variable
Call delkey("var1")
Call outputkey(3)

'delete lines of all
'Call dellines
'Call outputkey(4)

End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX    setkey: add variables to modules(abc_key)                               XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub setkey(Optional val As Variant = "abc,string,a")
Dim rw As Integer, lrw As Integer
Dim abc As Variant
val = Split(val, com)
If UBound(val) < 2 Then Exit Sub
rw = checkkey(CStr(val(0))): lrw = rowcount
abc = "const " & val(0) & " as " & val(1) & " = " & val(2) & ""
abc = Split(abc, spe)
If StrConv(abc(3), vbLowerCase) = "string" Then
  abc(UBound(abc)) = """" & abc(UBound(abc)) & """"
End If
With ThisWorkbook.VBProject.VBComponents(mo).codemodule
  If rw = 0 Then
    .InsertLines lrw + 1, Join(abc, spe)
  Else
    .ReplaceLine rw, Join(abc, spe)
  End If
End With
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX    rowcount: return the last line of modules(abc_key)                      XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function rowcount()
rowcount = ThisWorkbook.VBProject.VBComponents(mo).codemodule.countoflines
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX    checkkey: return the line number of the modules(abc_key) variable       XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function checkkey(Optional var As String = "abc")
Dim i As Integer, lrw As Integer: lrw = rowcount
Dim abc As Variant
checkkey = 0
With ThisWorkbook.VBProject.VBComponents(mo).codemodule
  For i = 1 To lrw
    abc = .Lines(i, 1): abc = Split(abc, spe)
    If abc(1) = var Then checkkey = i: Exit For
  Next
End With
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX    rowcount: return the value of modules(abc_key) variable                 XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function getkey(Optional var As String = "abc", Optional n As Integer = 0)
Dim rw As Integer: rw = checkkey(var)
Dim abc As Variant
If rw = 0 Then Exit Function
abc = ThisWorkbook.VBProject.VBComponents(mo).codemodule.Lines(rw, 1)
If n > 0 Then _
  abc = Split(abc, spe): abc = abc(UBound(abc))
getkey = abc
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX    delkey: delete the variable of modules(abc_key)                         XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function delkey(Optional var As String = "abc")
Dim rw As Integer: rw = checkkey(var)
If rw = 0 Then Exit Function
ThisWorkbook.VBProject.VBComponents(mo).codemodule.DeleteLines rw
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX    dellines: delete all variable in the modules(abc_key)                   XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Function dellines()
Dim lrw As Integer: lrw = rowcount
ThisWorkbook.VBProject.VBComponents(mo).codemodule.DeleteLines 1, lrw
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XX    outputkey: outputs a list of modules(abc_key) variables                 XX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Sub outputkey(Optional n As Integer = 1)
Dim i As Integer, lrw As Integer: lrw = rowcount
With ThisWorkbook.VBProject.VBComponents(mo).codemodule
  Debug.Print "=================================="
  Debug.Print "   (" & n & ")  abc_key module : list"
  Debug.Print "----------------------------------"
  For i = 1 To lrw
    Debug.Print "Line" & i & ". " & .Lines(i, 1)
  Next
  If lrw = 0 Then Debug.Print "none Line"
  Debug.Print "=================================="
End With
End Sub
