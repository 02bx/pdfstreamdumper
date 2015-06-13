Attribute VB_Name = "modActionScript"

'quick and dirty processor..
Function ProcessActionScript(buffer As String) As String

    tmp = Replace(buffer, "public", Empty, , , vbTextCompare)
    tmp = Replace(tmp, "private", Empty, , , vbTextCompare)
    tmp = Replace(tmp, "static", Empty, , , vbTextCompare)
    
    'i dont want to loose this info..but it screws with the formatter which then
    'screws with the function scan feature..
    
'    tmp = Replace(tmp, ":ByteArray", " /*:ByteArray*/ ", , , vbTextCompare)
'    tmp = Replace(tmp, ":String", " /*:String*/ ", , , vbTextCompare)
'    tmp = Replace(tmp, ":Boolean", " /*:Boolean*/ ", , , vbTextCompare)
'    tmp = Replace(tmp, ":uint", " /*:uint*/ ", , , vbTextCompare)
'    tmp = Replace(tmp, ":int", " /*:int*/ ", , , vbTextCompare)
'    tmp = Replace(tmp, ":Number", " /*:Number*/ ", , , vbTextCompare)
'    tmp = Replace(tmp, ":Event", " /*:Event*/ ", , , vbTextCompare)
'    tmp = Replace(tmp, ":void", " /*:void*/ ", , , vbTextCompare)
    
    tmp = Replace(tmp, ":ByteArray", Empty, , , vbTextCompare)
    tmp = Replace(tmp, ":String", Empty, , , vbTextCompare)
    tmp = Replace(tmp, ":Boolean", Empty, , , vbTextCompare)
    tmp = Replace(tmp, ":uint", Empty, , , vbTextCompare)
    tmp = Replace(tmp, ":int", Empty, , , vbTextCompare)
    tmp = Replace(tmp, ":Number", Empty, , , vbTextCompare)
    tmp = Replace(tmp, ":Event", Empty, , , vbTextCompare)
    tmp = Replace(tmp, ":void", Empty, , , vbTextCompare)
    
    'Vector.<
    t = Split(tmp, vbCrLf)
    Dim w
    
    'first we strip all indenting..
    For i = 0 To UBound(t)
        t(i) = Trim(mltrim(t(i)))
    Next
    
    Dim hasPackage As Boolean
    Dim hasClass As Boolean
    Dim inFuncs As Boolean
    
    Dim gvars() As String
    Dim funcs() As String
    
    push gvars, Empty 'never empty
    push funcs, Empty
    
    For i = 0 To UBound(t)
        If i < UBound(t) - 1 Then
            If t(i) = "package" And t(i + 1) = "{" Then
                t(i) = "//package"
                t(i + 1) = "//{"
                i = i + 1
                hasPackage = True
                GoTo nextone
            End If
        End If
        
        w = getWord(t(i), 0)
        If w = "import" Or w = "class" Then
            t(i) = "//" & t(i)
            If w = "class" And t(i + 1) = "{" Then
                t(i + 1) = "//{"
                i = i + 1
                hasClass = True
                GoTo nextone
            End If
        End If
        
        If w = "function" Then inFuncs = True
        
        If Not inFuncs Or w = "function" Then
            a = InStr(1, t(i), "_SafeStr", vbTextCompare)
            If a > 0 Then
                ss = extractSafeStr(t(i), a)
                If Len(ss) > 0 Then
                    If Not inFuncs Then 'global variables
                        push gvars, ss & "->" & "g_var" & UBound(gvars)
                    Else
                        push funcs, ss & "->" & "func_" & UBound(funcs)
                    End If
                End If
            End If
        End If
                    
            
        
nextone:
    Next
     
   
   trimTrailingBrace = 0
   If hasPackage Then trimTrailingBrace = trimTrailingBrace + 1
   If hasClass Then trimTrailingBrace = trimTrailingBrace + 1
   
   If trimTrailingBrace > 0 Then
        For i = UBound(t) To 0 Step -1
            If trimTrailingBrace = 0 Then Exit For
            If left(t(i), 1) = "}" Then
                t(i) = "//" & t(i)
                trimTrailingBrace = trimTrailingBrace - 1
            End If
        Next
   End If
   
   tmp = Join(t, vbCrLf)
   
   For Each x In funcs
      If Len(x) > 0 Then
         Y = Split(x, "->")
         tmp = Replace(tmp, Y(0), Y(1))
      End If
   Next

   For Each x In gvars
      If Len(x) > 0 Then
         Y = Split(x, "->")
         tmp = Replace(tmp, Y(0), Y(1))
      End If
   Next
   
   ProcessActionScript = tmp
    
End Function

Function extractSafeStr(line, start) As String
     tmp = Mid(line, start)
     
     Dim a(8) As Long
     Dim lowest As Long
     lowest = 9999
     
     a(0) = InStr(tmp, ":")
     a(1) = InStr(tmp, "(")
     a(2) = InStr(tmp, ";")
     a(3) = InStr(tmp, vbCr)
     a(4) = InStr(tmp, vbLf)
     a(5) = InStr(tmp, " ")
     a(6) = InStr(tmp, "/")
     a(7) = InStr(tmp, "=")
     
     For i = 0 To UBound(a)
        If a(i) > 0 And a(i) < lowest Then lowest = a(i)
     Next
     
     If lowest <> 9999 Then
        extractSafeStr = Trim(Mid(tmp, 1, lowest - 1))
     End If
     
End Function


Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Function getWord(str, i) As String
    On Error Resume Next
    tmp = Split(str, " ")
    getWord = tmp(i)
End Function

Private Function mltrim(ByVal x) As String

    Do While VBA.left(x, 1) = " " Or VBA.left(x, 1) = vbTab
       If Len(x) = 1 Then Exit Function 'empty line
       x = Mid(x, 2)
    Loop
    
    mltrim = x
    
End Function

