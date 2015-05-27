VERSION 5.00
Begin VB.Form frmFuncGraph 
   Caption         =   "Function Graph"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   LinkTopic       =   "Form3"
   ScaleHeight     =   5250
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Save Image"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   45
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   4605
      Left            =   135
      ScaleHeight     =   4545
      ScaleWidth      =   6435
      TabIndex        =   0
      Top             =   495
      Width           =   6495
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmFuncGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim img As BinaryImage
Dim pGraph As CGraph
Dim loaded As Boolean

Private Sub Command2_Click()

   If img Is Nothing Then Exit Sub
   
   pth = App.path & "\sample.gif"
   If img.Save(pth) Then
        MsgBox "Saved to " & pth, vbInformation
   Else
        MsgBox "Save failed", vbExclamation
   End If
   
   'or SavePicture Picture1, App.Path & "\sample.bmp"
    
End Sub

Function GraphFrom(startfunc As String, Optional pNode As CNode)
    
    'On Error Resume Next
    
    Dim li As ListItem
    Dim data As String
    Dim foundEnd As Boolean
    Dim func() As String
    Dim n As CNode
    Dim existingNode As CNode
    Dim startLine As Long
    Dim topLevel As Boolean
    
    If Not loaded Then Form_Load
    If pGraph Is Nothing Then Set pGraph = New CGraph
    
    If pNode Is Nothing Then
        Set pNode = pGraph.AddNode(startfunc)
        topLevel = True
    End If

    For Each li In Form2.lvFunc.ListItems
        If li.Text = startfunc Then
            startLine = CLng(li.tag)
            Exit For
        End If
    Next
    
    data = Form2.ExtractFunction(startLine, foundEnd)
    
    'now we trim off the function xx(){ part..
    a = InStr(data, "{")
    If a > 0 Then data = Mid(data, a)
    
    For Each li In Form2.lvFunc.ListItems
        If InStr(data, li.Text & "(") > 0 Then
            Set existingNode = pGraph.NodeExists(li.Text)
            If Not existingNode Is Nothing Then
                existingNode.ConnectTo pNode
            Else
                Set n = pGraph.AddNode(li.Text)
                pNode.ConnectTo n
                GraphFrom li.Text, n
            End If
        End If
    Next
    
    If Not topLevel Then Exit Function
    
    pGraph.GenerateGraph

    Set img = pGraph.dot.ToGIF(pGraph.lastGraph)
    
    If img Is Nothing Then
        Text1.Visible = True
        Text1.Text = "Graph generation failed?" & vbCrLf & vbCrLf & pGraph.lastGraph
    Else
        Set Picture1.Picture = img.Picture
        Me.Width = Picture1.Width
        Me.height = Picture1.height
    End If

End Function

Function GraphTo(func As String)

    MsgBox "todo!", vbExclamation
    
' 'On Error Resume Next
'
'    Dim li As ListItem
'    Dim data As String
'    Dim foundEnd As Boolean
'    Dim func() As String
'    Dim n As CNode
'    Dim existingNode As CNode
'    Dim startLine As Long
'    Dim topLevel As Boolean
'
'    If Not loaded Then Form_Load
'    If pGraph Is Nothing Then Set pGraph = New CGraph
'
'    If pNode Is Nothing Then
'        Set pNode = pGraph.AddNode(startfunc)
'        topLevel = True
'    End If
'
'    For Each li In Form2.lvFunc.ListItems
'        If li.Text = startfunc Then
'            startLine = CLng(li.tag)
'            Exit For
'        End If
'    Next
'
'    data = Form2.ExtractFunction(startLine, foundEnd)
'
'    'now we trim off the function xx(){ part..
'    a = InStr(data, "{")
'    If a > 0 Then data = Mid(data, a)
'
'    For Each li In Form2.lvFunc.ListItems
'        If InStr(data, li.Text & "(") > 0 Then
'            Set existingNode = pGraph.NodeExists(li.Text)
'            If Not existingNode Is Nothing Then
'                existingNode.ConnectTo pNode
'            Else
'                Set n = pGraph.AddNode(li.Text)
'                pNode.ConnectTo n
'                GraphFrom li.Text, n
'            End If
'        End If
'    Next
'
'    If Not topLevel Then Exit Function
'
'    pGraph.GenerateGraph
'
'    Set img = pGraph.dot.ToGIF(pGraph.lastGraph)
'
'    If img Is Nothing Then
'        Text1.Visible = True
'        Text1.Text = "Graph generation failed?" & vbCrLf & vbCrLf & pGraph.lastGraph
'    Else
'        Set Picture1.Picture = img.Picture
'        Me.Width = Picture1.Width
'        Me.height = Picture1.height
'    End If
'
End Function

'example
'Dim g As New CGraph
'   Dim n0 As CNode, n1 As CNode, n2 As CNode, n3 As CNode, n4 As CNode, n5 As CNode
'
'   Set n0 = g.AddNode("this is my" & vbCrLf & "multiline\nnode")
'   n0.shape = "box"
'   n0.style = "filled"
'   n0.color = "lightyellow"
'   n0.fontcolor = "#c0c0c0"
'
'   Set n1 = g.AddNode
'   Set n2 = g.AddNode
'   Set n3 = g.AddNode
'   Set n4 = g.AddNode
'   Set n5 = g.AddNode
'
'   n0.ConnectTo n2
'   n1.ConnectTo n2
'   n2.ConnectTo n3
'   n1.ConnectTo n4
'   n0.ConnectTo n5
'
'   Call g.GenerateGraph
'   Text1.Text = g.lastGraph
'
'   Set img = g.dot.ToGIF(g.lastGraph)
'   If img Is Nothing Then Exit Sub
'
'   Set Picture1.Picture = img.Picture

Private Sub Form_Load()
    Text1.Visible = False
    With Picture1
        Text1.Move .left, .Top, .Width, .height
    End With
    Me.Visible = True
    loaded = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With Picture1
        .Width = Me.Width - .left - 200
        .height = Me.height - .Top - 200
        Text1.Move .left, .Top, .Width, .height
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set pGraph = Nothing
    Set img = Nothing
    loaded = False
End Sub
