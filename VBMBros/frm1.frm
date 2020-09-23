VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Map Editor"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   837
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   28
      Top             =   4440
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Object"
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   27
      Top             =   1920
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Tile"
      Height          =   255
      Index           =   0
      Left            =   7200
      TabIndex        =   26
      Top             =   1680
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   8400
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   24
      Top             =   4440
      Width           =   2895
   End
   Begin VB.ListBox List2 
      Height          =   2985
      ItemData        =   "frm1.frx":0000
      Left            =   11400
      List            =   "frm1.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Prop"
      Height          =   375
      Left            =   11400
      TabIndex        =   22
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   375
      Left            =   6720
      TabIndex        =   20
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton ReadMap 
      Caption         =   "Read Map"
      Height          =   375
      Left            =   7200
      TabIndex        =   19
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Write"
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Top             =   2880
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   7200
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Read Def"
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Property's"
      Height          =   375
      Left            =   10440
      TabIndex        =   16
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "SetRes"
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      Height          =   4215
      Left            =   120
      ScaleHeight     =   277
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   13
      Top             =   120
      Width           =   6975
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         Height          =   4095
         Left            =   0
         ScaleHeight     =   269
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   765
         TabIndex        =   14
         Top             =   0
         Width           =   11535
      End
      Begin Project1.ScrollBars ScrollBars1 
         Left            =   5760
         Top             =   1440
         _ExtentX        =   1085
         _ExtentY        =   1085
         ScrollContents  =   -1  'True
         HorizMaxVal     =   5000
         VertMaxVal      =   5000
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   10440
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   12
      Top             =   480
      Width           =   240
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "frm1.frx":0004
      Left            =   11400
      List            =   "frm1.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add >>"
      Height          =   375
      Left            =   10440
      TabIndex        =   10
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   10440
      TabIndex        =   9
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Tiles"
      Height          =   375
      Left            =   7200
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   4440
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   8400
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox Text3 
      Height          =   1455
      Left            =   5760
      TabIndex        =   5
      Top             =   4440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2566
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frm1.frx":0008
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Active Objects"
      Height          =   3975
      Left            =   8280
      TabIndex        =   25
      Top             =   4200
      Width           =   4095
      Begin VB.CommandButton Command14 
         Caption         =   "Delete"
         Height          =   375
         Left            =   1200
         TabIndex        =   33
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Replace"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   3240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Background"
      Height          =   3975
      Left            =   8280
      TabIndex        =   21
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton Command10 
         Caption         =   "Find Dup"
         Height          =   375
         Left            =   2160
         TabIndex        =   34
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Delete"
         Height          =   375
         Left            =   2160
         TabIndex        =   31
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Replace"
         Height          =   375
         Left            =   2160
         TabIndex        =   30
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Rebuild"
         Height          =   375
         Left            =   2160
         TabIndex        =   35
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Not Included!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2160
         TabIndex        =   29
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Col"
      Height          =   255
      Left            =   6720
      TabIndex        =   3
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Row"
      Height          =   255
      Left            =   7200
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public cBlocks As New Collection
Public actualBlock As New clsBlock
Public actualObj As New clsObjProp
Public MapFile As String, MapPlayersFile As String, TilesFile As String, PlayersDefFile As String, DefFile As String
Public colObjProp As Collection

Private Sub Command1_Click()
Text3.Text = "SourceLeft=" & """" & Val(Text1.Text) * 16 & """" & _
             Chr$(13) & "SourceRight=" & """" & (Val(Text1.Text) + 1) * 16 & """" & _
             Chr$(13) & "SourceTop=" & """" & Val(Text2.Text) * 16 & """" & _
             Chr$(13) & "SourceBottom=" & """" & (Val(Text2.Text) + 1) * 16 & """"
End Sub

Private Sub Command10_Click()
Dim cBlocksDup As New Collection, bl As clsBlock
On Error GoTo 0
For Each bl In cBlocks
    cBlocksDup.Add bl
Next
Do While cBlocksDup.Count = 0
    bl = cBlocksDup(1)
    cBlocksDup.Remove 1
    For i = 1 To cBlocksDup.Count
        If bl.SourceLeft = cBlocksDup(i).SourceLeft And _
            bl.SourceRight = cBlocksDup(i).SourceRight And _
            bl.SourceTop = cBlocksDup(i).SourceTop And _
            bl.SourceBottom = cBlocksDup(i).SourceBottom Then
            founded = "x"
            MsgBox "Char " & bl.CharType & " is in same position as " & cBlocksDup(i).CharType
        End If
    Next
Loop
If founded = "" Then
    MsgBox "No duplicates founded!"
End If
End Sub

Private Sub Command11_Click()
    On Error GoTo err
    If List1.ListIndex = -1 Then Exit Sub
    vVar = InputBox("Enter a new character (1)", "Replace character")
    For i = 1 To cBlocks.Count
        Set actualBlock = cBlocks(i)
        If actualBlock.CharType = List1.List(List1.ListIndex) Then
            old_char = actualBlock.CharType
            cBlocks.Remove i
            List1.RemoveItem List1.ListIndex
            actualBlock.CharType = vVar
            cBlocks.Add actualBlock, "c" & Asc(vVar)
            List1.AddItem vVar
            Exit For
        End If
    Next
    str1 = Text4.Text
    Do While InStr(1, str1, old_char)
        str1 = Replace$(str1, old_char, vVar)
    Loop
    Text4.Text = str1
    Exit Sub
err:
End Sub

Private Sub Command13_Click()
    On Error GoTo err
    If List2.ListIndex = -1 Then Exit Sub
    vVar = InputBox("Enter a new character (1)", "Replace character")
    For i = 1 To colObjProp.Count
        Set actualObj = colObjProp(i)
        If actualObj.CharType = List2.List(List2.ListIndex) Then
            old_char = actualObj.CharType
            colObjProp.Remove i
            List2.RemoveItem List2.ListIndex
            actualObj.CharType = vVar
            colObjProp.Add actualObj, "c" & Asc(vVar)
            List2.AddItem vVar
            Exit For
        End If
    Next
    str1 = Text6.Text
    Do While InStr(1, str1, old_char)
        str1 = Replace$(str1, old_char, vVar)
    Loop
    Text6.Text = str1
    Exit Sub
err:
End Sub

Private Sub Command15_Click()
Dim bl As clsBlock, cBlocksDup As Collection
    'On Error GoTo err
    asci = Asc("!")
    List1.Clear
    Set cBlocksDup = New Collection
    'If List1.ListIndex = -1 Then Exit Sub
    For j = 0 To Picture1.ScaleHeight - 16 Step 16
        For i = 0 To Picture1.ScaleWidth - 16 Step 16
            Set bl = New clsBlock
            'Debug.Print i & " - " & j
            bl.SourceLeft = i
            bl.SourceRight = i + 16
            bl.SourceTop = j
            bl.SourceBottom = j + 16
            bl.Col = (bl.SourceLeft + 1) \ 16
            bl.Row = (bl.SourceTop + 1) \ 16
            Set actualBlock = Nothing
            For X = 1 To cBlocks.Count
                Set actualBlock = cBlocks(X)
                If bl.SourceLeft = actualBlock.SourceLeft And _
                   bl.SourceRight = actualBlock.SourceRight And _
                   bl.SourceTop = actualBlock.SourceTop And _
                   bl.SourceBottom = actualBlock.SourceBottom Then
                   Exit For
                End If
            Next
            'Debug.Print cBlocks.Count
            If X > cBlocks.Count Then
                bl.CharType = Chr$(asci)
                cBlocksDup.Add bl
            Else
                actualBlock.CharType = Chr$(asci)
                cBlocksDup.Add actualBlock, "c" & asci
            End If
            List1.AddItem Chr$(asci)
            'Debug.Print Chr$(asci)
            asci = asci + 1
            If asci > Asc("z") Then Exit For
        Next
        If asci > Asc("z") Then Exit For
    Next
    Do While cBlocks.Count > 0
        cBlocks.Remove 1
    Loop
    For Each bl In cBlocksDup
        cBlocks.Add bl, "c" & Asc(bl.CharType)
    Next
err:
End Sub

Private Sub Command2_Click()
    CommonDialog1.FileName = TilesFile
    CommonDialog1.Filter = "*.gif;*.bmp;*.jpg"
    CommonDialog1.ShowOpen
    TilesFile = CommonDialog1.FileName
    Picture1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub Command3_Click()
    On Error GoTo err
    If Text5.Text = "" Then Exit Sub
    Set actualBlock = cBlocks("c" & Asc(Text5.Text))
    Text5.Text = ""
    Exit Sub
err:
    actualBlock.CharType = Text5.Text
    cBlocks.Add actualBlock, "c" & Asc(Text5.Text)
    List1.AddItem actualBlock.CharType
    List1.ListIndex = List1.ListCount - 1
    Label3.Visible = False
End Sub

Private Sub Command4_Click()
    vWidth = Val(InputBox("Set a value for width...", "Resolution"))
    If vWidth <> 0 Then
        Picture2.Width = vWidth
    End If
    vHeight = Val(InputBox("Set a value for height...", "Resolution"))
    If vHeight <> 0 Then
        Picture2.Height = vHeight
    End If
    ScrollBars1.HorizMaxVal = Abs(Picture2.ScaleWidth - Picture4.ScaleWidth + 24)
    ScrollBars1.VertMaxVal = Abs(Picture2.ScaleHeight - Picture4.ScaleHeight + 24)
    ScrollBars1.AdjustScrollInfo Picture4.hwnd
End Sub

Private Sub Command5_Click()
On Error GoTo err
    Set actualBlock = cBlocks("c" & Asc(List1.List(List1.ListIndex)))
    Form2.txtFixed = actualBlock.Fixed
    Form2.txtNonSolid = actualBlock.NonSolid
    Form2.txtJumpNextLevel = actualBlock.JumpNextLevel
    Form2.txtIsFloor = actualBlock.IsFloor
    Form2.Visible = True
err:
End Sub

Private Sub Command6_Click()
Dim def As String, arrObjects() As String
Dim vVar As String
    
    Set cBlocks = Nothing
    Set cBlocks = New Collection
    Set actualBlock = Nothing
    Set actualBlock = New clsBlock
    List1.Clear
    CommonDialog2.ShowOpen
    DefFile = CommonDialog2.FileName
    Me.Caption = "Map Editor - " & DefFile
'/ Read Title Map file
    TilesFile = GetFromIni("World", "Tiles", DefFile)
    TilesFile = Replace$(TilesFile, "AppPath", App.Path)
    Picture1.Picture = LoadPicture(TilesFile)
'/ Init Map Objects
    vVar = GetFromIni("World", "Objects", DefFile)
    'Debug.Print vVar
    arrObjects = Split(vVar, ",")
    For i = 0 To UBound(arrObjects)
        Call InitMapObjects(arrObjects(i), DefFile)
    Next
'/ Init Game Objects
    PlayersDefFile = GetFromIni("World", "PlayersDef", DefFile)
    vVar2 = GetFromIni("Game", "Objects", Replace$(PlayersDefFile, "AppPath", App.Path))
    arrObjects = Split(vVar2, ",")
    
    Set colObjProp = New Collection
    
    For i = 0 To UBound(arrObjects)
        Call InitObjects(arrObjects(i), PlayersDefFile)
    Next
'/ Read Map
    MapFile = GetFromIni("World", "Map", DefFile)
    MapFile = Replace$(MapFile, "AppPath", App.Path)
    If MapFile <> "" Then ReadFileByLine MapFile
    
    MapPlayersFile = GetFromIni("World", "MapPlayers", DefFile)
    MapPlayersFile = Replace$(MapPlayersFile, "AppPath", App.Path)
    If MapPlayersFile <> "" Then
        ReadPlayersByLine MapPlayersFile
    Else
        Command8_Click
    End If
    If Option1(1).Value = False Then Text6.Visible = False
End Sub

Public Sub InitObjects(vObj As String, vFilePath As String)

    Set vObjProp = New clsObjProp
    vFilePath = Replace$(vFilePath, "AppPath", App.Path)
    vObjProp.CharType = GetFromIni(vObj, "CharType", Replace$(vFilePath, "AppPath", App.Path))
    vObjProp.Picture = Replace$(GetFromIni(vObj, "Picture", Replace$(vFilePath, "AppPath", App.Path)), "AppPath", App.Path)
    vObjProp.CanHit = Val(GetFromIni(vObj, "CanHit", vFilePath))
    vObjProp.HittedByTop = Val(GetFromIni(vObj, "HittedByTop", vFilePath))
    vObjProp.HittedByLeft = Val(GetFromIni(vObj, "HittedByLeft", vFilePath))
    vObjProp.HittedByRight = Val(GetFromIni(vObj, "HittedByRight", vFilePath))
    vObjProp.HittedByBottom = Val(GetFromIni(vObj, "HittedByBottom", vFilePath))
    vObjProp.SourceLeft = Val(GetFromIni(vObj, "SourceLeft", vFilePath))
    vObjProp.SourceRight = Val(Val(GetFromIni(vObj, "SourceRight", vFilePath)))
    vObjProp.SourceTop = Val(GetFromIni(vObj, "SourceTop", vFilePath))
    vObjProp.SourceBottom = Val(GetFromIni(vObj, "SourceBottom", vFilePath))
    vObjProp.Velocity = Val(GetFromIni(vObj, "Velocity", vFilePath))
    vObjProp.JumpVelocity = Val(GetFromIni(vObj, "JumpVelocity", vFilePath))
    'vObjProp.JumpSize = Val(GetFromIni(vObj, "JumpSize", vFilePath))
    'vObjProp.AnimFrom = Val(GetFromIni(vObj, "AnimFrom", vFilePath))
    vObjProp.AnimTo = Val(GetFromIni(vObj, "AnimTo", vFilePath))
    vObjProp.MinRunframe = Val(GetFromIni(vObj, "MinRunframe", vFilePath))
    vObjProp.MaxRunframe = Val(GetFromIni(vObj, "MaxRunframe", vFilePath))
    vObjProp.PosLeft = Val(GetFromIni(vObj, "PosLeft", vFilePath))
    vObjProp.PosTop = Val(GetFromIni(vObj, "PosTop", vFilePath))
    vObjProp.PosRight = Val(GetFromIni(vObj, "PosRight", vFilePath))
    vObjProp.PosBottom = Val(GetFromIni(vObj, "PosBottom", vFilePath))
    vObjProp.User = Val(GetFromIni(vObj, "User", vFilePath))
    vObjProp.AI = GetFromIni(vObj, "AI", vFilePath)
    vObjProp.InitedAI = Val(GetFromIni(vObj, "InitedAI", vFilePath))
    vObjProp.InitAIWhenHitted = Val(GetFromIni(vObj, "InitAIWhenHitted", vFilePath))
    vObjProp.Visible = Val(GetFromIni(vObj, "Visible", vFilePath))
    vObjProp.CanFall = Val(GetFromIni(vObj, "CanFall", vFilePath))
    vObjProp.ChangeFrom = Val(GetFromIni(vObj, "ChangeFrom", vFilePath))
    vObjProp.JumpFrom = Val(GetFromIni(vObj, "JumpFrom", vFilePath))
    vObjProp.JumpTo = Val(GetFromIni(vObj, "JumpTo", vFilePath))
    vObjProp.DieFrame = Val(GetFromIni(vObj, "DieFrame", vFilePath))
    vObjProp.RemoveWhenDies = Val(GetFromIni(vObj, "RemoveWhenDies", vFilePath))
    vObjProp.MakeJumpWhenHitted = Val(GetFromIni(vObj, "MakeJumpWhenHitted", vFilePath))
    vObjProp.JumpWhenHitted = Val(GetFromIni(vObj, "JumpWhenHitted", vFilePath))
    vObjProp.Solid = Val(GetFromIni(vObj, "Solid", vFilePath))
    vObjProp.DieTiming = Val(GetFromIni(vObj, "DieTiming", vFilePath))
    vObjProp.Fixed = Val(GetFromIni(vObj, "Fixed", vFilePath))
    vObjProp.CreateWhenHitted = GetFromIni(vObj, "CreateWhenHitted", vFilePath)
    vObjProp.CreatePlace = GetFromIni(vObj, "CreatePlace", vFilePath)
    vObjProp.Hibernating = Val(GetFromIni(vObj, "Hibernating", vFilePath))
    vObjProp.CanHitEnemies = Val(GetFromIni(vObj, "CanHitEnemies", vFilePath))
    vObjProp.CannotHitUser = Val(GetFromIni(vObj, "CannotHitUser", vFilePath))
    vObjProp.ChangePlayerTo = GetFromIni(vObj, "ChangePlayerTo", vFilePath)
    vObjProp.CanBeBreaked = Val(GetFromIni(vObj, "CanBeBreaked", vFilePath))
    vObjProp.CanBreak = Val(GetFromIni(vObj, "CanBreak", vFilePath))
    vObjProp.NextLevel = Val(GetFromIni(vObj, "NextLevel", vFilePath))
    'vObjProp.direction = Val(GetFromIni(vObj, "Direction", vFilePath))
    vObjProp.UserSelection = Val(GetFromIni(vObj, "UserSelection", vFilePath))
    vObjProp.GrowTo = GetFromIni(vObj, "GrowTo", vFilePath)
    vObjProp.MakeGrow = Val(GetFromIni(vObj, "MakeGrow", vFilePath))
    vObjProp.DieWhenHits = Val(GetFromIni(vObj, "DieWhenHits", vFilePath))
    vObjProp.FireBall = GetFromIni(vObj, "FireBall", vFilePath)
    vObjProp.FireSnd = Replace$(GetFromIni(vObj, "FireSnd", vFilePath), "AppPath", App.Path)
    vObjProp.JumpSnd = Replace$(GetFromIni(vObj, "JumpSnd", vFilePath), "AppPath", App.Path)
    vObjProp.DieSnd = Replace$(GetFromIni(vObj, "DieSnd", vFilePath), "AppPath", App.Path)
    vObjProp.Raising = Val(GetFromIni(vObj, "Raising", vFilePath))
    'vObjProp.Raisetime = Val(GetFromIni(vObj, "RaiseTime", vFilePath))
    'vObjProp.JumpSnd = GetShortPath(vObjProp.JumpSnd)
    'vObjProp.DieSnd = GetShortPath(vObjProp.DieSnd)
    'vObjProp.FireSnd = GetShortPath(vObjProp.FireSnd)
    
    If vObjProp.Fixed Then
        vFixed = vFixed & vObjProp.CharType
    End If
    'Debug.Print vObjProp.CharType
    Set actualObj = vObjProp
    'Debug.Print vObjProp.SourceBottom
    colObjProp.Add vObjProp, "c" & Asc(vObjProp.CharType)
    List2.AddItem vObjProp.CharType
    
End Sub

Public Sub InitMapObjects(vObj As String, vFilePath As String)

    On Error GoTo err
    Set actualBlock = New clsBlock
    
    actualBlock.CharType = (GetFromIni(vObj, "CharType", vFilePath))
    actualBlock.CharType = Replace$(actualBlock.CharType, "Space", " ")
    actualBlock.JumpNextLevel = Val(GetFromIni(vObj, "JumpNextLevel", vFilePath))
    actualBlock.SourceLeft = Val(GetFromIni(vObj, "SourceLeft", vFilePath))
    actualBlock.SourceRight = Val(Val(GetFromIni(vObj, "SourceRight", vFilePath)))
    actualBlock.SourceTop = Val(GetFromIni(vObj, "SourceTop", vFilePath))
    actualBlock.SourceBottom = Val(GetFromIni(vObj, "SourceBottom", vFilePath))
    actualBlock.NonSolid = Val(GetFromIni(vObj, "NonSolid", vFilePath))
    actualBlock.IsFloor = Val(GetFromIni(vObj, "IsFloor", vFilePath))
    actualBlock.Fixed = Val(GetFromIni(vObj, "Fixed", vFilePath))
    actualBlock.Col = (actualBlock.SourceLeft + 1) \ 16
    actualBlock.Row = (actualBlock.SourceTop + 1) \ 16
    
    cBlocks.Add actualBlock, "c" & Asc(actualBlock.CharType)
    List1.AddItem actualBlock.CharType
    
err:
End Sub

Public Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    '*** DESCRIPTION:Reads from an *.INI file strFileName (full path & file name)
    '*** RETURNS:The string stored in [strSectionHeader], line beginning strVariableName=
    '*** NOTE: Requires declaration of API call GetPrivateProfileString
    'Initialise variable
    Dim strReturn As String
    'Blank the return string
    strReturn = String(1024, Chr(0))
    'Get requested information, trimming the
    '     returned string
    GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function

Private Sub Command7_Click()
    CommonDialog2.DialogTitle = "Saving Map File"
    CommonDialog2.FileName = MapFile
    CommonDialog2.ShowSave
    MapFile = CommonDialog2.FileName
    i = FreeFile
    Open MapFile For Binary As #i
        Put #i, , Text4.Text
    Close #i
    
    CommonDialog2.DialogTitle = "Saving Map Players File"
    CommonDialog2.FileName = MapPlayersFile
    CommonDialog2.ShowSave
    MapPlayersFile = CommonDialog2.FileName
    i = FreeFile
    Open MapPlayersFile For Binary As #i
        Put #i, , Text6.Text
    Close #i
    str1 = ""
    For i = 0 To List1.ListCount - 1
        Set actualBlock = cBlocks("c" & Asc(List1.List(i)))
        str1 = "w" & (i + 1)
        If i = 0 Then
            str2 = str1
        Else
            str2 = str2 & "," & str1
        End If
        str1 = "[" & str1 & "]" & vbNewLine
        
        str1 = str1 & "CharType = " & """" & Replace$(actualBlock.CharType, " ", "Space") & """" & vbNewLine
        str1 = str1 & "SourceLeft = " & """" & actualBlock.SourceLeft & """" & vbNewLine
        str1 = str1 & "SourceRight = " & """" & actualBlock.SourceRight & """" & vbNewLine
        str1 = str1 & "SourceTop = " & """" & actualBlock.SourceTop & """" & vbNewLine
        str1 = str1 & "SourceBottom = " & """" & actualBlock.SourceBottom & """" & vbNewLine
        str1 = str1 & "Fixed = " & """" & actualBlock.Fixed & """" & vbNewLine
        str1 = str1 & "IsFloor = " & """" & actualBlock.IsFloor & """" & vbNewLine
        str1 = str1 & "JumpNextLevel = " & """" & actualBlock.JumpNextLevel & """" & vbNewLine
        str1 = str1 & "NonSolid = " & """" & actualBlock.NonSolid & """" & vbNewLine
        str3 = str3 & str1 & vbNewLine
    Next
    str1 = "[World]" & vbNewLine
    str1 = str1 & "Map=" & """" & MapFile & """" & vbNewLine
    str1 = str1 & "Tiles=" & """" & TilesFile & """" & vbNewLine
    str1 = str1 & "MapPlayers=" & """" & MapPlayersFile & """" & vbNewLine
    str1 = str1 & "PlayersDef=" & """" & PlayersDefFile & """" & vbNewLine
    str1 = Replace$(str1, App.Path, "AppPath")
    str1 = str1 & "Objects=" & """" & str2 & """" & vbNewLine & vbNewLine & str3
    
    CommonDialog2.DialogTitle = "Saving Def file"
    CommonDialog2.FileName = DefFile
    CommonDialog2.ShowSave
    DefFile = CommonDialog2.FileName
    i = FreeFile
    Open DefFile For Binary As #i
        Put #i, , str1
    Close #i
    
End Sub

Private Sub Command8_Click()
Dim str1 As String, char1 As String
Dim iBlock As New clsBlock
    str1 = Text4.Text
    For i = 1 To Len(str1)
        char1 = Mid$(str1, i, 1)
        If char1 <> vbNewLine And char1 <> " " And _
           char1 <> vbCr And char1 <> vbLf And char1 <> "X" Then
            'Debug.Print Asc(char1)
            If IsInCol(cBlocks, char1) = char1 Then
                For Each iBlock In cBlocks
                    If iBlock.CharType = char1 Then Exit For
                Next
                If iBlock.Fixed = 0 And iBlock.IsFloor = 0 Then
                    str1 = Replace$(str1, char1, " ") 'Left$(str1, i - 1) & "X" & Mid$(str1, i + 1)
                Else
                    str1 = Replace$(str1, char1, "X") 'Left$(str1, i - 1) & "X" & Mid$(str1, i + 1)
                End If
            Else
                'str1 = Replace$(str1, char1, " ") 'Left$(str1, i - 1) & "X" & Mid$(str1, i + 1)
            End If
        End If
    Next
    'Debug.Print Len(str1)
    'Debug.Print str1
    Text6.Text = str1
End Sub

Public Function IsInCol(pObj As Collection, char As String) As String
Dim ObjProp As clsObjProp, i As Long
On Error GoTo errFound
    For i = 1 To pObj.Count
        If pObj(i).CharType = char Then
            IsInCol = char
            Exit Function
        End If
    Next
    Exit Function
errFound:
On Error GoTo 0
    IsInCol = ""
End Function

Private Sub Form_Load()
    Picture2.BackColor = vbWhite
    Picture2.Cls
End Sub

Private Sub List1_Click()
On Error Resume Next
    Set actualBlock = cBlocks("c" & Asc(List1.List(List1.ListIndex)))
    'Debug.Print actualBlock.SourceBottom
    Call Picture3.PaintPicture(Picture1, 0, 0, CInt(actualBlock.SourceRight - actualBlock.SourceLeft), CInt(actualBlock.SourceBottom - actualBlock.SourceTop), actualBlock.SourceLeft, actualBlock.SourceTop, 16, 16, vbSrcCopy)
    If actualBlock.CharType = "" Then Label3.Visible = True Else Label3.Visible = False
    'TransparentBlt Picture3.hdc, 0, 0, CInt(actualBlock.SourceRight - actualBlock.SourceLeft), CInt(actualBlock.SourceBottom - actualBlock.SourceTop), Picture1.hdc, actualBlock.SourceLeft, actualBlock.SourceTop, 16, 16, vbMagenta
    Picture3.Refresh
End Sub

Private Sub List2_Click()
On Error Resume Next
    Set actualObj = colObjProp("c" & Asc(List2.List(List2.ListIndex)))
    Picture5 = LoadPicture(actualObj.Picture)
End Sub

Private Sub Option1_Click(Index As Integer)
    Text6.Visible = False
    ShowFileByLine
    If Option1(1).Value <> False Then ShowPlayersByLine
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text5.Text = ""
    Set actualBlock = Nothing
    For Each block In cBlocks
        If block.Col = X \ 16 And block.Row = Y \ 16 Then
            Set actualBlock = block
            Text5.Text = actualBlock.CharType
            Exit For
        End If
    Next
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = actualBlock.CharType Then
            List1.ListIndex = i
            Exit For
        End If
    Next
    If actualBlock.CharType = "" Then Label3.Visible = True Else Label3.Visible = False
    If Text5.Text = "" Then
        Set actualBlock = New clsBlock
        actualBlock.CharType = ""
        actualBlock.Col = Val(X \ 16)
        actualBlock.Row = Val(Y \ 16)
        actualBlock.SourceLeft = Val(X \ 16) * 16
        actualBlock.SourceRight = actualBlock.SourceLeft + 16
        actualBlock.SourceTop = Val(Y \ 16) * 16
        actualBlock.SourceBottom = actualBlock.SourceTop + 16
        Call Picture3.PaintPicture(Picture1, 0, 0, CInt(actualBlock.SourceRight - actualBlock.SourceLeft), CInt(actualBlock.SourceBottom - actualBlock.SourceTop), actualBlock.SourceLeft, actualBlock.SourceTop, 16, 16, vbSrcCopy)
        'TransparentBlt Picture3.hdc, 0, 0, CInt(actualBlock.SourceRight - actualBlock.SourceLeft), CInt(actualBlock.SourceBottom - actualBlock.SourceTop), Picture1.hdc, actualBlock.SourceLeft, actualBlock.SourceTop, 16, 16, vbMagenta
        Picture3.Refresh
    End If
    
    Text3.Text = "SourceLeft=" & """" & actualBlock.SourceLeft & """" & _
             Chr$(13) & "SourceRight=" & """" & actualBlock.SourceRight & """" & _
             Chr$(13) & "SourceTop=" & """" & actualBlock.SourceTop & """" & _
             Chr$(13) & "SourceBottom=" & """" & actualBlock.SourceBottom & """"
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture1.Picture = LoadPicture(Data.Files.Item(1))
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Text4 = "" Then
        For i = 0 To Picture2.Height \ 16 + 1
            str1 = str1 & Space$(Picture2.Width \ 16 + 1) & vbNewLine
        Next
        Text4 = str1
    End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim t As clsBlock, arrtext() As String
On Error GoTo err:
    If Button = vbLeftButton Then
        If Option1(0).Value = True Then
            If actualBlock Is Nothing Then Exit Sub
            If actualBlock.CharType = "" Then Exit Sub
'    Set t = cBlocks("c" & Asc(actualBlock.CharType))
            'Call Picture2.PaintPicture(Picture1, CInt((x \ 16) * 16), CInt((y \ 16) * 16), CInt(actualBlock.SourceRight - actualBlock.SourceLeft), CInt(actualBlock.SourceBottom - actualBlock.SourceTop), actualBlock.SourceLeft, actualBlock.SourceTop, 16, 16, vbSrcCopy)
            TransparentBlt Picture2.hdc, CInt((X \ 16) * 16), CInt((Y \ 16) * 16), CInt(actualBlock.SourceRight - actualBlock.SourceLeft), CInt(actualBlock.SourceBottom - actualBlock.SourceTop), Picture1.hdc, actualBlock.SourceLeft, actualBlock.SourceTop, 16, 16, vbMagenta
            Picture2.Refresh
            Replace Text4.Text, Chr$(10), ""
            arrtext = Split(Text4.Text, Chr$(13))
            str1 = ""
            X = X + 16
            For i = 0 To UBound(arrtext)
                arrtext(i) = Replace$(arrtext(i), Chr(10), "")
                arrtext(i) = Replace$(arrtext(i), Chr(13), "")
                If i > 0 Then
                    str1 = str1 & vbNewLine
                End If
                If i = (Y \ 16) Then
                    str1 = str1 & Left$(arrtext(i), (X \ 16) - 1) & actualBlock.CharType & Mid$(arrtext(i), (X \ 16) + 1)
                Else
                    str1 = str1 & arrtext(i)
                End If
            Next
            Text4.Text = str1
        Else
            If actualObj Is Nothing Then Exit Sub
            If actualObj.CharType = "" Then Exit Sub
            'Call Picture2.PaintPicture(Picture5, 16 + CInt((x \ 16) * 16) - CInt(actualObj.SourceRight - actualObj.SourceLeft), 16 + CInt((y \ 16) * 16) - CInt(actualObj.SourceBottom - actualObj.SourceTop), , , actualObj.SourceLeft, actualObj.SourceTop, CInt(actualObj.SourceRight - actualObj.SourceLeft), CInt(actualObj.SourceBottom - actualObj.SourceTop), vbSrcCopy)
            TransparentBlt Picture2.hdc, 16 + CInt((X \ 16) * 16) - CInt(actualObj.SourceRight - actualObj.SourceLeft), 16 + CInt((Y \ 16) * 16) - CInt(actualObj.SourceBottom - actualObj.SourceTop), CInt(actualObj.SourceRight - actualObj.SourceLeft), CInt(actualObj.SourceBottom - actualObj.SourceTop), Picture5.hdc, actualObj.SourceLeft, actualObj.SourceTop, CInt(actualObj.SourceRight - actualObj.SourceLeft), CInt(actualObj.SourceBottom - actualObj.SourceTop), vbMagenta
            Picture2.Refresh
            Replace Text6.Text, Chr$(10), ""
            arrtext = Split(Text6.Text, Chr$(13))
            str1 = ""
            X = X + 16
            For i = 0 To UBound(arrtext)
                arrtext(i) = Replace$(arrtext(i), Chr(10), "")
                arrtext(i) = Replace$(arrtext(i), Chr(13), "")
                If i > 0 Then
                    str1 = str1 & vbNewLine
                End If
                If i = (Y \ 16) Then
                    str1 = str1 & Left$(arrtext(i), (X \ 16) - 1) & actualObj.CharType & Mid$(arrtext(i), (X \ 16) + 1)
                Else
                    str1 = str1 & arrtext(i)
                End If
            Next
            Text6.Text = str1
        End If
    End If
err:
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim t As clsBlock, arrtext() As String
On Error GoTo err:
    
    If Option1(0).Value = True Then
        If actualBlock Is Nothing Then Exit Sub
        If actualBlock.CharType = "" Then Exit Sub
        Set t = cBlocks("c" & Asc(actualBlock.CharType))
        '¡Call Picture2.PaintPicture(Picture1, CInt((x \ 16) * 16), CInt((y \ 16) * 16), CInt(actualBlock.SourceRight - actualBlock.SourceLeft), CInt(actualBlock.SourceBottom - actualBlock.SourceTop), actualBlock.SourceLeft, actualBlock.SourceTop, 16, 16, vbSrcCopy)
        TransparentBlt Picture2.hdc, CInt((X \ 16) * 16), CInt((Y \ 16) * 16), CInt(actualBlock.SourceRight - actualBlock.SourceLeft), CInt(actualBlock.SourceBottom - actualBlock.SourceTop), Picture1.hdc, actualBlock.SourceLeft, actualBlock.SourceTop, 16, 16, vbMagenta
        Picture2.Refresh
        Replace Text4.Text, Chr$(10), ""
        arrtext = Split(Text4.Text, Chr$(13))
        str1 = ""
        X = X + 16
        For i = 0 To UBound(arrtext)
            arrtext(i) = Replace$(arrtext(i), Chr(10), "")
            arrtext(i) = Replace$(arrtext(i), Chr(13), "")
            If i > 0 Then
                str1 = str1 & vbNewLine
            End If
            If i = (Y \ 16) Then
                str1 = str1 & Left$(arrtext(i), (X \ 16) - 1) & actualBlock.CharType & Mid$(arrtext(i), (X \ 16) + 1)
            Else
                str1 = str1 & arrtext(i)
            End If
        Next
        Text4.Text = str1
    Else
        If actualObj Is Nothing Then Exit Sub
        If actualObj.CharType = "" Then Exit Sub
        Set t = colObjProp("c" & Asc(actualObj.CharType))
        'Call Picture2.PaintPicture(Picture5, CInt((x \ 16) * 16), CInt((y \ 16) * 16), CInt(actualObj.SourceRight - actualObj.SourceLeft), CInt(actualObj.SourceBottom - actualObj.SourceTop), actualObj.SourceLeft, actualObj.SourceTop, 16, 16, vbSrcCopy)
        TransparentBlt Picture2.hdc, CInt((X \ 16) * 16), CInt((Y \ 16) * 16), CInt(actualObj.SourceRight - actualObj.SourceLeft), CInt(actualObj.SourceBottom - actualObj.SourceTop), Picture5.hdc, actualObj.SourceLeft, actualObj.SourceTop, 16, 16, vbMagenta
        Picture2.Refresh
        Replace Text6.Text, Chr$(10), ""
        arrtext = Split(Text6.Text, Chr$(13))
        str1 = ""
        X = X + 16
        For i = 0 To UBound(arrtext)
            arrtext(i) = Replace$(arrtext(i), Chr(10), "")
            arrtext(i) = Replace$(arrtext(i), Chr(13), "")
            If i > 0 Then
                str1 = str1 & vbNewLine
            End If
            If i = (Y \ 16) Then
                str1 = str1 & Left$(arrtext(i), (X \ 16) - 1) & actualObj.CharType & Mid$(arrtext(i), (X \ 16) + 1)
            Else
                str1 = str1 & arrtext(i)
            End If
        Next
        Text6.Text = str1
    End If
err:
End Sub

Private Sub ReadMap_Click()
CommonDialog2.ShowOpen
MapFile = CommonDialog2.FileName
ReadFileByLine MapFile
End Sub

Private Function ReadFileByLine(vFile As String) As String
'On Error GoTo errHandler::
X = FreeFile
On Error Resume Next
Open vFile For Input As #X
    Do While Not EOF(1)
        Line Input #X, txtVar
        If str1 = "" Then
            str1 = str1 & txtVar
        Else
            str1 = str1 & vbNewLine & txtVar
        End If
    Loop
Close #X
Text4.Text = str1
ShowFileByLine
End Function

Private Function ShowFileByLine()
Dim txtVar As String, i As Long
On Error Resume Next
arrlines = Split(Text4.Text, vbNewLine)
For i = 0 To UBound(arrlines)
    If vWidth < Len(arrlines(i)) Then vWidth = Len(arrlines(i))
Next
Picture2.Width = (vWidth * 16) + 4
Picture2.Height = ((1 + UBound(arrlines)) * 16) + 4
Picture2.Cls
ScrollBars1.ReSetScrollBar Picture4.hwnd
ScrollBars1.HorizMaxVal = Abs(Picture2.ScaleWidth - Picture4.ScaleWidth + 16)
ScrollBars1.VertMaxVal = Abs(Picture2.ScaleHeight - Picture4.ScaleHeight + 16)
ScrollBars1.AdjustScrollInfo Picture4.hwnd
DoEvents
Set t = cBlocks("c" & Asc(" "))
TransparentBlt Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hdc, t.SourceLeft, t.SourceTop, 16, 16, vbMagenta
Picture2.Refresh
'Debug.Print (cBlocks Is Nothing)
    For i = 0 To UBound(arrlines)
        For j = 1 To Len(arrlines(i))
            char1 = Mid$(arrlines(i), j, 1)
            Set t = Nothing
            Set t = cBlocks("c" & Asc(char1))
            If t Is Nothing Then
                char1 = " "
                Set t = cBlocks("c" & Asc(char1))
            End If
            'Call Picture2.PaintPicture(Picture1, CInt((j - 1) * 16), CInt(i * 16), CInt(t.SourceRight - t.SourceLeft), CInt(t.SourceBottom - t.SourceTop), t.SourceLeft, t.SourceTop, 16, 16, vbSrcCopy)
            TransparentBlt Picture2.hdc, CInt((j - 1) * 16), CInt(i * 16), CInt(t.SourceRight - t.SourceLeft), CInt(t.SourceBottom - t.SourceTop), Picture1.hdc, t.SourceLeft, t.SourceTop, 16, 16, vbMagenta
        Next
    Next
    Picture2.Refresh
Exit Function
errHandler:
    MsgBox err.Description
End Function

Private Function ReadPlayersByLine(vFile As String) As String
'On Error GoTo errHandler::
X = FreeFile
On Error Resume Next
Open vFile For Input As #X
    Do While Not EOF(X)
        Line Input #X, txtVar
        If str1 = "" Then
            str1 = str1 & txtVar
        Else
            str1 = str1 & vbNewLine & txtVar
        End If
    Loop
Close #X
Text6.Text = str1
ShowPlayersByLine
End Function

Private Function ShowPlayersByLine()
Dim txtVar As String, i As Long

On Error Resume Next
Text6.Visible = True
arrlines = Split(Text6.Text, vbNewLine)
For i = 0 To UBound(arrlines)
    If vWidth < Len(arrlines(i)) Then vWidth = Len(arrlines(i))
Next
'Picture2.Width = (vWidth * 16) + 4
'Picture2.Height = (UBound(arrlines) * 16) + 4
'Picture2.Cls
'ScrollBars1.ReSetScrollBar Picture4.hwnd
'ScrollBars1.HorizMaxVal = Abs(Picture2.ScaleWidth - Picture4.ScaleWidth + 16)
'ScrollBars1.VertMaxVal = Abs(Picture2.ScaleHeight - Picture4.ScaleHeight + 16)
'ScrollBars1.AdjustScrollInfo Picture4.hwnd
    For i = 0 To UBound(arrlines)
        For j = 1 To Len(arrlines(i))
            char1 = Mid$(arrlines(i), j, 1)
            Set t = Nothing
            Set t = colObjProp("c" & Asc(char1))
            If Not t Is Nothing Then
                Picture5 = LoadPicture(t.Picture)
                DoEvents
                'Debug.Print CInt(t.SourceBottom - t.SourceTop)
                'Debug.Print t.SourceBottom
                'Call Picture2.PaintPicture(Picture5, CInt((j - 1) * 16), CInt(i * 16), CInt(t.SourceRight - t.SourceLeft), CInt(t.SourceBottom - t.SourceTop), t.SourceLeft, t.SourceTop, t.SourceRight, t.SourceBottom, vbSrcCopy)
                'Call Picture2.PaintPicture(Picture5, 16 + CInt((j - 1) * 16) - CInt(t.SourceRight - t.SourceLeft), 16 + CInt(i * 16) - CInt(t.SourceBottom - t.SourceTop), , , t.SourceLeft, t.SourceTop, CInt(t.SourceRight - t.SourceLeft), CInt(t.SourceBottom - t.SourceTop), vbSrcCopy)
                TransparentBlt Picture2.hdc, 16 + CInt((j - 1) * 16) - CInt(t.SourceRight - t.SourceLeft), 16 + CInt(i * 16) - CInt(t.SourceBottom - t.SourceTop), CInt(t.SourceRight - t.SourceLeft), CInt(t.SourceBottom - t.SourceTop), Picture5.hdc, t.SourceLeft, t.SourceTop, CInt(t.SourceRight - t.SourceLeft), CInt(t.SourceBottom - t.SourceTop), vbMagenta
            End If
        Next
    Next
    Picture2.Refresh
    List2.ListIndex = 0
    List2_Click
Exit Function
errHandler:
    MsgBox err.Description
End Function


