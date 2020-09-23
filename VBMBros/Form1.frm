VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2685
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4335
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTerrain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   2040
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox pForm1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   3840
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.PictureBox picGoomba 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   3000
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   144
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox picTiles 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3660
      Left            =   960
      Picture         =   "Form1.frx":1B42
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox PicMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      Height          =   1095
      Left            =   3360
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox PicMapMask 
      AutoRedraw      =   -1  'True
      Height          =   1095
      Left            =   2280
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cScreenWidth  As Long
Public cScreenHeight As Long

Const cKeyRight = vbKeyNumpad6
Const cKeyLeft = vbKeyNumpad4
Const cKeyDown = vbKeyNumpad2
Const cKeyJump = vbKeySpace
Const cKeyFire = vbKeyControl
Const cKeySpeed = vbKeyZ

Public Player As clsObject
Dim col_oldPlayers As Collection
Dim colLevels As Collection
Dim Enemy As clsObject
Dim oTimer As clsTimer
Dim oScreen As clsScreen
Dim m_exit As Boolean
Dim m_passlevel As Boolean
Dim mKeyRight As Integer
Dim mKeyLeft As Integer
Dim mKeyDown As Integer
Public mKeyJump As Integer
Public mKeyFire As Integer
Public mKeySpeed As Integer
Dim mKey As Integer
Dim mStopMotion As Boolean
Dim mEnableJump As Boolean
Dim mVel As Long
Dim acum As Long

Private Sub GameStart()
Dim bot As clsObject, i As Long, tPosLeft As Long, tPosTop As Long
Dim cbBlock As clsbBlock, j As Long, vObjProp As clsObjProp, x As Long
Dim tmpCol As Collection, BrkFx As Variant, vVel As Long, vProm As Long

Set tmpCol = New Collection
For Each vObjProp In colObjProp
    If vObjProp.UserSelection Then
        Set bot = New clsObject
        MoveProp vObjProp, bot
        bot.PosLeft = (ScreenWidth - Abs(bot.PosRight - bot.PosLeft)) \ 2: bot.PosRight = (ScreenWidth + Abs(bot.PosRight - bot.PosLeft)) \ 2: bot.PosTop = (ScreenHeight - Abs(bot.PosBottom - bot.PosTop)) \ 2: bot.PosBottom = (ScreenHeight + Abs(bot.PosBottom - bot.PosTop)) \ 2
        i = i + 1
        If i = 1 Then
            bot.User = True
        Else
            bot.User = False
        End If
        tmpCol.Add bot
    End If
Next

Set bot = Nothing
j = 1
On Error GoTo err1
Set bot = tmpCol(j)
bot.User = True

Do While mKey <> vbKeyReturn And mKey <> vbKeyEscape And True = False
    Do While cGetInputState() <> 0
        DoEvents
    Loop
    If mKeyRight <> 0 And j < tmpCol.Count Then
        j = j + 1
        bot.User = False
        Set bot = tmpCol(j)
        bot.User = True
    End If
    If mKeyLeft <> 0 And j > 1 Then
        j = j - 1
        bot.User = False
        Set bot = tmpCol(j)
        bot.User = True
    End If
    
    pForm1.Cls
    bot.Draw 0
    rctDst.Left = 0: rctDst.Right = ScreenWidth: rctDst.Top = 0: rctDst.Bottom = ScreenHeight
    rctSrc.Left = 0: rctSrc.Right = ScreenWidth: rctSrc.Top = 0: rctSrc.Bottom = ScreenHeight
    Paint Form1.hdc, rctDst, pForm1.hdc, rctSrc, vbSrcCopy
    mKeyRight = 0
    mKeyLeft = 0
Loop

For x = tmpCol.Count To 1 Step -1
    Set bot = tmpCol(x)
    If bot.User Then
        For i = 1 To cLevel.VBlocks
            For j = 1 To cLevel.HBlocks
                If cLevel.PBlock(j, i) = "p" Then
                    For Each vObjProp In colObjProp
                        If vObjProp.CharType = bot.CharType Then Exit For
                    Next
                    MoveProp vObjProp, Player
                    Player.PosLeft = (j - 1) * 16
                    Player.PosTop = (i - 2) * 16
                    Player.PosRight = Player.PosLeft + 16
                    Player.PosBottom = Player.PosTop + 32
                    Player.NextLevelChar = gNextlevelChar
                    KeyIdx = KeyIdx + 1
                    Player.Key = Player.CharType & KeyIdx
                    Col.Add Player, Player.Key
                    Player.RScreenLimit = ScreenWidth / 2
                    Set vObjProp = Nothing
                End If
            Next
        Next
    End If
    tmpCol.Remove x
Next

Set tmpCol = Nothing

If LenB(vGameMusic) <> 0 Then
    PlayMIDI vGameMusic
End If

Set cBrkFx = New Collection
Dim Tim As Long
oTimer.CaptureFrames = 10
Form1.Cls
Do While m_exit = False
    Do While cGetInputState() <> 0
        DoEvents
    Loop
' Si se tardo demasiado darle aire a las demas aplicaciones
    If oTimer.TimePerFrame() > 50 Then
        DoEvents
    End If
' si se realizo todo en menos de 65fps (0.015 sec p/frame) esperar
    'Debug.Print "-"
    Tim = oTimer.TimePerFrame()
    Do While Tim < 17
      Tim = oTimer.TimePerFrame()
      DoEvents
    Loop
    
    If oTimer.fps > 0 Then
        Form1.Caption = oTimer.fps & "fps"
    End If
    
    oTimer.StartFrame
    
    cLevel.Redraw cLevel.OffsetX, 0, cLevel.OffsetX + ScreenWidth, ScreenHeight, _
               0, 0, ScreenWidth, ScreenHeight
'    For i = 1 To Col.Count
'        If i <= Col.Count Then
'            Set bot = Col(i)
     For Each bot In Col
            'Debug.Print bot.CharType
            If bot.Visible Then
                'If bot.User Then
                '    Debug.Print bot.CharType
                'End If
                If bot.Amnisty > 0 Then
                    bot.Amnisty = bot.Amnisty - oTimer.ElapsedTime
                ElseIf bot.Amnisty < 0 Then
                    bot.Amnisty = 0
                End If
                If bot.CheckIfInScreen(cLevel.OffsetX, ScreenWidth) Then
                   'If bot.CharType = "M" Then
                   '     Debug.Print bot.PosBottom
                   ' End If
                    If bot.InitedAI And Not bot.Died Then
                        'Debug.Print bot.CharType
                        bot.PerformAI
                    End If
                    'If bot.CharType = "M" Then
                    '    Debug.Print bot.PosBottom
                    'End If
                    'If bot.CharType = "R" Then
                    Call bot.Animate(Col, oTimer.ElapsedTime)
                    bot.Draw (cLevel.OffsetX)
                End If
                
            Else
                If bot.Died Then
                    Col.Remove bot.Key
                    Set bot = Nothing
                End If
            End If
        'End If
    Next
    
    For Each cbBlock In cBreakeable
        If cbBlock.Visible Then
            If cbBlock.StartAnim Then
                rctDst.Left = cbBlock.PosLeft: rctDst.Right = cbBlock.PosRight: rctDst.Top = cbBlock.PosTop: rctDst.Bottom = cbBlock.PosBottom
                rctSrc.Left = cbBlock.PosLeft: rctSrc.Right = cbBlock.PosRight: rctSrc.Top = cbBlock.PosTop: rctSrc.Bottom = cbBlock.PosBottom
                Paint PicMap.hdc, rctDst, PicMapMask.hdc, rctSrc, vbSrcCopy
                'Paint picTerrain.hdc, rctDst, PicMapMask.hdc, rctSrc, vbSrcCopy
                Call cbBlock.Draw(1, cLevel.OffsetX)
            End If
        End If
    Next
    
    rctDst.Left = 0: rctDst.Right = cScreenWidth: rctDst.Top = 0: rctDst.Bottom = cScreenHeight
    rctSrc.Left = 0: rctSrc.Right = ScreenWidth: rctSrc.Top = 0: rctSrc.Bottom = ScreenHeight
    Paint Form1.hdc, rctDst, pForm1.hdc, rctSrc, vbSrcCopy
    
    Select Case mKey
        Case Is = -1
            Player.StopMotion
            If Player.PosLeft > (ScreenWidth \ 2) Then
                cLevel.OffsetX = Player.PosLeft - (ScreenWidth \ 2)
            End If
            mKey = 0
        Case Is = vbKeyF3
            Player.JumpNextLevel = True
            mKey = 0
        Case Is = vbKeyEscape
            m_exit = True
        Case Is = vbKeyF8
            If oScreen.ResChanged Then
                Form1.Cls
                Call oScreen.ReturnRes
            Else
                Form1.Cls
                cScreenWidth = 320: cScreenHeight = 240
                If oScreen.ChangeRes(cScreenWidth, cScreenHeight) <> 0 Then
                    cScreenWidth = 640: cScreenHeight = 480
                    Call oScreen.ChangeRes(cScreenWidth, cScreenHeight)
                End If
            End If
            mKey = 0
        Case Is = vbKeyF9
            m_exit = True
            mKey = 0
    End Select
    
    If Not Player.Died Or Player.JumpNextLevel Then
        If Player.CanCrouch And Player.OnFloor(Player.PosBottom) Then Player.Crouched = mKeyDown
        If Not Player.Crouched Then
            If mKeyRight <> 0 Or (Player.direction And mVel <> 0 And mKeyLeft = 0) Then
                'If mVel < 0 Then mVel = mVel * -1
                If mVel < 10 And mKeyRight <> 0 Then
                    mVel = mVel + 1
                ElseIf mVel < 20 And mKeySpeed <> 0 And mKeyRight <> 0 Then
                    mVel = mVel + 2
                ElseIf mKeyRight = 0 Then
                    If mVel > 0 Then
                        mVel = mVel - 1
                    ElseIf mVel < 0 Then
                        mVel = mVel + 1
                    End If
                End If
                If mVel > 10 And mKeySpeed = 0 Then mVel = 10
                vProm = Player.Velocity * mVel
                If Abs(vProm) < 10 Then
                    vVel = Sgn(vProm)
                Else
                    vVel = (vProm / 10) \ 1
                End If
                If (mKeyRight = 0 And mVel <> 0) Or mVel < 0 Then
                    If mKeyRight = 0 And vVel = 0 Then
                    Else
                    If vVel < 0 Then
                        Player.MoveLeft oTimer.ElapsedTime, True, vVel
                    ElseIf vVel > 0 Then
                        Player.MoveRight oTimer.ElapsedTime, True, vVel
                    End If
                    'If mKeyLeft <> 0 And
                    If mKeySpeed <> 0 And vVel <= 0 And -vVel > Player.Velocity / 3 Then Player.RunFrame = Player.ChangeFrom
                    End If
                Else
                    'If vVel = 0 Then vVel = 1
                    If vVel < Player.Velocity And vVel > 0 Then
                        Player.MoveRight oTimer.ElapsedTime, True, Abs(vVel)
                    Else
                        If mKeySpeed <> 0 Then
                            Player.MoveRight oTimer.ElapsedTime, True, Player.Velocity + 1
                        Else
                            Player.MoveRight oTimer.ElapsedTime, True
                        End If
                    End If
                End If
            ElseIf (mKeyLeft <> 0 Or (Not Player.direction And mVel <> 0)) And Player.PosLeft > cLevel.OffsetX Then
                'If mVel > 0 Then mVel = mVel * -1
                If mVel > -10 And mKeyLeft <> 0 Then
                    mVel = mVel - 1
                ElseIf mVel > -20 And mKeySpeed <> 0 And mKeyLeft <> 0 Then
                    mVel = mVel - 2
                ElseIf mKeyLeft = 0 Then
                    If mVel > 0 Then
                        mVel = mVel - 1
                    ElseIf mVel < 0 Then
                        mVel = mVel + 1
                    End If
                End If
                If mVel < -10 And mKeySpeed = 0 Then mVel = -10
                vProm = Player.Velocity * mVel
                If Abs(vProm) < 10 Then
                    vVel = Sgn(vProm)
                Else
                    vVel = (vProm / 10) \ 1
                End If
                If (mKeyLeft = 0 And mVel <> 0) Or mVel > 0 Then
                    If mKeyRight = 0 And vVel = 0 Then
                    Else
                    If vVel > 0 Then
                        Player.MoveRight oTimer.ElapsedTime, True, vVel
                    ElseIf vVel < 0 Then
                        Player.MoveLeft oTimer.ElapsedTime, True, vVel
                    End If
                    'If mKeyRight <> 0 And
                    If mKeySpeed <> 0 And vVel > 0 And vVel > Player.Velocity / 3 Then Player.RunFrame = Player.ChangeFrom
                    End If
                Else
                    'If vVel > 0 Then vVel = -1
                    If Abs(vVel) < Player.Velocity And mVel < 0 Then
                        'If vVel <> 0 Then
                        Player.MoveLeft oTimer.ElapsedTime, True, -Abs(vVel)
                    Else
                        If mKeySpeed <> 0 Then
                            Player.MoveLeft oTimer.ElapsedTime, True, -Player.Velocity - 1
                        Else
                            Player.MoveLeft oTimer.ElapsedTime, True
                        End If
                    End If
                End If
            Else
                mVel = 0
                vVel = 0
                Player.RunFrame = Player.AnimFrom
            'If Player.Inercia <> 0 Then
                'If Player.Inercia > 0 Then
                '    Player.MoveRight oTimer.ElapsedTime, False
                'Else
                '    Player.MoveLeft oTimer.ElapsedTime, False
                'End If
            End If
        
            If mKeyRight <> 0 Then Player.direction = True
            If mKeyLeft <> 0 Then Player.direction = False
            If Player.PosLeft > (ScreenWidth \ 2) Then
                cLevel.OffsetX = Player.PosLeft - (ScreenWidth \ 2)
            End If
            If mKeyJump <> 0 And Not Player.Falling Then
                Player.Jump = True
                mEnableJump = False
            Else
                Player.Jump = False
            End If
            
            If Player.ForceJump Then
                Player.Jump = True
                Player.ForceJump = False
            End If
            If mKeyFire > 0 And LenB(Player.FireBall) Then
                Player.Fire
                mKeyFire = -1
            End If
        End If
        If Player.JumpNextLevel Then
            m_exit = True
            Player.JumpNextLevel = False
            m_passlevel = True
        End If
        
        If Player.JumpStart Then
            Player.JumpStart = False
        End If
    Else
        If Not Player.Visible Then m_exit = True
    End If
Loop

For Each bot In Col
    Col.Remove bot.Key
Next

KeyIdx = 0

For Each BrkFx In cBrkFx
    Call mciSendString("Close " & BrkFx, 0&, 0, 0)
    cBrkFx.Remove BrkFx
Next

Set cBrkFx = Nothing

If LenB(vGameMusic) <> 0 Then
    StopMIDI vGameMusic
End If
err1:
End Sub

Private Sub main()
Dim vVar As String, arrWorlds() As String, arrObjects() As String
Dim i As Long, tmpLevel As clsLevel

vVar = GetFromIni("Game", "WorldDef", App.Path & "\Game.ini")
vSwitchGifts = GetFromIni("Game", "SwitchGifts", App.Path & "\Game.ini")
arrWorlds = Split(vVar, ",")

For i = 0 To UBound(arrWorlds)
    Set tmpLevel = New clsLevel
    tmpLevel.Def = Replace$(arrWorlds(i), "AppPath", App.Path)
    'Debug.Print tmpLevel.Def
    colLevels.Add tmpLevel
Next

Set tmpLevel = Nothing

Me.Visible = True

Do While mKey <> vbKeyEscape And colLevels.Count > 0
'Load the Current Level
    m_passlevel = False
    Set cLevel = colLevels(1)
    Run_Level cLevel
'Start a New game
    m_exit = False
    GameStart
'Unload the Current Level
    Unload_Objects
    If m_passlevel Then
'Unload previus Level
         colLevels.Remove 1
    End If
    DoEvents
Loop

For i = colLevels.Count To 1 Step -1
    colLevels.Remove (i)
Next
Set colLevels = Nothing

For i = colObjProp.Count To 1 Step -1
    colObjProp.Remove (i)
Next
Set colObjProp = Nothing

For i = picGoomba.Count - 1 To 1 Step -1
    Set picGoomba(i).Picture = LoadPicture("")
    Unload picGoomba(i)
Next

If oScreen.ResChanged Then
    Call oScreen.ReturnRes
End If
Set oScreen = Nothing
Unload Me
End Sub

Private Sub Run_Level(pLevel As clsLevel)
    Dim i As Long, j As Long, vVar As String, arrObjects() As String
    Dim strDef As String, x As Long, y As Long
    
    Set Col = New Collection
    Set col_oldPlayers = New Collection
    Set oTimer = New clsTimer
    Set Player = New clsObject
    Set cBreakeable = New Collection
    
    m_mult = 1
    
' Clear everything except the path to de definition
    strDef = pLevel.Def
    Set pLevel = Nothing
    Set pLevel = New clsLevel
    pLevel.Def = strDef
' Load the definition
    gNextlevelChar = pLevel.LoadMap()
' -------------------------------------------------------------------
' Load Game Characters
' -------------------------------------------------------------------
  
    vVar = GetFromIni("Game", "Objects", pLevel.PlayersDef)
    arrObjects = Split(vVar, ",")
    
    Set colObjProp = New Collection
    
    For i = 0 To UBound(arrObjects)
        Call InitObjects(arrObjects(i), pLevel.PlayersDef)
    Next
    
    Erase arrObjects
    
    For i = 1 To pLevel.VBlocks
        For j = 1 To pLevel.HBlocks
            'Debug.Print pLevel.PBlock(j, i)
            Select Case pLevel.PBlock(j, i)
                Case Is = " "
                Case Is = "X"
                Case Is = IsInCol(colObjProp, pLevel.PBlock(j, i))
                    x = x + 1
                    Set Enemy = New clsObject
                    On Error GoTo errFounded
                    For Each vObjProp In colObjProp
                        If vObjProp.CharType = pLevel.PBlock(j, i) Then Exit For
                    Next
                    If Not vObjProp Is Nothing Then
                     On Error GoTo 0
                        If vObjProp.UserSelection = False Then
                            MoveProp vObjProp, Enemy
                            Enemy.PosLeft = (j - 1) * 16 + Enemy.StartPosOffsetX
                            Enemy.PosTop = (i - 1) * 16
                            Enemy.PosRight = Enemy.PosLeft + (Enemy.SourceRight - Enemy.SourceLeft)
                            Enemy.PosBottom = Enemy.PosTop + (Enemy.SourceBottom - Enemy.SourceTop)
                            Enemy.ID = Col.Count + 1
                            If LenB(Enemy.AI) = 0 Then
                                Enemy.InitedAI = False
                            End If
                            KeyIdx = KeyIdx + 1
                            Enemy.Key = Enemy.CharType & KeyIdx
                            Enemy.RunFrame = Enemy.AnimFrom
                            Col.Add Enemy, Enemy.Key
                        End If
                    End If
                    Set vObjProp = Nothing
            End Select
        Next
    Next
    Set vObjProp = Nothing
    Set Enemy = Nothing
    
    Me.ZOrder 0
    m_exit = True
    Unload frmMenu
    Exit Sub
errFounded:
    MsgBox "Object not founded Check game.ini!!"
    Set vObjProp = Nothing
    Set Enemy = Nothing
    Unload frmMenu
End Sub

Public Sub Start()
    frmMenu.Visible = False
    DoEvents
    
    If m_exit = False Then Exit Sub
    
    Set colLevels = New Collection
    
    gMusicOn = frmMenu.mnuMusicOn.Checked
    gEffectsOn = frmMenu.mnuEffectsOn.Checked
    
    If oScreen Is Nothing Then Set oScreen = New clsScreen
    
    If frmMenu.mnuFullScreen.Checked Then
        cScreenWidth = 320: cScreenHeight = 240
        If oScreen.ChangeRes(cScreenWidth, cScreenHeight) <> 0 Then
            cScreenWidth = 640: cScreenHeight = 480
            Call oScreen.ChangeRes(cScreenWidth, cScreenHeight)
        End If
        DoEvents
    End If
    
    Call main
End Sub

Private Sub Form_Initialize()
    m_exit = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    mKey = KeyCode
    Select Case KeyCode
        Case cKeyRight
            mKeyRight = KeyCode
        Case cKeyLeft
            mKeyLeft = KeyCode
        Case cKeyDown
            mKeyDown = KeyCode
        Case cKeyJump
            mKeyJump = KeyCode
        Case cKeyFire
            If mKeyFire = 0 Then mKeyFire = KeyCode
        Case cKeySpeed
            mKeySpeed = KeyCode
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case cKeyRight
            mKeyRight = 0
            If mKey = cKeyRight Or mKey = cKeyLeft Then mKey = -1
        Case cKeyLeft
            mKeyLeft = 0
            If mKey = cKeyRight Or mKey = cKeyLeft Then mKey = -1
        Case cKeyDown
            mKeyDown = 0
        Case cKeyJump
            'If mKey = cKeyRight Or mKey = cKeyLeft Then mKey = -1
            mKeyJump = 0
            'mEnableJump = True
            'Player.Jump = False
        Case cKeyFire
            mKeyFire = 0
        Case cKeySpeed
            mKeySpeed = 0
    End Select
End Sub

Private Sub Form_Load()
    Form1.Width = (ScreenWidth + (GetSystemMetrics(SM_CXDLGFRAME) * 2)) * Screen.TwipsPerPixelX
    Form1.Height = (ScreenHeight + GetSystemMetrics(SM_CXDLGFRAME) * 2 + GetSystemMetrics(SM_CYCAPTION)) * Screen.TwipsPerPixelY
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If m_exit = False Then
        m_exit = True
        Cancel = True
    End If
End Sub

Public Function GetBlock(PosX As Long, PosY As Long) As String
    If PosX > cLevel.HBlocks Then Exit Function
    If PosY > cLevel.VBlocks Then GetBlock = "-1": Exit Function
    If PosY < 1 Then Exit Function
    If PosX < 1 Then PosX = 1
    GetBlock = cLevel.Block(PosX, PosY)
End Function

Public Function RollBackPlayer() As Boolean
Dim oldPlayer As clsObject
    If col_oldPlayers.Count > 0 Then
        Set oldPlayer = col_oldPlayers.Item(col_oldPlayers.Count)
        'Debug.Print Player.CharType
        'Debug.Print oldPlayer.CharType
        oldPlayer.PosLeft = Player.PosLeft
        oldPlayer.PosBottom = Player.PosBottom
        oldPlayer.PosRight = Player.PosRight
        oldPlayer.PosTop = Player.PosTop
        oldPlayer.StopMotion
        oldPlayer.direction = Not Player.direction
        oldPlayer.ForceJump = True
        oldPlayer.Jump = True
        oldPlayer.Amnisty = Player.Amnisty
        Col.Remove Player.Key
        Set Player = oldPlayer
        col_oldPlayers.Remove col_oldPlayers.Count
        KeyIdx = KeyIdx + 1
        Player.Key = Player.CharType & KeyIdx
        Col.Add Player, Player.Key
        RollBackPlayer = True
    End If
End Function

Public Sub SetPlayer(char As String, Optional p_addtocollection = True)
Dim newPlayer As clsObject, vObjProp As clsObjProp
    
    If LenB(char) = 0 Then Exit Sub
    Set newPlayer = New clsObject
    For Each vObjProp In colObjProp
        If vObjProp.CharType = char Then Exit For
    Next
    MoveProp vObjProp, newPlayer
    newPlayer.NextLevelChar = Player.NextLevelChar
    newPlayer.PosLeft = Player.PosLeft
    newPlayer.PosBottom = Player.PosBottom
    newPlayer.PosTop = newPlayer.PosBottom - Abs(newPlayer.SourceBottom - newPlayer.SourceTop)
    newPlayer.PosRight = newPlayer.PosLeft + Abs(newPlayer.PosRight - newPlayer.PosLeft)
    newPlayer.direction = Player.direction
    Player.Inercia = 0
    If p_addtocollection Then
        col_oldPlayers.Add Player
    End If
    Col.Remove Player.Key
    Set Player = New clsObject
    MoveProp vObjProp, Player
    Player.NextLevelChar = newPlayer.NextLevelChar
    Player.PosLeft = newPlayer.PosLeft
    Player.PosBottom = newPlayer.PosBottom
    Player.PosRight = newPlayer.PosRight
    Player.PosTop = newPlayer.PosTop
    Player.direction = newPlayer.direction
    KeyIdx = KeyIdx + 1
    Player.Key = Player.CharType & KeyIdx
    Col.Add Player, Player.Key
    Set newPlayer = Nothing
End Sub

Public Sub InitObjects(vObj As String, vFilePath As String)

    Set vObjProp = New clsObjProp
    vObjProp.CharType = GetFromIni(vObj, "CharType", vFilePath)
    Load picGoomba(picGoomba.Count)
    picGoomba(picGoomba.Count - 2).Picture = LoadPicture(Replace$(GetFromIni(vObj, "Picture", vFilePath), "AppPath", App.Path))
    vObjProp.hdc = picGoomba(picGoomba.Count - 2).hdc
    vObjProp.CanHit = Val(GetFromIni(vObj, "CanHit", vFilePath))
    'vObjProp.FireBall = Val(GetFromIni(vObj, "FireBall", vFilePath))
    vObjProp.HittedByTop = Val(GetFromIni(vObj, "HittedByTop", vFilePath))
    vObjProp.FireBallCanHit = Val(GetFromIni(vObj, "FireBallCanHit", vFilePath))
    vObjProp.HittedByLeft = Val(GetFromIni(vObj, "HittedByLeft", vFilePath))
    vObjProp.HittedByRight = Val(GetFromIni(vObj, "HittedByRight", vFilePath))
    vObjProp.HittedByBottom = Val(GetFromIni(vObj, "HittedByBottom", vFilePath))
    vObjProp.SourceLeft = Val(GetFromIni(vObj, "SourceLeft", vFilePath))
    vObjProp.SourceRight = Val(Val(GetFromIni(vObj, "SourceRight", vFilePath)))
    vObjProp.SourceTop = Val(GetFromIni(vObj, "SourceTop", vFilePath))
    vObjProp.SourceBottom = Val(GetFromIni(vObj, "SourceBottom", vFilePath))
    vObjProp.Velocity = Val(GetFromIni(vObj, "Velocity", vFilePath))
    vObjProp.JumpVelocity = Val(GetFromIni(vObj, "JumpVelocity", vFilePath))
    vObjProp.JumpSize = Val(GetFromIni(vObj, "JumpSize", vFilePath))
    vObjProp.AnimFrom = Val(GetFromIni(vObj, "AnimFrom", vFilePath))
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
    vObjProp.StartPosOffsetX = Val(GetFromIni(vObj, "StartPosOffsetX", vFilePath))
    vObjProp.ChangeFrom = Val(GetFromIni(vObj, "ChangeFrom", vFilePath))
    vObjProp.JumpFrom = Val(GetFromIni(vObj, "JumpFrom", vFilePath))
    vObjProp.JumpTo = Val(GetFromIni(vObj, "JumpTo", vFilePath))
    vObjProp.DieFrame = Val(GetFromIni(vObj, "DieFrame", vFilePath))
    vObjProp.CanCrouch = Val(GetFromIni(vObj, "CanCrouch", vFilePath))
    vObjProp.CrouchFrame = Val(GetFromIni(vObj, "CrouchFrame", vFilePath))
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
    vObjProp.direction = Val(GetFromIni(vObj, "Direction", vFilePath))
    vObjProp.UserSelection = Val(GetFromIni(vObj, "UserSelection", vFilePath))
    vObjProp.GrowTo = GetFromIni(vObj, "GrowTo", vFilePath)
    vObjProp.MakeGrow = Val(GetFromIni(vObj, "MakeGrow", vFilePath))
    vObjProp.DieWhenHits = Val(GetFromIni(vObj, "DieWhenHits", vFilePath))
    vObjProp.FireBall = GetFromIni(vObj, "FireBall", vFilePath)
    vObjProp.FireSnd = Replace$(GetFromIni(vObj, "FireSnd", vFilePath), "AppPath", App.Path)
    vObjProp.JumpSnd = Replace$(GetFromIni(vObj, "JumpSnd", vFilePath), "AppPath", App.Path)
    vObjProp.DieSnd = Replace$(GetFromIni(vObj, "DieSnd", vFilePath), "AppPath", App.Path)
    vObjProp.Raising = Val(GetFromIni(vObj, "Raising", vFilePath))
    vObjProp.Descending = Val(GetFromIni(vObj, "Descending", vFilePath))
    vObjProp.Raisetime = Val(GetFromIni(vObj, "RaiseTime", vFilePath))
    vObjProp.JumpSnd = GetShortPath(vObjProp.JumpSnd)
    vObjProp.DieSnd = GetShortPath(vObjProp.DieSnd)
    vObjProp.FireSnd = GetShortPath(vObjProp.FireSnd)
    
    If vObjProp.Fixed Then
        vFixed = vFixed & vObjProp.CharType
    End If
    
    colObjProp.Add vObjProp
    
End Sub

Public Sub Unload_Objects()
Dim i As Long
    On Error Resume Next
    For i = Col.Count To 1
        Set Col.Item(i) = Nothing
        Col.Remove (i)
    Next
    
    Set Col = Nothing
    
    For i = cBreakeable.Count To 1
        Set cBreakeable.Item(i) = Nothing
        cBreakeable.Remove (i)
    Next
    
    Set cBreakeable = Nothing
    Set col_oldPlayers = Nothing
    Set oTimer = Nothing
    Set Player = Nothing
    
    vFixed = ""
End Sub

