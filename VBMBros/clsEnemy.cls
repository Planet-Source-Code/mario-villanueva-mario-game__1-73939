VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Implements clsPlayer
Implements clsAI

Private jumpcnt As Long
Private fallcnt As Long
Private RunFrame As Long
Private mMaxRunframe As Integer
Private mMinRunframe As Integer
Private mvarAnimFrom As Integer
Private mvarAnimTo   As Integer
Private mclsPlayer As New clsPlayer
Private mvarHittedByLeft As Boolean
Private mvarHittedByRight As Boolean
Private mvarHittedByTop As Boolean
Private mvarHittedByBottom As Boolean
Private mvarDie As Boolean
Private mvarUser As Boolean
Private mvarJumping As Boolean
Private mvarFalling As Boolean
Private mvarMoving As Boolean
Private mClsAI As New clsAI
Private Vel As Long
Private mvarCreateWhenHitted As String
Private mvarCreatePlace As String
Private mvarInercia As Boolean

Public Property Let Inercia(ByVal vData As Boolean)
    mvarInercia = vData
End Property

Public Property Get Inercia() As Boolean
    Inercia = mvarInercia
End Property

Private Sub Class_Initialize()
    mclsPlayer.Velocity = 2
    mclsPlayer.JumpSize = 32
End Sub

Public Property Let Fixed(ByVal vData As Boolean)
    mclsPlayer.Fixed = vData
End Property

Public Property Get Fixed() As Boolean
    Fixed = mclsPlayer.Fixed
End Property

Public Property Let Solid(ByVal vData As Boolean)
    mclsPlayer.Solid = vData
End Property

Public Property Get Solid() As Boolean
    Solid = mclsPlayer.Solid
End Property

Public Property Let MakeJumpWhenHitted(ByVal vData As Boolean)
    mclsPlayer.MakeJumpWhenHitted = vData
End Property

Public Property Get MakeJumpWhenHitted() As Boolean
    MakeJumpWhenHitted = mclsPlayer.MakeJumpWhenHitted
End Property

Public Property Let AnimTo(ByVal vData As Integer)
    mvarAnimTo = vData
End Property

Public Property Get RemoveWhenDies() As Boolean
    RemoveWhenDies = mclsPlayer.RemoveWhenDies
End Property

Public Property Let RemoveWhenDies(ByVal vData As Boolean)
    mclsPlayer.RemoveWhenDies = vData
End Property

Public Property Get DieFrame() As Integer
    DieFrame = mclsPlayer.DieFrame
End Property

Public Property Let DieFrame(ByVal vData As Integer)
    mclsPlayer.DieFrame = vData
End Property

Public Property Get AnimTo() As Integer
    AnimTo = mvarAnimTo
End Property

Public Property Let AnimFrom(ByVal vData As Integer)
    mvarAnimFrom = vData
End Property

Public Property Get AnimFrom() As Integer
    AnimFrom = mvarAnimFrom
End Property

Public Property Let CreatePlace(ByVal vData As String)
    mvarCreatePlace = vData
End Property

Public Property Get CreatePlace() As String
    CreatePlace = mvarCreatePlace
End Property

Public Property Let CreateWhenHitted(ByVal vData As String)
    mvarCreateWhenHitted = vData
End Property

Public Property Get CreateWhenHitted() As String
    CreateWhenHitted = mvarCreateWhenHitted
End Property

Friend Property Get Died() As Boolean
    Died = mvarDie
End Property

Friend Property Let Died(ByVal vData As Boolean)
    mvarDie = vData
End Property

Friend Property Let User(ByVal vData As Long)
    mvarUser = vData
End Property

Friend Property Get User() As Long
    User = mvarUser
End Property

Friend Property Get HittedByLeft() As Long
    HittedByLeft = mvarHittedByLeft
End Property

Friend Property Let HittedByLeft(ByVal vData As Long)
    mvarHittedByLeft = vData
End Property

Friend Property Get HittedByTop() As Long
    HittedByTop = mvarHittedByTop
End Property

Friend Property Let HittedByTop(ByVal vData As Long)
    mvarHittedByTop = vData
End Property

Friend Property Get HittedByRight() As Long
    HittedByRight = mvarHittedByRight
End Property

Friend Property Let HittedByRight(ByVal vData As Long)
    mvarHittedByRight = vData
End Property

Friend Property Get HittedByBottom() As Long
    HittedByBottom = mvarHittedByBottom
End Property

Friend Property Let HittedByBottom(ByVal vData As Long)
    mvarHittedByBottom = vData
End Property

Friend Property Get SourceBottom() As Long
    SourceBottom = mclsPlayer.SourceBottom()
End Property

Friend Property Let SourceBottom(ByVal vData As Long)
    mclsPlayer.SourceBottom = vData
End Property

Friend Property Get HorizontalMove() As Boolean
    HorizontalMove = mClsAI.HorizontalMove()
End Property

Friend Property Let HorizontalMove(ByVal vData As Boolean)
    mClsAI.HorizontalMove = vData
End Property

Friend Property Get CanFall() As Boolean
    CanFall = mclsPlayer.CanFall()
End Property

Friend Property Let CanFall(ByVal vData As Boolean)
    mclsPlayer.CanFall = vData
End Property

Friend Property Get SourceTop() As Long
    SourceTop = mclsPlayer.SourceTop()
End Property

Friend Property Let SourceTop(ByVal vData As Long)
    mclsPlayer.SourceTop = vData
End Property

Friend Property Get SourceRight() As Long
    SourceRight = mclsPlayer.SourceRight()
End Property

Friend Property Let SourceRight(ByVal vData As Long)
    mclsPlayer.SourceRight = vData
End Property

Friend Property Get SourceLeft() As Long
    SourceLeft = mclsPlayer.SourceLeft()
End Property

Friend Property Let SourceLeft(ByVal vData As Long)
    mclsPlayer.SourceLeft = vData
End Property

Public Sub StopMotion()
    RunFrame = mvarAnimFrom
    mvarMoving = False
    mvarInercia = True
End Sub

Public Sub PerformAI()
    mClsAI.PerformAI Me
End Sub

Friend Property Get InitedAI() As Boolean
    InitedAI = mClsAI.Inited
End Property

Friend Property Let InitedAI(ByVal vData As Boolean)
    mClsAI.Inited = vData
End Property

Friend Property Get MinRunframe() As Integer
    MinRunframe = mMinRunframe
End Property

Friend Property Let MinRunframe(ByVal vData As Integer)
    mMinRunframe = vData
End Property

Friend Property Get MaxRunframe() As Integer
    MaxRunframe = mMaxRunframe
End Property

Friend Property Let MaxRunframe(ByVal vData As Integer)
    mMaxRunframe = vData
End Property

Friend Property Get Visible() As Boolean
    Visible = mclsPlayer.Visible()
End Property

Friend Property Let Visible(ByVal vData As Boolean)
    mclsPlayer.Visible = vData
End Property

Friend Property Get RScreenLimit() As Long
    RScreenLimit = mclsPlayer.RScreenLimit()
End Property

Friend Property Let RScreenLimit(ByVal vData As Long)
    mclsPlayer.RScreenLimit = vData
End Property

Friend Property Let PosBottom(ByVal vData As Long)
    mclsPlayer.PosBottom = vData
End Property

Friend Property Get PosBottom() As Long
    PosBottom = mclsPlayer.PosBottom()
End Property

Friend Property Let PosTop(ByVal vData As Long)
    mclsPlayer.PosTop = vData
End Property

Friend Property Get PosTop() As Long
    PosTop = mclsPlayer.PosTop()
End Property

Friend Property Let PosRight(ByVal vData As Long)
    mclsPlayer.PosRight = vData
End Property

Friend Property Get PosRight() As Long
    PosRight = mclsPlayer.PosRight()
End Property

Friend Property Let PosLeft(ByVal vData As Long)
    mclsPlayer.PosLeft = vData
End Property

Friend Property Get PosLeft() As Long
    PosLeft = mclsPlayer.PosLeft
End Property

Friend Property Let Direction(ByVal vData As String)
    mclsPlayer.Direction = vData
End Property

Friend Property Get Direction() As String
    Direction = mclsPlayer.Direction
End Property

Friend Property Let Velocity(ByVal vData As Long)
    mclsPlayer.Velocity = vData
End Property

Friend Property Get Velocity() As Long
    Velocity = mclsPlayer.Velocity
End Property

Friend Property Let Jump(ByVal vData As Boolean)
Dim coordX As Long, coordY As Long
    coordY = Fix(mclsPlayer.PosBottom / 16) + 1 'Chequeo el proximo bloque si esta ocupado por una X
    coordX = Fix((mclsPlayer.CharPos + mclsPlayer.PosLeft + 8) / 16) + 1
    If InStrB(1, vFixed, Form1.GetBlock(coordX, coordY)) <> 0 Then
        mclsPlayer.Jump = vData
    End If
    If vData = False Then
        mclsPlayer.Jump = vData
    End If
End Property

Friend Property Get Jump() As Boolean
    Jump = mclsPlayer.Jump
End Property

Friend Property Let hdc(ByVal vData As Long)
    mclsPlayer.hdc = vData
End Property

Friend Property Get hdc() As Long
    hdc = mclsPlayer.hdc
End Property

Public Sub Draw(OffsetX As Long)
    mclsPlayer.Draw RunFrame, OffsetX
End Sub

Private Sub clsPlayer_Draw(RunFrame As Long, OffsetX As Long)
    mclsPlayer.Draw RunFrame, OffsetX
End Sub

Private Function clsPlayer_MoveLeft() As Long
    'mclsPlayer.MoveLeft
End Function

Private Function clsPlayer_MoveRight() As Long
    'mclsPlayer.MoveRight
End Function

Private Property Get clsPlayer_RScreenLimit() As Long
    PL_RScreenLimit = mclsPlayer.RScreenLimit()
End Property

Private Property Let clsPlayer_RScreenLimit(ByVal vData As Long)
    mclsPlayer.RScreenLimit = vData
End Property

Private Property Let clsPlayer_Fixed(ByVal RHS As Boolean)

End Property

Private Property Get clsPlayer_Fixed() As Boolean

End Property

Private Property Let clsPlayer_MakeJumpWhenHitted(ByVal RHS As Boolean)

End Property

Private Property Get clsPlayer_MakeJumpWhenHitted() As Boolean

End Property

Private Property Get clsAI_HorizontalMove() As Boolean

End Property

Private Property Let clsAI_HorizontalMove(ByVal RHS As Boolean)

End Property

Private Property Get clsAI_Inited() As Boolean

End Property

Private Property Let clsAI_Inited(ByVal RHS As Boolean)

End Property

Private Sub clsAI_PerformAI(mclsEnemy As clsEnemy)

End Sub

Private Property Let clsPlayer_CanFall(ByVal RHS As Boolean)

End Property

Private Property Get clsPlayer_CanFall() As Boolean

End Property

Private Property Let clsPlayer_DieFrame(ByVal RHS As Integer)

End Property

Private Property Get clsPlayer_DieFrame() As Integer

End Property

Private Property Let clsPlayer_RemoveWhenDies(ByVal RHS As Boolean)

End Property

Private Property Get clsPlayer_RemoveWhenDies() As Boolean

End Property

Private Property Let clsPlayer_Solid(ByVal RHS As Boolean)

End Property

Private Property Get clsPlayer_Solid() As Boolean

End Property

Private Property Let clsPlayer_SourceBottom(ByVal RHS As Long)

End Property

Private Property Get clsPlayer_SourceBottom() As Long

End Property

Private Property Let clsPlayer_SourceLeft(ByVal RHS As Long)

End Property

Private Property Get clsPlayer_SourceLeft() As Long

End Property

Private Property Let clsPlayer_SourceRight(ByVal RHS As Long)

End Property

Private Property Get clsPlayer_SourceRight() As Long

End Property

Private Property Let clsPlayer_SourceTop(ByVal RHS As Long)

End Property

Private Property Get clsPlayer_SourceTop() As Long

End Property

Private Property Let clsPlayer_Visible(ByVal RHS As Boolean)

End Property

Private Property Get clsPlayer_Visible() As Boolean

End Property

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'                       Implementaciones propias del personaje
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Friend Sub Animate(Col As Collection, time As Long)
Dim coordX As Long, coordY As Long, coordXL As Long, coordXR As Long
Dim bot As clsEnemy
Static dieCounter As Long
Static mtime
mtime = mtime + time

'Matar al personaje
If mvarDie = True Then
    If mvarUser = True Then
        'Muere el Jugador
        RunFrame = mclsPlayer.DieFrame
        coordY = Fix(mclsPlayer.PosBottom / 16) + 1
        If coordY > 3 And dieCounter = 0 Then
            mclsPlayer.PosTop = mclsPlayer.PosTop - 4
        ElseIf coordY > 16 Then
            mclsPlayer.Visible = False
        ElseIf coordY = 3 Or dieCounter = 1 Then
            dieCounter = 1
            mclsPlayer.PosTop = mclsPlayer.PosTop + 4
        End If
    Else
        'Muere un Villano
        RunFrame = mclsPlayer.DieFrame
        If dieCounter > 20 And mclsPlayer.RemoveWhenDies Then
            mclsPlayer.Visible = False
        End If
        dieCounter = dieCounter + 1
    End If
    Exit Sub
End If

' Hacer animacion de caida
If mclsPlayer.Jump = False And mvarJumping = False Then
    If CanFall Then
        coordY = Fix(mclsPlayer.PosBottom / 16) + 1 'Chequeo el proximo bloque si esta ocupado por una X
        coordXL = Fix((mclsPlayer.CharPos + mclsPlayer.PosLeft + 2) / 16) + 1
        coordXR = Fix((mclsPlayer.CharPos + mclsPlayer.PosRight - 2) / 16) + 1
        If InStrB(1, vFixed, Form1.GetBlock(coordXL, coordY)) = 0 _
           And InStrB(1, vFixed, Form1.GetBlock(coordXR, coordY)) = 0 Then
'          Si el bloque de abajo no esta ocupado, caer
            If (fallcnt < 4 And jumpcnt > 12) Then
                fallcnt = fallcnt + 1
                mclsPlayer.PosTop = mclsPlayer.PosTop + 1 'mclsPlayer.Velocity
            Else
                mclsPlayer.PosTop = mclsPlayer.PosTop + 3
            End If
            RunFrame = mMaxRunframe + 2
            mvarFalling = True
        Else
'          Si el bloque de abajo esta ocupado, parase sobre el
            If mvarUser Then
                Debug.Print ""
            End If
            mvarJumping = False
            mvarFalling = False
            mclsPlayer.PosTop = (coordY - 1) * 16 - (mclsPlayer.SourceBottom - mclsPlayer.SourceTop)
            fallcnt = 0
        End If
    End If
    If mvarJumping = False And mvarFalling = False And mvarMoving = False Then
        If mtime > 500 Then
            If RunFrame + 1 > mvarAnimTo Then
                RunFrame = mvarAnimFrom
            Else
                RunFrame = RunFrame + 1
            End If
            mtime = 0
        End If
    End If
    jumpcnt = 0
' Hacer animacion de salto
Else
    fallcnt = 0
    coordY = Fix(mclsPlayer.PosBottom / 16)  'Chequeo el proximo bloque si esta ocupado por una X
    coordXL = Fix((mclsPlayer.CharPos + mclsPlayer.PosLeft + 2) / 16) + 1
    coordXR = Fix((mclsPlayer.CharPos + mclsPlayer.PosRight - 2) / 16) + 1
        If InStrB(1, vFixed, Form1.GetBlock(coordXL, coordY)) = 0 _
       And InStrB(1, vFixed, Form1.GetBlock(coordXR, coordY)) = 0 Then
        If jumpcnt < mclsPlayer.JumpSize Then
            If (mclsPlayer.Jump = True And jumpcnt < mclsPlayer.JumpSize - 4) Or _
                (mclsPlayer.Jump = False And mvarJumping = True And jumpcnt < 12) Then
                mvarJumping = True 'Para que siga entrando por este IF
                jumpcnt = jumpcnt + 1
                mclsPlayer.PosTop = mclsPlayer.PosTop - 3 'mclsPlayer.Velocity
                If jumpcnt + 8 > mclsPlayer.JumpSize Then
                    mclsPlayer.Jump = False
                End If
            Else
                If (mvarJumping = True And jumpcnt < 4) Or _
                    (mclsPlayer.Jump = True And jumpcnt < mclsPlayer.JumpSize) Then
                    jumpcnt = jumpcnt + 1
                    mclsPlayer.PosTop = mclsPlayer.PosTop - 1
                Else
                    mvarJumping = False
                    mvarFalling = True
                End If
            End If
            If Vel <> 0 And Vel < 32 Then Vel = Vel + Sgn(Vel)
        Else
            mvarJumping = False 'Para que no entre mas por este IF
            mvarFalling = True
            mclsPlayer.Jump = False
            jumpcnt = 0
        End If
    Else
        mvarJumping = False
        mvarFalling = True
        mclsPlayer.Jump = False
        jumpcnt = 0
    End If
    RunFrame = mMaxRunframe + 2
End If

' Chequear si hubo colisiones con otros personajes
For Each bot In Col
    If bot.Died = False Then
        ' Este personaje golpeo a otro de Izquierda a Derecha
        If PosLeft < bot.PosRight And PosRight > bot.PosRight And _
           PosBottom = bot.PosBottom Then
            If bot.HittedByRight Then
                bot.WasHitted
                If bot.MakeJumpWhenHitted Then
                    mclsPlayer.Jump = True
                End If
            End If
        End If
        ' Este personaje golpeo a otro de Derecha a Izquierda
        If PosRight > bot.PosLeft And PosLeft < bot.PosLeft And _
           PosBottom = bot.PosBottom Then
            If bot.HittedByLeft Then
                bot.WasHitted
                If bot.MakeJumpWhenHitted Then
                    mclsPlayer.Jump = True
                End If
            End If
        End If
        ' Este personaje golpeo a otro Pisandolo
        If PosBottom > bot.PosTop And PosBottom < bot.PosBottom And _
         ((PosLeft >= bot.PosLeft And PosLeft < bot.PosRight) Or _
          (PosLeft < bot.PosLeft And PosRight > bot.PosLeft)) Then
            If bot.HittedByTop Then
                bot.WasHitted
                If bot.MakeJumpWhenHitted Then
                    mclsPlayer.Jump = True
                End If
            End If
        End If
        ' Este personaje golpeo a otro de abajo hacia arriba
        If PosBottom > bot.PosBottom And PosTop < bot.PosBottom And _
         ((PosLeft >= bot.PosLeft And PosLeft < bot.PosRight) Or _
          (PosLeft < bot.PosLeft And PosRight > bot.PosLeft)) Then
            If bot.HittedByBottom Then
                bot.WasHitted
                If bot.MakeJumpWhenHitted Then
                    mclsPlayer.Jump = True
                End If
            End If
        End If
    End If
Next
End Sub

Public Sub WasHitted()
Dim Enemy As clsEnemy
    'Matar al personaje (mostrar animacion)
    mvarDie = True
    'Si el personaje crea a otro, por ej una caja ? crear el objeto nuevo y posicionarlo
    If CreateWhenHitted <> "" Then
        Set Enemy = New clsEnemy
        Set vObjProp = cObj(CreateWhenHitted)
        MoveProp vObjProp, Enemy
        Enemy.PosLeft = PosLeft - 16
        Enemy.PosRight = PosRight - 16
        If CreatePlace = "OnTop" Then
            Enemy.PosTop = PosTop - 16
            Enemy.PosBottom = PosBottom - 16
        Else
            Enemy.PosBottom = PosBottom
            Enemy.PosTop = PosTop
        End If
        Col.Add Enemy
        CreateWhenHitted = ""
    End If
End Sub

Public Function MoveRight(time As Long) As Long
Dim coordX As Long, coordY As Long, inc As Long
Static mtime As Long
    mtime = mtime + time
    If (mvarJumping Or mvarFalling) And mvarMoving And mclsPlayer.Direction <> True Then Exit Function
    mvarInercia = False
    coordY = Fix((mclsPlayer.PosBottom + 10) / 16) 'Chequeo el proximo bloque si esta ocupado por una X
    coordX = Fix((mclsPlayer.CharPos + mclsPlayer.PosRight + 2) / 16) + 1
    mvarMoving = True
    
    If InStrB(1, vFixed, Form1.GetBlock(coordX, coordY)) = 0 Then
        If Vel < 4 And mvarUser Then
            inc = 1
        Else
            inc = mclsPlayer.Velocity
        End If
        mclsPlayer.PosLeft = mclsPlayer.PosLeft + inc
        If Vel < 16 And mvarUser Then
            Vel = Vel + inc
        End If
        mclsPlayer.PosRight = mclsPlayer.PosLeft + 16
    End If
    
    'If mvarJumping = False And mvarFalling = False And mvarUser Then
        mclsPlayer.MoveRight
    'End If
    
    If mtime > 300 Then
        RunFrame = RunFrame + 1
        mtime = 0
    End If
    If RunFrame > mMaxRunframe Then
        RunFrame = mMinRunframe
    End If
End Function

Public Function MoveLeft(time As Long) As Long
Dim coordX As Long, coordY As Long, inc As Long
Static mtime
mtime = mtime + time
    If (mvarJumping Or mvarFalling) And mvarMoving And mclsPlayer.Direction <> False Then Exit Function
    mvarInercia = False
    coordY = Fix((mclsPlayer.PosBottom + 10) / 16) 'Chequeo el proximo bloque si esta ocupado por una X
    coordX = Fix((mclsPlayer.CharPos + mclsPlayer.PosLeft - 2) / 16) + 1
    mvarMoving = True
    
    If InStrB(1, vFixed, Form1.GetBlock(coordX, coordY)) = 0 Then
        If mclsPlayer.PosLeft > 0 Then
            If Vel > -4 And mvarUser Then
                inc = -1
            Else
                inc = -mclsPlayer.Velocity
            End If
            mclsPlayer.PosLeft = mclsPlayer.PosLeft + inc
            If Vel > -16 And mvarUser Then
                Vel = Vel + inc
            End If
            mclsPlayer.PosRight = mclsPlayer.PosLeft + 16
        Else
            mclsPlayer.PosLeft = 0
        End If
    End If
    
    'If mvarJumping = False And mvarFalling = False And mvarUser Then
        MoveLeft = mclsPlayer.MoveLeft()
    'End If
    
    If mtime > 300 Then
        RunFrame = RunFrame + 1
        mtime = 0
    End If
    If RunFrame > mMaxRunframe Then
        RunFrame = mMinRunframe
    End If
End Function

Public Sub ManejarInercia(time As Long)
Dim coordX As Long, coordY As Long
Dim vSign As Long
Static mtime
mtime = mtime + time

    If Vel = 0 Or mvarUser = False Then Exit Sub
    
    vSign = Sgn(Vel)
    
    coordY = Fix((mclsPlayer.PosBottom + 10) / 16) 'Chequeo el proximo bloque si esta ocupado por una X
    coordX = Fix((mclsPlayer.CharPos + mclsPlayer.PosLeft + (16 * vSign)) _
                  / 16) + 1
    mvarMoving = True
    
    If InStrB(1, vFixed, Form1.GetBlock(coordX, coordY)) = 0 Then
        If mclsPlayer.PosLeft > 0 Then
            If mvarJumping = False And mvarFalling = False Then
                mclsPlayer.PosLeft = mclsPlayer.PosLeft + vSign
            Else
                mclsPlayer.PosLeft = mclsPlayer.PosLeft + vSign * 2
            End If
            mclsPlayer.PosRight = mclsPlayer.PosLeft + 16
        Else
            mclsPlayer.PosLeft = 0
        End If
    End If
    
    Vel = Vel - vSign
    
    If mvarJumping = False And mvarFalling = False Then
        If vSign > 0 Then
            mclsPlayer.MoveRight
        Else
            mclsPlayer.MoveLeft
        End If
    End If
    
    If mtime > 500 Then
        RunFrame = RunFrame + 1
        mtime = 0
    End If
    If RunFrame > mMaxRunframe Then
        RunFrame = mMinRunframe
    End If
    If Vel = 0 Then
        mvarMoving = False
    End If
End Sub
