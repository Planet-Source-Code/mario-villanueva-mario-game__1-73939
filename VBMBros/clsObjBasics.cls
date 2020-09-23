VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjBasics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarhdc As Long 'copia local
Private mvarVelocity As Long  ' Velocidad para moverse
Private mvarJumpVelocity As Long  ' Velocidad para moverse
Private mvarDirection As Boolean 'copia local
Private mvarPosLeft As Long 'copia local
Private mvarPosRight As Long 'copia local
Private mvarPosTop As Long 'copia local
Private mvarPosBottom As Long 'copia local
Private mvarJump As Boolean   ' Esta Saltando?
Private mvarCharPos As Long   ' Posicion dentro del Mapa
Private mvarJumpSize As Long  ' Limite de Salto
Private mvarflip As Boolean
Private mvarVisible As Boolean
Private mvarRScreenLimit As Long 'copia local
Private mvarSourceLeft As Long 'copia local
Private mvarSourceRight As Long 'copia local
Private mvarSourceTop As Long 'copia local
Private mvarSourceBottom As Long 'copia local
Private mvarCanFall As Boolean 'copia local
Private mvarStartPosOffsetX As Long
Private mvarDieFrame As Integer 'copia local
Private mvarRemoveWhenDies As Boolean 'copia local
Private mvarMakeJumpWhenHitted As Boolean 'copia local
Private mvarJumpWhenHitted As Boolean 'copia local
Private mvarSolid As Boolean 'copia local
Private mvarFixed As Boolean 'copia local
Private mvarCharType As String 'local copy
Private mvarStep As Long
Private mvarBreaking As Boolean 'local copy
Private mvarRaising As Boolean 'local copy
Private mvarHidden As Boolean 'local copy
Private mvarDescending As Boolean 'local copy
Private mvarBreaked As Boolean 'local copy
Private mvarcHeight As Long
Private mvarcWidth As Long
Private raisingstep As Long
Private mvarRaisetime As Long
'variables locales para almacenar los valores de las propiedades
Private mvarCanCrouch As Boolean 'copia local
Private mvarCrouchFrame As Integer 'copia local
Private mvarCrouched As Boolean 'copia local
Public Property Let Crouched(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Crouched = 5
    mvarCrouched = vData
End Property


Public Property Get Crouched() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Crouched
    Crouched = mvarCrouched
End Property



Public Property Let CrouchFrame(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CrouchFrame = 5
    mvarCrouchFrame = vData
End Property


Public Property Get CrouchFrame() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CrouchFrame
    CrouchFrame = mvarCrouchFrame
End Property



Public Property Let CanCrouch(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CanCrouch = 5
    mvarCanCrouch = vData
End Property


Public Property Get CanCrouch() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CanCrouch
    CanCrouch = mvarCanCrouch
End Property




Public Property Let Raisetime(ByVal vData As Long)
    mvarRaisetime = vData
End Property

Public Property Get Raisetime() As Long
    Raisetime = mvarRaisetime
End Property

Public Property Let Raising(ByVal vData As Boolean)
    mvarRaising = vData
End Property

Public Property Get Raising() As Boolean
    Raising = mvarRaising
End Property

Public Property Let Descending(ByVal vData As Boolean)
    mvarDescending = vData
End Property

Public Property Get Descending() As Boolean
    Descending = mvarDescending
End Property


Public Property Let cHeight(ByVal vData As Long)
    mvarcHeight = vData
End Property

Public Property Get cHeight() As Long
    cHeight = mvarcHeight
End Property

Public Property Let cWidth(ByVal vData As Long)
    mvarcWidth = vData
End Property

Public Property Get cWidth() As Long
    cWidth = mvarcWidth
End Property

Public Property Let Breaked(ByVal vData As Boolean)
    mvarBreaked = vData
End Property

Public Property Get Breaked() As Boolean
    Breaked = mvarBreaked
End Property

Public Property Let Breaking(ByVal vData As Boolean)
    mvarBreaking = vData
End Property

Public Property Get Breaking() As Boolean
    Breaking = mvarBreaking
End Property

Public Property Let CharType(ByVal vData As String)
    mvarCharType = vData
End Property

Public Property Get CharType() As String
    CharType = mvarCharType
End Property

Public Property Let Fixed(ByVal vData As Boolean)
    mvarFixed = vData
End Property

Public Property Get Fixed() As Boolean
    Fixed = mvarFixed
End Property

Public Property Let Solid(ByVal vData As Boolean)
    mvarSolid = vData
End Property

Public Property Get Solid() As Boolean
    Solid = mvarSolid
End Property

Public Property Let JumpWhenHitted(ByVal vData As Boolean)
    mvarJumpWhenHitted = vData
End Property

Public Property Get JumpWhenHitted() As Boolean
    JumpWhenHitted = mvarJumpWhenHitted
End Property

Public Property Let MakeJumpWhenHitted(ByVal vData As Boolean)
    mvarMakeJumpWhenHitted = vData
End Property

Public Property Get MakeJumpWhenHitted() As Boolean
    MakeJumpWhenHitted = mvarMakeJumpWhenHitted
End Property

Public Property Let RemoveWhenDies(ByVal vData As Boolean)
    mvarRemoveWhenDies = vData
End Property

Public Property Get RemoveWhenDies() As Boolean
    RemoveWhenDies = mvarRemoveWhenDies
End Property

Public Property Let DieFrame(ByVal vData As Integer)
    mvarDieFrame = vData
End Property

Public Property Get DieFrame() As Integer
    DieFrame = mvarDieFrame
End Property

Public Property Let StartPosOffsetX(ByVal vData As Long)
    mvarStartPosOffsetX = vData
End Property

Public Property Get StartPosOffsetX() As Long
    StartPosOffsetX = mvarStartPosOffsetX
End Property

Public Property Let CanFall(ByVal vData As Boolean)
    mvarCanFall = vData
End Property

Public Property Get CanFall() As Boolean
    CanFall = mvarCanFall
End Property

Public Property Let SourceBottom(ByVal vData As Long)
    mvarSourceBottom = vData
End Property

Public Property Get SourceBottom() As Long
    SourceBottom = mvarSourceBottom
End Property

Public Property Let SourceTop(ByVal vData As Long)
    mvarSourceTop = vData
End Property

Public Property Get SourceTop() As Long
    SourceTop = mvarSourceTop
End Property

Public Property Let SourceRight(ByVal vData As Long)
    mvarSourceRight = vData
End Property

Public Property Get SourceRight() As Long
    SourceRight = mvarSourceRight
End Property

Public Property Let SourceLeft(ByVal vData As Long)
    mvarSourceLeft = vData
End Property

Public Property Get SourceLeft() As Long
    SourceLeft = mvarSourceLeft
End Property

Public Property Let Hidden(ByVal vData As Boolean)
    mvarHidden = vData
End Property

Public Property Get Hidden() As Boolean
    Hidden = mvarHidden
End Property

Public Property Let Visible(ByVal vData As Boolean)
    mvarVisible = vData
End Property

Public Property Get Visible() As Boolean
    Visible = mvarVisible
End Property

Public Property Let RScreenLimit(ByVal vData As Long)
    mvarRScreenLimit = vData
End Property

Public Property Get RScreenLimit() As Long
    RScreenLimit = mvarRScreenLimit
End Property

Friend Property Let PosBottom(ByVal vData As Long)
    mvarPosTop = vData - (mvarSourceBottom - mvarSourceTop)
    mvarPosBottom = vData
    mvarcHeight = Abs(mvarPosBottom - mvarPosTop)
End Property

Friend Property Get PosBottom() As Long
    PosBottom = mvarPosBottom
End Property

Friend Property Let PosTop(ByVal vData As Long)
    mvarPosBottom = vData + (mvarSourceBottom - mvarSourceTop)
    mvarPosTop = vData
    mvarcHeight = Abs(mvarPosBottom - mvarPosTop)
End Property

Friend Property Get PosTop() As Long
    PosTop = mvarPosTop
End Property

Friend Property Let PosRight(ByVal vData As Long)
    mvarPosRight = vData
    mvarcWidth = Abs(mvarPosRight - mvarPosLeft)
End Property

Friend Property Get PosRight() As Long
    PosRight = mvarPosRight
End Property

Friend Property Let PosLeft(ByVal vData As Long)
    mvarPosLeft = vData
    mvarPosRight = mvarPosLeft + (mvarSourceRight - mvarSourceLeft)
    mvarcWidth = Abs(mvarPosRight - mvarPosLeft)
End Property

Friend Property Get PosLeft() As Long
    PosLeft = mvarPosLeft
End Property

Friend Property Let direction(ByVal vData As Boolean)
    mvarDirection = vData
    mvarflip = Not vData
End Property

Friend Property Get direction() As Boolean
    direction = mvarDirection
End Property


Public Property Let JumpVelocity(ByVal vData As Long)
    mvarJumpVelocity = vData
End Property

Public Property Get JumpVelocity() As Long
     JumpVelocity = mvarJumpVelocity
End Property

Friend Property Let Velocity(ByVal vData As Long)
    mvarVelocity = vData
End Property

Friend Property Get Velocity() As Long
    Velocity = mvarVelocity
End Property

Friend Property Let Jump(ByVal vData As Boolean)
    mvarJump = vData
End Property

Friend Property Get Jump() As Boolean
    Jump = mvarJump
End Property

Friend Property Let hdc(ByVal vData As Long)
    mvarhdc = vData
End Property

Friend Property Get hdc() As Long
    hdc = mvarhdc
End Property

Friend Property Let CharPos(ByVal vData As Long)
    mvarCharPos = vData
End Property

Friend Property Get CharPos() As Long
    CharPos = mvarCharPos
End Property

Friend Property Let JumpSize(ByVal vData As Long)
    mvarJumpSize = vData
End Property

Friend Property Get JumpSize() As Long
    JumpSize = mvarJumpSize
End Property

Public Function MoveRight(Optional direction As Long) As Long
    mvarflip = False
    If direction >= 0 Then
        mvarDirection = True
    Else
        mvarDirection = False
    End If
End Function

Public Function MoveLeft(Optional direction As Long) As Long
    mvarflip = True
    If direction <= 0 Then
        mvarDirection = False
    Else
        mvarDirection = True
    End If
End Function

Public Sub Draw(RunFrame As Long, OffsetX As Long)
Dim sWidth As Long, sHeight As Long, mult As Long
sWidth = mvarSourceRight - mvarSourceLeft
sHeight = mvarSourceBottom - mvarSourceTop

    If mvarBreaking Then
        If mvarStep < 50 Then
            mvarStep = mvarStep + 1
        Else
            mvarBreaked = True
            mvarBreaking = False
        End If
        rctSrc.Left = RunFrame * sWidth + sWidth / 2
        rctSrc.Right = RunFrame * sWidth + sWidth
        rctSrc.Top = mvarSourceTop
        rctSrc.Bottom = mvarSourceTop + sHeight / 2
        rctDst.Top = mvarPosTop + ((mvarStep / 2) - 9) ^ 2 - 81
        rctDst.Bottom = rctDst.Top + mvarcHeight / 2
        rctDst.Left = mvarPosLeft - OffsetX + mvarStep * 3
        rctDst.Right = rctDst.Left + mvarcWidth / 2
        If mvarFixed Or mvarSolid Then
            StretchPaint Form1.pForm1.hdc, rctDst, hdc, rctSrc
        Else
            TransparentBlt Form1.pForm1.hdc, rctDst, hdc, rctSrc, vbMagenta, mvarflip
        End If
        
        rctSrc.Left = RunFrame * sWidth + sWidth / 2
        rctSrc.Right = RunFrame * sWidth + sWidth
        rctSrc.Top = mvarSourceTop + sHeight / 2
        rctSrc.Bottom = mvarSourceTop + sHeight
        rctDst.Top = mvarPosTop + ((mvarStep / 2) - 5) ^ 2 - 25
        rctDst.Bottom = rctDst.Top + mvarcHeight / 2
        rctDst.Left = mvarPosLeft - OffsetX + mvarStep * 3
        rctDst.Right = rctDst.Left + mvarcWidth / 2
        If mvarFixed Or mvarSolid Then
            StretchPaint Form1.pForm1.hdc, rctDst, hdc, rctSrc
        Else
            TransparentBlt Form1.pForm1.hdc, rctDst, hdc, rctSrc, vbMagenta, mvarflip
        End If
        '
        rctSrc.Top = mvarSourceTop
        rctSrc.Bottom = mvarSourceTop + sHeight / 2
        rctSrc.Left = RunFrame * sWidth
        rctSrc.Right = RunFrame * sWidth + sWidth / 2
        rctDst.Top = mvarPosTop + ((mvarStep / 2) - 9) ^ 2 - 81
        rctDst.Bottom = rctDst.Top + mvarcHeight / 2
        rctDst.Left = mvarPosLeft - OffsetX - mvarStep * 3
        rctDst.Right = rctDst.Left + mvarcWidth / 2
        If mvarFixed Or mvarSolid Then
            StretchPaint Form1.pForm1.hdc, rctDst, hdc, rctSrc
        Else
            TransparentBlt Form1.pForm1.hdc, rctDst, hdc, rctSrc, vbMagenta, mvarflip
        End If
        
        rctSrc.Top = mvarSourceTop + sHeight / 2
        rctSrc.Bottom = mvarSourceBottom
        rctSrc.Left = RunFrame * sWidth
        rctSrc.Right = RunFrame * sWidth + sWidth / 2
        rctDst.Top = mvarPosTop + ((mvarStep / 2) - 5) ^ 2 - 25
        rctDst.Bottom = rctDst.Top + mvarcHeight / 2
        rctDst.Left = mvarPosLeft - OffsetX - mvarStep * 3
        rctDst.Right = rctDst.Left + mvarcWidth / 2
        If mvarFixed Or mvarSolid Then
            StretchPaint Form1.pForm1.hdc, rctDst, hdc, rctSrc
        Else
            TransparentBlt Form1.pForm1.hdc, rctDst, hdc, rctSrc, vbMagenta, mvarflip
        End If
    Else
        rctSrc.Left = RunFrame * sWidth
        rctSrc.Right = rctSrc.Left + sWidth
        rctSrc.Top = mvarSourceTop
        rctSrc.Bottom = mvarSourceBottom
        rctDst.Left = mvarPosLeft - OffsetX
        rctDst.Right = mvarPosRight - OffsetX
        rctDst.Top = mvarPosTop
        rctDst.Bottom = mvarPosBottom
        If Me.Hidden And Not mvarRaising Then Exit Sub
        If mvarRaising Then
            If Me.Hidden Then Me.Hidden = False
            mult = (mvarRaisetime / ((PosBottom - PosTop) / 64))
            If raisingstep >= (PosBottom - PosTop) * mult Then
                raisingstep = 0
                mvarRaising = False
            Else
                raisingstep = raisingstep + 1
            End If
            rctSrc.Top = mvarSourceTop
            rctSrc.Bottom = raisingstep \ mult
            rctDst.Top = mvarPosBottom - raisingstep \ mult
        ElseIf mvarDescending Then
            mult = (mvarRaisetime / ((PosBottom - PosTop) / 64))
            If raisingstep >= (PosBottom - PosTop) * mult Then
                raisingstep = 0
                mvarDescending = False
                Me.Hidden = True
                Exit Sub
            Else
                raisingstep = raisingstep + 1
            End If
            rctSrc.Top = mvarSourceTop
            rctSrc.Bottom = mvarSourceBottom - raisingstep \ mult
            rctDst.Top = mvarPosTop + raisingstep \ mult
        End If
        If Crouched Then
            rctSrc.Left = CrouchFrame * sWidth
            rctSrc.Right = rctSrc.Left + sWidth
        End If
        If Not mvarBreaked Then
            If mvarFixed Or mvarSolid Then
                StretchPaint Form1.pForm1.hdc, rctDst, hdc, rctSrc
            Else
                TransparentBlt Form1.pForm1.hdc, rctDst, hdc, rctSrc, vbMagenta, mvarflip
            End If
        End If
    End If
    
End Sub
