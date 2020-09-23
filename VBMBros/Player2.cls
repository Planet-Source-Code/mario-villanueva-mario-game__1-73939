VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarhdc As Long 'copia local
Private mvarVelocity As Long  ' Velocidad para moverse
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
Private mvarDieFrame As Integer 'copia local
Private mvarRemoveWhenDies As Boolean 'copia local
Private mvarMakeJumpWhenHitted As Boolean 'copia local
Private mvarSolid As Boolean 'copia local
Private mvarFixed As Boolean 'copia local

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
End Property

Friend Property Get PosBottom() As Long
    PosBottom = mvarPosBottom
End Property

Friend Property Let PosTop(ByVal vData As Long)
    mvarPosBottom = vData + (mvarSourceBottom - mvarSourceTop)
    mvarPosTop = vData
End Property

Friend Property Get PosTop() As Long
    PosTop = mvarPosTop
End Property

Friend Property Let PosRight(ByVal vData As Long)
    mvarPosRight = vData
End Property

Friend Property Get PosRight() As Long
    PosRight = mvarPosRight
End Property

Friend Property Let PosLeft(ByVal vData As Long)
    mvarPosLeft = vData
    mvarPosRight = mvarPosLeft + 16 '(vData - mvarPosLeft)
End Property

Friend Property Get PosLeft() As Long
    PosLeft = mvarPosLeft
End Property

Friend Property Let Direction(ByVal vData As Boolean)
    mvarDirection = vData
End Property

Friend Property Get Direction() As Boolean
    Direction = mvarDirection
End Property

Friend Property Let Velocity(ByVal vData As Long)
    mvarVelocity = vData
End Property

Friend Property Get Velocity() As Long
    Velocity = mvarVelocity
End Property

Friend Property Let Jump(ByVal vData As Boolean)
    If vData = True Then
        mvarJump = vData
    Else
        mvarJump = vData
    End If
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

Public Function MoveRight() As Long
    mvarflip = False
    mvarDirection = True
End Function

Public Function MoveLeft() As Long
    mvarflip = True
    mvarDirection = False
End Function

Public Sub Draw(RunFrame As Long, OffsetX As Long)
    rctSrc.left = RunFrame * mvarSourceLeft
    rctSrc.right = rctSrc.left + 16
    rctSrc.top = mvarSourceTop
    rctSrc.bottom = mvarSourceBottom
    rctDst.left = mvarPosLeft - OffsetX
    rctDst.right = mvarPosRight - OffsetX
    rctDst.top = mvarPosTop
    rctDst.bottom = mvarPosBottom
    TransparentBlt Form1.hdc, rctDst, hdc, rctSrc, vbMagenta, mvarflip
End Sub
