VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsbBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarhdc As Long 'copia local
Private mvarHittedByTop As Boolean 'copia local
Private mvarHittedByBottom As Boolean 'copia local
Private mvarHittedByLeft As Boolean 'copia local
Private mvarHittedByRight As Boolean 'copia local
Private mvarSourceLeft As Long 'copia local
Private mvarSourceTop As Long 'copia local
Private mvarSourceBottom As Long 'copia local
Private mvarSourceRight As Long 'copia local
'Private mvarMinRunframe As Long 'copia local
'Private mvarMaxRunframe As Long 'copia local
Private mvarVelocity As Long 'copia local
Private mvarJumpVelocity As Long 'copia local
Private mvarInitedAIBoolean As Boolean 'copia local
Private mvarPosLeft As Long 'copia local
Private mvarPosRight As Long 'copia local
Private mvarPosTop As Long 'copia local
Private mvarPosBottom As Long 'copia local
'Private mvarCreateWhenHitted As String 'copia local
'Private mvarCreatePlace As String 'copia local
Private mvarVisible As Boolean 'copia local
'Private mvarUser As Boolean 'copia local
'Private mvarCanFall As Boolean 'copia local
'Private mvarAnimFrom As Integer 'copia local
'Private mvarAnimTo As Integer 'copia local
'Private mvarDieFrame As Integer 'copia local
'Private mvarRemoveWhenDies As Boolean 'copia local
'Private mvarMakeJumpWhenHitted As Boolean 'copia local
'Private mvarSolid As Boolean 'copia local
'Private mvarFixed As Boolean 'copia local
'Private mvarInitAIWhenHitted As Boolean 'local copy
Private mvarCharType As String 'local copy
'Private mvarDieTiming As Long 'local copy
'Private mvarAI As String 'local copy
'Private mvarJumpFrom As Integer 'local copy
'Private mvarJumpTo As Integer 'local copy
'Private mvarJumpSize As Long
'Private mvarCanHit As Boolean 'local copy
'Private mvarChangePlayerTo As String 'local copy
'Private mvarNextLevel As Boolean 'local copy
'Private mvarHibernating As Boolean 'local copy
'Private mvarCanHitEnemies As Boolean 'local copy
'Private mvarCanBeBreaked As Boolean 'local copy
'Private mvarCanBreak As Boolean 'local copy
'Private mvarDirection As Boolean
'Private mvarJumpNextLevel As Boolean 'local copy
Private mvarNonSolid As Long 'local copy
Private mvarDieSnd As String 'local copy
Private mvarCoordX As Long 'local copy
Private mvarCoordY As Long 'local copy
Private mvarStartAnim As Boolean 'local copy
Private mvarcHeight As Long
Private mvarcWidth As Long
Private mvarKeyIndex As Long

Public Property Let DieSnd(ByVal vData As String)
    mvarDieSnd = vData
End Property

Public Property Get DieSnd() As String
    DieSnd = mvarDieSnd
End Property

Public Function SetBlock(pChar As String) As String
    SetBlock = cLevel.SetBlock(CoordX, CoordY, pChar)
End Function

Public Property Let KeyIndex(ByVal vData As Long)
    mvarKeyIndex = vData
End Property

Public Property Get KeyIndex() As Long
    KeyIndex = mvarKeyIndex
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

Public Property Let StartAnim(ByVal vData As Boolean)
    mvarStartAnim = vData
End Property

Public Property Get StartAnim() As Boolean
    StartAnim = mvarStartAnim
End Property

Public Property Let CoordY(ByVal vData As Long)
    mvarCoordY = vData
End Property

Public Property Get CoordY() As Long
    CoordY = mvarCoordY
End Property

Public Property Let CoordX(ByVal vData As Long)
    mvarCoordX = vData
End Property

Public Property Get CoordX() As Long
    CoordX = mvarCoordX
End Property

Public Property Let NonSolid(ByVal vData As Long)
    mvarNonSolid = vData
End Property

Public Property Get NonSolid() As Long
    NonSolid = mvarNonSolid
End Property

'Public Property Let JumpNextLevel(ByVal vData As Boolean)
'    mvarJumpNextLevel = vData
'End Property

'Public Property Get JumpNextLevel() As Boolean
'    JumpNextLevel = mvarJumpNextLevel
'End Property

'Friend Property Let Direction(ByVal vData As Boolean)
'    mvarDirection = vData
'End Property

'Friend Property Get Direction() As Boolean
'    Direction = mvarDirection
'End Property

'Public Property Let CanBreak(ByVal vData As Boolean)
'    mvarCanBreak = vData
'End Property

'Public Property Get CanBreak() As Boolean
'    CanBreak = mvarCanBreak
'End Property

'Public Property Let CanBeBreaked(ByVal vData As Boolean)
'    mvarCanBeBreaked = vData
'End Property

'Public Property Get CanBeBreaked() As Boolean
'    CanBeBreaked = mvarCanBeBreaked
'End Property

'Public Property Let CanHitEnemies(ByVal vData As Boolean)
'    mvarCanHitEnemies = vData
'End Property

'Public Property Get CanHitEnemies() As Boolean
'    CanHitEnemies = mvarCanHitEnemies
'End Property

'Public Property Let Hibernating(ByVal vData As Boolean)
'    mvarHibernating = vData
'End Property

'Public Property Get Hibernating() As Boolean
'    Hibernating = mvarHibernating
'End Property

'Public Property Let NextLevel(ByVal vData As Boolean)
'    mvarNextLevel = vData
'End Property

'Public Property Get NextLevel() As Boolean
'    NextLevel = mvarNextLevel
'End Property

'Public Property Let ChangePlayerTo(ByVal vData As String)
'    mvarChangePlayerTo = vData
'End Property

'Public Property Get ChangePlayerTo() As String
'    ChangePlayerTo = mvarChangePlayerTo
'End Property

'Public Property Let CanHit(ByVal vData As Boolean)
'    mvarCanHit = vData
'End Property

'Public Property Get CanHit() As Boolean
'    CanHit = mvarCanHit
'End Property

'Friend Property Let JumpSize(ByVal vData As Long)
'    mvarJumpSize = vData
'End Property

'Friend Property Get JumpSize() As Long
'    JumpSize = mvarJumpSize
'End Property

'Public Property Let JumpTo(ByVal vData As Integer)
'    mvarJumpTo = vData
'End Property

'Public Property Get JumpTo() As Integer
'    JumpTo = mvarJumpTo
'End Property

'Public Property Let JumpFrom(ByVal vData As Integer)
'    mvarJumpFrom = vData
'End Property

'Public Property Get JumpFrom() As Integer
'    'JumpFrom = mvarJumpFrom
'End Property

'Public Property Let AI(ByVal vData As String)
'    mvarAI = vData
'End Property

'Public Property Get AI() As String
'    AI = mvarAI
'End Property

'Public Property Let DieTiming(ByVal vData As Long)
'    mvarDieTiming = vData
'End Property

'Public Property Get DieTiming() As Long
'    DieTiming = mvarDieTiming
'End Property

Public Property Let CharType(ByVal vData As String)
    mvarCharType = vData
End Property


Public Property Get CharType() As String
    CharType = mvarCharType
End Property

'Public Property Let InitAIWhenHitted(ByVal vData As Boolean)
'    mvarInitAIWhenHitted = vData
'End Property

'Public Property Get InitAIWhenHitted() As Boolean
'    InitAIWhenHitted = mvarInitAIWhenHitted
'End Property

'Public Property Let Fixed(ByVal vData As Boolean)
'    mvarFixed = vData
'End Property

'Public Property Get Fixed() As Boolean
'    Fixed = mvarFixed
'End Property

'Public Property Let Solid(ByVal vData As Boolean)
'    mvarSolid = vData
'End Property

'Public Property Get Solid() As Boolean
'    Solid = mvarSolid
'End Property

'Public Property Let MakeJumpWhenHitted(ByVal vData As Boolean)
'    mvarMakeJumpWhenHitted = vData
'End Property

'Public Property Get MakeJumpWhenHitted() As Boolean
'    MakeJumpWhenHitted = mvarMakeJumpWhenHitted
'End Property

'Public Property Let RemoveWhenDies(ByVal vData As Boolean)
'    mvarRemoveWhenDies = vData
'End Property

'Public Property Get RemoveWhenDies() As Boolean
'    RemoveWhenDies = mvarRemoveWhenDies
'End Property

'Public Property Let DieFrame(ByVal vData As Integer)
'    mvarDieFrame = vData
'End Property

'Public Property Get DieFrame() As Integer
'    DieFrame = mvarDieFrame
'End Property

'Public Property Let AnimTo(ByVal vData As Integer)
'    mvarAnimTo = vData
'End Property

'Public Property Get AnimTo() As Integer
'    AnimTo = mvarAnimTo
'End Property

'Public Property Let AnimFrom(ByVal vData As Integer)
'    mvarAnimFrom = vData
'End Property

'Public Property Get AnimFrom() As Integer
'    AnimFrom = mvarAnimFrom
'End Property

'Public Property Let CanFall(ByVal vData As Boolean)
'    mvarCanFall = vData
'End Property

'Public Property Get CanFall() As Boolean
'    CanFall = mvarCanFall
'End Property

'Public Property Let User(ByVal vData As Boolean)
'    mvarUser = vData
'End Property

'Public Property Get User() As Boolean
'     User = mvarUser
'End Property

Public Property Let Visible(ByVal vData As Boolean)
    mvarVisible = vData
End Property

Public Property Get Visible() As Boolean
    Visible = mvarVisible
End Property

'Public Property Let CreatePlace(ByVal vData As String)
'    mvarCreatePlace = vData
'End Property
'
'Public Property Get CreatePlace() As String
'     CreatePlace = mvarCreatePlace
'End Property

'Public Property Let CreateWhenHitted(ByVal vData As String)
'    mvarCreateWhenHitted = vData
'End Property

'Public Property Get CreateWhenHitted() As String
'     CreateWhenHitted = mvarCreateWhenHitted
'End Property

Public Property Let PosBottom(ByVal vData As Long)
    mvarPosBottom = vData
    mvarcHeight = Abs(mvarPosBottom - mvarPosTop)
End Property

Public Property Get PosBottom() As Long
     PosBottom = mvarPosBottom
End Property

Public Property Let PosTop(ByVal vData As Long)
    mvarPosTop = vData
    mvarcHeight = Abs(mvarPosBottom - mvarPosTop)
End Property

Public Property Get PosTop() As Long
     PosTop = mvarPosTop
End Property

Public Property Let PosRight(ByVal vData As Long)
    mvarPosRight = vData
    mvarcWidth = Abs(mvarPosRight - mvarPosLeft)
End Property

Public Property Get PosRight() As Long
     PosRight = mvarPosRight
End Property

Public Property Let PosLeft(ByVal vData As Long)
    mvarPosLeft = vData
    mvarcWidth = Abs(mvarPosRight - mvarPosLeft)
End Property

Public Property Get PosLeft() As Long
     PosLeft = mvarPosLeft
End Property

'Public Property Let InitedAI(ByVal vData As Boolean)
'    mvarInitedAIBoolean = vData
'End Property

'Public Property Get InitedAI() As Boolean
'    InitedAI = mvarInitedAIBoolean
'End Property

'Public Property Let JumpVelocity(ByVal vData As Long)
'    mvarJumpVelocity = vData
'End Property

'Public Property Get JumpVelocity() As Long
'     JumpVelocity = mvarJumpVelocity
'End Property

'Public Property Let Velocity(ByVal vData As Long)
'    mvarVelocity = vData
'End Property

'Public Property Get Velocity() As Long
'     Velocity = mvarVelocity
'End Property

'Public Property Let MaxRunframe(ByVal vData As Long)
'    mvarMaxRunframe = vData
'End Property

'Public Property Get MaxRunframe() As Long
'     MaxRunframe = mvarMaxRunframe
'End Property

'Public Property Let MinRunframe(ByVal vData As Long)
'    mvarMinRunframe = vData
'End Property

'Public Property Get MinRunframe() As Long
'     MinRunframe = mvarMinRunframe
'End Property

Public Property Let SourceRight(ByVal vData As Long)
    mvarSourceRight = vData
End Property

Public Property Get SourceRight() As Long
     SourceRight = mvarSourceRight
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

Public Property Let SourceLeft(ByVal vData As Long)
    mvarSourceLeft = vData
End Property

Public Property Get SourceLeft() As Long
     SourceLeft = mvarSourceLeft
End Property

Public Property Let HittedByRight(ByVal vData As Boolean)
    mvarHittedByRight = vData
End Property

Public Property Get HittedByRight() As Boolean
     HittedByRight = mvarHittedByRight
End Property

Public Property Let HittedByLeft(ByVal vData As Boolean)
    mvarHittedByLeft = vData
End Property

Public Property Get HittedByLeft() As Boolean
     HittedByLeft = mvarHittedByLeft
End Property

Public Property Let HittedByBottom(ByVal vData As Boolean)
    mvarHittedByBottom = vData
End Property

Public Property Get HittedByBottom() As Boolean
     HittedByBottom = mvarHittedByBottom
End Property

Public Property Let HittedByTop(ByVal vData As Boolean)
    mvarHittedByTop = vData
End Property

Public Property Get HittedByTop() As Boolean
    HittedByTop = mvarHittedByTop
End Property

Public Property Let hdc(ByVal vData As Long)
    mvarhdc = vData
End Property

Public Property Get hdc() As Long
    hdc = mvarhdc
End Property

Public Sub Draw(RunFrame As Long, OffsetX As Long)
Dim sWidth As Long, sHeight As Long
Static mvarStep
sWidth = mvarSourceRight - mvarSourceLeft
sHeight = mvarSourceBottom - mvarSourceTop

        If mvarStep >= 50 Then
            mvarStartAnim = False
            Visible = False
        End If
        mvarStep = mvarStep + 1
        rctSrc.left = mvarSourceLeft + sWidth / 2
        rctSrc.right = mvarSourceLeft + sWidth
        rctSrc.top = mvarSourceTop
        rctSrc.bottom = mvarSourceTop + sHeight / 2
        rctDst.top = mvarPosTop + ((mvarStep / 2) - 9) ^ 2 - 81
        rctDst.bottom = rctDst.top + mvarcHeight / 2
        rctDst.left = mvarPosLeft - OffsetX + mvarStep * 3
        rctDst.right = rctDst.left + mvarcWidth / 2
        'If mvarFixed Or mvarSolid Then
        '    StretchPaint Form1.pForm1.hdc, rctDst, hdc, rctSrc
        'Else
            TransparentBlt Form1.pForm1.hdc, rctDst, hdc, rctSrc, vbMagenta, 0
        'End If
        
        rctSrc.left = mvarSourceLeft + sWidth / 2
        rctSrc.right = mvarSourceLeft + sWidth
        rctSrc.top = mvarSourceTop + sHeight / 2
        rctSrc.bottom = mvarSourceTop + sHeight
        rctDst.top = mvarPosTop + ((mvarStep / 2) - 5) ^ 2 - 25
        rctDst.bottom = rctDst.top + mvarcHeight / 2
        rctDst.left = mvarPosLeft - OffsetX + mvarStep * 3
        rctDst.right = rctDst.left + mvarcWidth / 2
        'If mvarFixed Or mvarSolid Then
        '    StretchPaint Form1.pForm1.hdc, rctDst, hdc, rctSrc
        'Else
            TransparentBlt Form1.pForm1.hdc, rctDst, hdc, rctSrc, vbMagenta, 0
        'End If
        '
        rctSrc.top = mvarSourceTop
        rctSrc.bottom = mvarSourceTop + sHeight / 2
        rctSrc.left = mvarSourceLeft
        rctSrc.right = mvarSourceLeft + sWidth / 2
        rctDst.top = mvarPosTop + ((mvarStep / 2) - 9) ^ 2 - 81
        rctDst.bottom = rctDst.top + mvarcHeight / 2
        rctDst.left = mvarPosLeft - OffsetX - mvarStep * 3
        rctDst.right = rctDst.left + mvarcWidth / 2
        'If mvarFixed Or mvarSolid Then
        '    StretchPaint Form1.pForm1.hdc, rctDst, hdc, rctSrc
        'Else
            TransparentBlt Form1.pForm1.hdc, rctDst, hdc, rctSrc, vbMagenta, 0
        'End If
        
        rctSrc.top = mvarSourceTop + sHeight / 2
        rctSrc.bottom = mvarSourceBottom
        rctSrc.left = mvarSourceLeft
        rctSrc.right = mvarSourceLeft + sWidth / 2
        rctDst.top = mvarPosTop + ((mvarStep / 2) - 5) ^ 2 - 25
        rctDst.bottom = rctDst.top + mvarcHeight / 2
        rctDst.left = mvarPosLeft - OffsetX - mvarStep * 3
        rctDst.right = rctDst.left + mvarcWidth / 2
        'If mvarFixed Or mvarSolid Then
        '    StretchPaint Form1.pForm1.hdc, rctDst, hdc, rctSrc
        'Else
            TransparentBlt Form1.pForm1.hdc, rctDst, hdc, rctSrc, vbMagenta, 0
        'End If
    
End Sub



