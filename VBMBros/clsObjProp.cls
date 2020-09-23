VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarhdc As Long 'copia local
Private mvarHittedByTop As Boolean 'copia local
Private mvarFireBallCanHit As Boolean 'copia local
Private mvarHittedByBottom As Boolean 'copia local
Private mvarHittedByLeft As Boolean 'copia local
Private mvarHittedByRight As Boolean 'copia local
Private mvarSourceLeft As Long 'copia local
Private mvarSourceTop As Long 'copia local
Private mvarSourceBottom As Long 'copia local
Private mvarSourceRight As Long 'copia local
Private mvarMinRunframe As Long 'copia local
Private mvarMaxRunframe As Long 'copia local
Private mvarVelocity As Long 'copia local
Private mvarJumpVelocity As Long 'copia local
Private mvarInitedAIBoolean As Boolean 'copia local
Private mvarPosLeft As Long 'copia local
Private mvarPosRight As Long 'copia local
Private mvarPosTop As Long 'copia local
Private mvarPosBottom As Long 'copia local
Private mvarCreateWhenHitted As String 'copia local
Private mvarCreatePlace As String 'copia local
Private mvarVisible As Boolean 'copia local
Private mvarUser As Boolean 'copia local
Private mvarCanFall As Boolean 'copia local
Private mvarStartPosOffsetX As Long
Private mvarAnimFrom As Integer 'copia local
Private mvarAnimTo As Integer 'copia local
Private mvarDieFrame As Integer 'copia local
Private mvarRemoveWhenDies As Boolean 'copia local
Private mvarMakeJumpWhenHitted As Boolean 'copia local
Private mvarJumpWhenHitted As Boolean 'copia local
Private mvarSolid As Boolean 'copia local
Private mvarFixed As Boolean 'copia local
Private mvarInitAIWhenHitted As Boolean 'local copy
Private mvarCharType As String 'local copy
Private mvarDieTiming As Long 'local copy
Private mvarAI As String 'local copy
Private mvarJumpFrom As Integer 'local copy
Private mvarChangeFrom As Integer 'local copy
Private mvarJumpTo As Integer 'local copy
Private mvarJumpSize As Long
Private mvarCanHit As Boolean 'local copy
'Private mvarFireBall As Boolean 'local copy
Private mvarChangePlayerTo As String 'local copy
Private mvarNextLevel As Boolean 'local copy
Private mvarHibernating As Boolean 'local copy
Private mvarCanHitEnemies As Boolean 'local copy
Private mvarCannotHitUser As Boolean 'local copy
Private mvarCanBeBreaked As Boolean 'local copy
Private mvarCanBreak As Boolean 'local copy
Private mvarDirection As Boolean
Private mvarJumpNextLevel As Boolean 'local copy
Private mvarNonSolid As Long 'local copy
Private mvarUserSelection As Boolean 'local copy
Private mvarGrowTo As String 'local copy
Private mvarMakeGrow As Boolean 'local copy
Private mvarIsFloor As Boolean 'local copy
Private mvarJumpSnd As String 'local copy
Private mvarFireSnd As String 'local copy
Private mvarDieSnd As String 'local copy
Private mvarDieWhenHits As Boolean 'local copy
Private mvarFireBall As String 'local copy
Private mvarRaising As Boolean
Private mvarDescending As Boolean
Private mvarRaisetime As Long
'variables locales para almacenar los valores de las propiedades
Private mvarPicture As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarCrouchFrame As Integer 'copia local
Private mvarCanCrouch As Boolean 'copia local
'variables locales para almacenar los valores de las propiedades
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



Public Property Let Picture(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Picture = 5
    mvarPicture = vData
End Property


Public Property Get Picture() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Picture
    Picture = mvarPicture
End Property




Friend Property Get Raisetime() As Long
    Raisetime = mvarRaisetime
End Property

Friend Property Let Raisetime(ByVal vData As Long)
    mvarRaisetime = vData
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

Public Property Let FireBall(ByVal vData As String)
    mvarFireBall = vData
End Property

Public Property Get FireBall() As String
    FireBall = mvarFireBall
End Property

Public Property Let DieWhenHits(ByVal vData As Boolean)
    mvarDieWhenHits = vData
End Property

Public Property Get DieWhenHits() As Boolean
    DieWhenHits = mvarDieWhenHits
End Property

Public Property Let DieSnd(ByVal vData As String)
    mvarDieSnd = vData
End Property

Public Property Get DieSnd() As String
    DieSnd = mvarDieSnd
End Property

Public Property Let FireSnd(ByVal vData As String)
    mvarFireSnd = vData
End Property

Public Property Get FireSnd() As String
    FireSnd = mvarFireSnd
End Property

Public Property Let JumpSnd(ByVal vData As String)
    mvarJumpSnd = vData
End Property

Public Property Get JumpSnd() As String
    JumpSnd = mvarJumpSnd
End Property

Public Property Let IsFloor(ByVal vData As Boolean)
    mvarIsFloor = vData
End Property

Public Property Get IsFloor() As Boolean
    IsFloor = mvarIsFloor
End Property

Public Property Let MakeGrow(ByVal vData As Boolean)
    mvarMakeGrow = vData
End Property

Public Property Get MakeGrow() As Boolean
    MakeGrow = mvarMakeGrow
End Property

Public Property Let GrowTo(ByVal vData As String)
    mvarGrowTo = vData
End Property

Public Property Get GrowTo() As String
    GrowTo = mvarGrowTo
End Property

Public Property Let UserSelection(ByVal vData As Boolean)
    mvarUserSelection = vData
End Property

Public Property Get UserSelection() As Boolean
    UserSelection = mvarUserSelection
End Property

Public Property Let NonSolid(ByVal vData As Long)
    mvarNonSolid = vData
End Property

Public Property Get NonSolid() As Long
    NonSolid = mvarNonSolid
End Property

Public Property Let JumpNextLevel(ByVal vData As Boolean)
    mvarJumpNextLevel = vData
End Property

Public Property Get JumpNextLevel() As Boolean
    JumpNextLevel = mvarJumpNextLevel
End Property

Friend Property Let direction(ByVal vData As Boolean)
    mvarDirection = vData
End Property

Friend Property Get direction() As Boolean
    direction = mvarDirection
End Property

Public Property Let CanBreak(ByVal vData As Boolean)
    mvarCanBreak = vData
End Property

Public Property Get CanBreak() As Boolean
    CanBreak = mvarCanBreak
End Property

Public Property Let CanBeBreaked(ByVal vData As Boolean)
    mvarCanBeBreaked = vData
End Property

Public Property Get CanBeBreaked() As Boolean
    CanBeBreaked = mvarCanBeBreaked
End Property

Public Property Let CannotHitUser(ByVal vData As Boolean)
    mvarCannotHitUser = vData
End Property

Public Property Get CannotHitUser() As Boolean
    CannotHitUser = mvarCannotHitUser
End Property

Public Property Let CanHitEnemies(ByVal vData As Boolean)
    mvarCanHitEnemies = vData
End Property

Public Property Get CanHitEnemies() As Boolean
    CanHitEnemies = mvarCanHitEnemies
End Property

Public Property Let Hibernating(ByVal vData As Boolean)
    mvarHibernating = vData
End Property

Public Property Get Hibernating() As Boolean
    Hibernating = mvarHibernating
End Property

Public Property Let NextLevel(ByVal vData As Boolean)
    mvarNextLevel = vData
End Property

Public Property Get NextLevel() As Boolean
    NextLevel = mvarNextLevel
End Property

Public Property Let ChangePlayerTo(ByVal vData As String)
    mvarChangePlayerTo = vData
End Property

Public Property Get ChangePlayerTo() As String
    ChangePlayerTo = mvarChangePlayerTo
End Property

'Public Property Let FireBall(ByVal vData As Boolean)
'    mvarFireBall = vData
'End Property'

'Public Property Get FireBall() As Boolean
'    FireBall = mvarFireBall
'End Property

Public Property Let CanHit(ByVal vData As Boolean)
    mvarCanHit = vData
End Property

Public Property Get CanHit() As Boolean
    CanHit = mvarCanHit
End Property

Friend Property Let JumpSize(ByVal vData As Long)
    mvarJumpSize = vData
End Property

Friend Property Get JumpSize() As Long
    JumpSize = mvarJumpSize
End Property

Public Property Let JumpTo(ByVal vData As Integer)
    mvarJumpTo = vData
End Property

Public Property Get JumpTo() As Integer
    JumpTo = mvarJumpTo
End Property

Public Property Let ChangeFrom(ByVal vData As Integer)
    mvarChangeFrom = vData
End Property

Public Property Get ChangeFrom() As Integer
    ChangeFrom = mvarChangeFrom
End Property

Public Property Let JumpFrom(ByVal vData As Integer)
    mvarJumpFrom = vData
End Property

Public Property Get JumpFrom() As Integer
    JumpFrom = mvarJumpFrom
End Property

Public Property Let AI(ByVal vData As String)
    mvarAI = vData
End Property

Public Property Get AI() As String
    AI = mvarAI
End Property

Public Property Let DieTiming(ByVal vData As Long)
    mvarDieTiming = vData
End Property

Public Property Get DieTiming() As Long
    DieTiming = mvarDieTiming
End Property

Public Property Let CharType(ByVal vData As String)
    mvarCharType = vData
End Property


Public Property Get CharType() As String
    CharType = mvarCharType
End Property

Public Property Let InitAIWhenHitted(ByVal vData As Boolean)
    mvarInitAIWhenHitted = vData
End Property

Public Property Get InitAIWhenHitted() As Boolean
    InitAIWhenHitted = mvarInitAIWhenHitted
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

Public Property Let AnimTo(ByVal vData As Integer)
    mvarAnimTo = vData
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

Public Property Let User(ByVal vData As Boolean)
    mvarUser = vData
End Property

Public Property Get User() As Boolean
     User = mvarUser
End Property

Public Property Let Visible(ByVal vData As Boolean)
    mvarVisible = vData
End Property

Public Property Get Visible() As Boolean
    Visible = mvarVisible
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

Public Property Let PosBottom(ByVal vData As Long)
    mvarPosBottom = vData
End Property

Public Property Get PosBottom() As Long
     PosBottom = mvarPosBottom
End Property

Public Property Let PosTop(ByVal vData As Long)
    mvarPosTop = vData
End Property

Public Property Get PosTop() As Long
     PosTop = mvarPosTop
End Property

Public Property Let PosRight(ByVal vData As Long)
    mvarPosRight = vData
End Property

Public Property Get PosRight() As Long
     PosRight = mvarPosRight
End Property

Public Property Let PosLeft(ByVal vData As Long)
    mvarPosLeft = vData
End Property

Public Property Get PosLeft() As Long
     PosLeft = mvarPosLeft
End Property

Public Property Let InitedAI(ByVal vData As Boolean)
    mvarInitedAIBoolean = vData
End Property

Public Property Get InitedAI() As Boolean
    InitedAI = mvarInitedAIBoolean
End Property

Public Property Let JumpVelocity(ByVal vData As Long)
    mvarJumpVelocity = vData
End Property

Public Property Get JumpVelocity() As Long
     JumpVelocity = mvarJumpVelocity
End Property

Public Property Let Velocity(ByVal vData As Long)
    mvarVelocity = vData
End Property

Public Property Get Velocity() As Long
     Velocity = mvarVelocity
End Property

Public Property Let MaxRunframe(ByVal vData As Long)
    mvarMaxRunframe = vData
End Property

Public Property Get MaxRunframe() As Long
     MaxRunframe = mvarMaxRunframe
End Property

Public Property Let MinRunframe(ByVal vData As Long)
    mvarMinRunframe = vData
End Property

Public Property Get MinRunframe() As Long
     MinRunframe = mvarMinRunframe
End Property

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

Public Property Let FireBallCanHit(ByVal vData As Boolean)
    mvarFireBallCanHit = vData
End Property

Public Property Get FireBallCanHit() As Boolean
    FireBallCanHit = mvarFireBallCanHit
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



