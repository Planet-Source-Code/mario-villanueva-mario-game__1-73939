VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements clsObjBasics
Implements clsAI

Private jumpcnt As Single, gravity As Single, endjump As Long
Private mvartotalJump As Single
Private fallcnt As Long
Private mMaxRunframe As Integer
Private mMinRunframe As Integer
Private mvarAnimFrom As Integer
Private mvarAnimTo   As Integer
Private mvarChangeFrom As Integer
Private mvarJumpFrom As Integer
Private mvarJumpTo   As Integer
Private mclsObjBasics As New clsObjBasics
Private mvarHittedByLeft As Boolean
Private mvarHittedByRight As Boolean
Private mvarFireBallCanHit As Boolean
Private mvarHittedByTop As Boolean
Private mvarHittedByBottom As Boolean
Private mvarDie As Boolean
Private mvarUser As Boolean
Private mvarJumping As Boolean
Private mvarFalling As Boolean
Private mvarMoving As Boolean
Private mClsAI As New clsAI
Private mvarCreateWhenHitted As String
Private mvarCreatePlace As String
Private mvarInercia As Long
Private mvarInitAIWhenHitted As Boolean 'local copy
Private mvarCharType As String 'local copy
Private mvarRunFrame As Long 'local copy
Private mvarDieTiming As Long 'local copy
Private mvarCanHit As Boolean 'local copy
Private mvarFireBall As String 'local copy
Private mvarChangePlayerTo As String 'local copy
Private mvarForceJump As Boolean 'local copy
Private mvarNextLevel As Boolean 'local copy
Private mvarJumpNextLevel As Boolean 'local copy
Private mvarHibernating As Boolean 'local copy
Private mvarCanHitEnemies As Boolean 'local copy
Private mvarCannotHitUser As Boolean 'local copy
Private mvarCanBeBreaked As Boolean 'local copy
Private mvarCanBreak As Boolean 'local copy
Private mvarInScreen As Boolean 'local copy
Private mvarNextLevelChar As String 'local copy
Private mvarID As Long 'local copy
Private mvarjumpStart As Boolean
Private mvarUserSelection As Boolean
Private mvarGrowTo As String 'local copy
Private mvarMakeGrow As Boolean 'local copy
Private mvarIsFloor As Boolean 'local copy
Private mvarIsFireBall As Boolean 'local copy
Private mvarJumpSnd As String 'local copy
Private mvarFireSnd As String 'local copy
Private mvarDieSnd As String 'local copy
Private mvarDieWhenHits As Boolean 'local copy
Private mvarKey As String 'local copy
Private mvarStand As Long
Private srcJump As Long
'local variable(s) to hold property value(s)
Private mvarCarried As Boolean 'local copy
'variables locales para almacenar los valores de las propiedades
Private mvarAmnisty As Long 'copia local
Public Property Let Amnisty(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Amnisty = 5
    mvarAmnisty = vData
End Property


Public Property Get Amnisty() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Amnisty
    Amnisty = mvarAmnisty
End Property



Public Property Let Carried(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Carried = 5
    mvarCarried = vData
End Property


Public Property Get Carried() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Carried
    Carried = mvarCarried
End Property

Public Property Get Crouched() As Boolean
    Crouched = mclsObjBasics.Crouched
End Property

Public Property Let Crouched(ByVal RHS As Boolean)
    mclsObjBasics.Crouched = RHS
End Property

Public Property Get CanCrouch() As Boolean
    CanCrouch = mclsObjBasics.CanCrouch
End Property

Public Property Let CanCrouch(ByVal RHS As Boolean)
    mclsObjBasics.CanCrouch = RHS
End Property

Public Property Get CrouchFrame() As Integer
    CrouchFrame = mclsObjBasics.CrouchFrame
End Property

Public Property Let CrouchFrame(ByVal RHS As Integer)
    mclsObjBasics.CrouchFrame = RHS
End Property

Friend Property Get Raisetime() As Long
    Raisetime = mclsObjBasics.Raisetime()
End Property

Friend Property Let Raisetime(ByVal vData As Long)
    mclsObjBasics.Raisetime = vData
End Property

Friend Property Get Raising() As Long
    Raising = mclsObjBasics.Raising()
End Property

Friend Property Let Raising(ByVal vData As Long)
    mclsObjBasics.Raising = vData
End Property

Friend Property Get Descending() As Long
    Descending = mclsObjBasics.Descending()
End Property

Friend Property Let Descending(ByVal vData As Long)
    mclsObjBasics.Descending = vData
End Property

Public Property Let Stand(ByVal vData As Long)
    mvarStand = vData
End Property

Public Property Get Stand() As Long
    Stand = mvarStand
End Property

Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property

Public Property Get Key() As String
    Key = mvarKey
End Property

Public Property Let IsFireBall(ByVal vData As Boolean)
    mvarIsFireBall = vData
End Property

Public Property Get IsFireBall() As Boolean
    IsFireBall = mvarIsFireBall
End Property

Public Property Let FireSnd(ByVal vData As String)
    mvarFireSnd = vData
End Property

Public Property Get FireSnd() As String
    FireSnd = mvarFireSnd
End Property

'Public Property Let FireBall(ByVal vData As Boolean)
'    mvarFireBall = vData
'End Property

'Public Property Get FireBall() As Boolean
'    FireBall = mvarFireBall
'End Property

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

Public Property Let JumpStart(ByVal vData As Boolean)
    mvarjumpStart = vData
End Property

Public Property Get JumpStart() As Boolean
    JumpStart = mvarjumpStart
End Property

Public Property Let ID(ByVal vData As Long)
    mvarID = vData
End Property

Public Property Get ID() As Long
    ID = mvarID
End Property

Public Property Let NextLevelChar(ByVal vData As String)
    mvarNextLevelChar = vData
End Property

Public Property Get NextLevelChar() As String
    NextLevelChar = mvarNextLevelChar
End Property

Public Property Let InScreen(ByVal vData As Boolean)
    mvarInScreen = vData
End Property

Public Property Get InScreen() As Boolean
    InScreen = mvarInScreen
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

Public Property Let JumpNextLevel(ByVal vData As Boolean)
    mvarJumpNextLevel = vData
End Property

Public Property Get JumpNextLevel() As Boolean
    JumpNextLevel = mvarJumpNextLevel
End Property

Public Property Let NextLevel(ByVal vData As Boolean)
    mvarNextLevel = vData
End Property

Public Property Get NextLevel() As Boolean
    NextLevel = mvarNextLevel
End Property

Public Property Let ForceJump(ByVal vData As Boolean)
    mvarForceJump = vData
End Property

Public Property Get ForceJump() As Boolean
    ForceJump = mvarForceJump
End Property

Public Property Let ChangePlayerTo(ByVal vData As String)
    mvarChangePlayerTo = vData
End Property

Public Property Get ChangePlayerTo() As String
    ChangePlayerTo = mvarChangePlayerTo
End Property

Public Property Let FireBall(ByVal vData As String)
    mvarFireBall = vData
End Property '

Public Property Get FireBall() As String
    FireBall = mvarFireBall
End Property

Public Property Let CanHit(ByVal vData As Boolean)
    mvarCanHit = vData
End Property

Public Property Get CanHit() As Boolean
    CanHit = mvarCanHit
End Property

Public Property Get JumpSize() As Long
    JumpSize = mclsObjBasics.JumpSize
End Property

Public Property Let JumpSize(ByVal vData As Long)
    mclsObjBasics.JumpSize = vData
End Property

Public Property Let AI(ByVal vData As String)
    mClsAI.AI = vData
End Property

Public Property Get AI() As String
    AI = mClsAI.AI
End Property

Public Property Let Jumping(ByVal vData As Boolean)
    mvarJumping = vData
End Property

Public Property Get Jumping() As Boolean
    Jumping = mvarJumping
End Property

Public Property Let Falling(ByVal vData As Boolean)
    mvarFalling = vData
End Property

Public Property Get Falling() As Boolean
    Falling = mvarFalling
End Property

Public Property Let DieTiming(ByVal vData As Long)
    mvarDieTiming = vData
End Property

Public Property Get DieTiming() As Long
    DieTiming = mvarDieTiming
End Property

Public Property Let RunFrame(ByVal vData As Long)
    mvarRunFrame = vData
End Property

Public Property Get RunFrame() As Long
    RunFrame = mvarRunFrame
End Property

Public Property Let Breaking(ByVal vData As Boolean)
    mclsObjBasics.Breaking = vData
End Property

Public Property Get Breaking() As Boolean
    Breaking = mclsObjBasics.Breaking
End Property

Public Property Let CharType(ByVal vData As String)
    mclsObjBasics.CharType = vData
End Property

Public Property Get CharType() As String
    CharType = mclsObjBasics.CharType
End Property

Public Property Let InitAIWhenHitted(ByVal vData As Boolean)
    mvarInitAIWhenHitted = vData
End Property

Public Property Get InitAIWhenHitted() As Boolean
    InitAIWhenHitted = mvarInitAIWhenHitted
End Property

Public Property Let Inercia(ByVal vData As Long)
    mvarInercia = vData
End Property

Public Property Get Inercia() As Long
    Inercia = mvarInercia
End Property

Private Sub Class_Initialize()
    mclsObjBasics.Velocity = 2
End Sub

Public Property Let Fixed(ByVal vData As Boolean)
    mclsObjBasics.Fixed = vData
End Property

Public Property Get Fixed() As Boolean
    Fixed = mclsObjBasics.Fixed
End Property

Public Property Let Solid(ByVal vData As Boolean)
    mclsObjBasics.Solid = vData
End Property

Public Property Get Solid() As Boolean
    Solid = mclsObjBasics.Solid
End Property

Public Property Let JumpWhenHitted(ByVal vData As Boolean)
    mclsObjBasics.JumpWhenHitted = vData
End Property

Public Property Get JumpWhenHitted() As Boolean
    JumpWhenHitted = mclsObjBasics.JumpWhenHitted
End Property

Public Property Let MakeJumpWhenHitted(ByVal vData As Boolean)
    mclsObjBasics.MakeJumpWhenHitted = vData
End Property

Public Property Get MakeJumpWhenHitted() As Boolean
    MakeJumpWhenHitted = mclsObjBasics.MakeJumpWhenHitted
End Property

Public Property Get RemoveWhenDies() As Boolean
    RemoveWhenDies = mclsObjBasics.RemoveWhenDies
End Property

Public Property Let RemoveWhenDies(ByVal vData As Boolean)
    mclsObjBasics.RemoveWhenDies = vData
End Property

Public Property Get DieFrame() As Integer
    DieFrame = mclsObjBasics.DieFrame
End Property

Public Property Let DieFrame(ByVal vData As Integer)
    mclsObjBasics.DieFrame = vData
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

Public Property Let JumpTo(ByVal vData As Integer)
    mvarJumpTo = vData
End Property

Public Property Get JumpTo() As Integer
    JumpTo = mvarJumpTo
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

Friend Property Let User(ByVal vData As Boolean)
    mvarUser = vData
End Property

Friend Property Get User() As Boolean
    User = mvarUser
End Property

Friend Property Get HittedByLeft() As Long
    HittedByLeft = mvarHittedByLeft
End Property

Friend Property Let HittedByLeft(ByVal vData As Long)
    mvarHittedByLeft = vData
End Property

Friend Property Get FireBallCanHit() As Long
    FireBallCanHit = mvarFireBallCanHit
End Property

Friend Property Let FireBallCanHit(ByVal vData As Long)
    mvarFireBallCanHit = vData
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
    SourceBottom = mclsObjBasics.SourceBottom()
End Property

Friend Property Let SourceBottom(ByVal vData As Long)
    mclsObjBasics.SourceBottom = vData
End Property

Friend Property Get StartPosOffsetX() As Long
    StartPosOffsetX = mclsObjBasics.StartPosOffsetX()
End Property

Friend Property Let StartPosOffsetX(ByVal vData As Long)
    mclsObjBasics.StartPosOffsetX = vData
End Property

Friend Property Get CanFall() As Boolean
    CanFall = mclsObjBasics.CanFall()
End Property

Friend Property Let CanFall(ByVal vData As Boolean)
    mclsObjBasics.CanFall = vData
End Property

Friend Property Get SourceTop() As Long
    SourceTop = mclsObjBasics.SourceTop()
End Property

Friend Property Let SourceTop(ByVal vData As Long)
    mclsObjBasics.SourceTop = vData
End Property

Friend Property Get SourceRight() As Long
    SourceRight = mclsObjBasics.SourceRight()
End Property

Friend Property Let SourceRight(ByVal vData As Long)
    mclsObjBasics.SourceRight = vData
End Property

Friend Property Get SourceLeft() As Long
    SourceLeft = mclsObjBasics.SourceLeft()
End Property

Friend Property Let SourceLeft(ByVal vData As Long)
    mclsObjBasics.SourceLeft = vData
End Property

Public Sub StopMotion()
    mvarRunFrame = mvarAnimFrom
    mvarMoving = False
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

Friend Property Get Hidden() As Boolean
    Hidden = mclsObjBasics.Hidden()
End Property

Friend Property Let Hidden(ByVal vData As Boolean)
    mclsObjBasics.Hidden = vData
End Property

Friend Property Get Visible() As Boolean
    Visible = mclsObjBasics.Visible()
End Property

Friend Property Let Visible(ByVal vData As Boolean)
    mclsObjBasics.Visible = vData
End Property

Friend Property Get RScreenLimit() As Long
    RScreenLimit = mclsObjBasics.RScreenLimit()
End Property

Friend Property Let RScreenLimit(ByVal vData As Long)
    mclsObjBasics.RScreenLimit = vData
End Property

Friend Property Let PosBottom(ByVal vData As Long)
    mclsObjBasics.PosBottom = vData
End Property

Friend Property Get PosBottom() As Long
    PosBottom = mclsObjBasics.PosBottom()
End Property

Friend Property Let PosTop(ByVal vData As Long)
    mclsObjBasics.PosTop = vData
End Property

Friend Property Get PosTop() As Long
    PosTop = mclsObjBasics.PosTop()
End Property

Friend Property Let PosRight(ByVal vData As Long)
    mclsObjBasics.PosRight = vData
End Property

Friend Property Get PosRight() As Long
    PosRight = mclsObjBasics.PosRight()
End Property

Friend Property Let PosLeft(ByVal vData As Long)
    mclsObjBasics.PosLeft = vData
End Property

Friend Property Get PosLeft() As Long
    PosLeft = mclsObjBasics.PosLeft
End Property

Friend Property Let direction(ByVal vData As Boolean)
    mclsObjBasics.direction = vData
End Property

Friend Property Get direction() As Boolean
    direction = mclsObjBasics.direction
End Property

Friend Property Let JumpVelocity(ByVal vData As Long)
    mclsObjBasics.JumpVelocity = vData
End Property

Friend Property Get JumpVelocity() As Long
    JumpVelocity = mclsObjBasics.JumpVelocity
End Property

Friend Property Let Velocity(ByVal vData As Long)
    mclsObjBasics.Velocity = vData
End Property

Friend Property Get Velocity() As Long
    Velocity = mclsObjBasics.Velocity
End Property

Friend Property Let Jump(ByVal vData As Boolean)
Dim newPos As Long
    If vData Then
        If Not ForceJump And JumpStart Then
            Playwav JumpSnd
        End If
    End If
    If OnFloor(newPos) Or ForceJump Or Form1.picTerrain.BackColor <> Form1.picTerrain.Point(PosLeft + 8, PosBottom + 3) Then
        mclsObjBasics.Jump = vData
        If newPos = 0 Then
            'mclsObjBasics.PosBottom = (((mclsObjBasics.PosBottom \ 16)) * 16)
        Else
            mclsObjBasics.PosBottom = newPos
        End If
        If ForceJump Then
            JumpStart = True
        End If
    End If
    'Debug.Print (Form1.picTerrain.BackColor <> Form1.picTerrain.Point(PosLeft + 8, PosBottom + 2))
    'ForceJump = False
    If vData = False Then
        mclsObjBasics.Jump = vData
    End If
End Property

Friend Property Get Jump() As Boolean
    Jump = mclsObjBasics.Jump
End Property

Friend Property Let hdc(ByVal vData As Long)
    mclsObjBasics.hdc = vData
End Property

Friend Property Get hdc() As Long
    hdc = mclsObjBasics.hdc
End Property

Public Sub Draw(OffsetX As Long)
    mclsObjBasics.Draw mvarRunFrame, OffsetX
End Sub

Private Property Get clsAI_AI() As String

End Property

Private Property Let clsAI_AI(ByVal RHS As String)

End Property

Private Property Get clsObjBasics_Breaked() As Boolean

End Property

Private Property Let clsObjBasics_Breaked(ByVal RHS As Boolean)

End Property

Private Property Get clsObjBasics_Breaking() As Boolean

End Property

Private Property Let clsObjBasics_Breaking(ByVal RHS As Boolean)

End Property

Private Property Let clsObjBasics_CharType(ByVal RHS As String)

End Property

Private Property Get clsObjBasics_CharType() As String

End Property

Private Property Let clsObjBasics_cHeight(ByVal RHS As Long)

End Property

Private Property Get clsObjBasics_cHeight() As Long

End Property

Private Property Let clsObjBasics_cWidth(ByVal RHS As Long)

End Property

Private Property Get clsObjBasics_cWidth() As Long

End Property

Private Sub clsObjBasics_Draw(mvarRunFrame As Long, OffsetX As Long)
    mclsObjBasics.Draw mvarRunFrame, OffsetX
End Sub

Private Property Get clsObjBasics_Crouched() As Boolean
    Crouched = mclsObjBasics.Crouched
End Property

Private Property Let clsObjBasics_Crouched(ByVal RHS As Boolean)
    mclsObjBasics.Crouched = RHS
End Property

Private Property Get clsObjBasics_CanCrouch() As Boolean
    CanCrouch = mclsObjBasics.CanCrouch
End Property

Private Property Let clsObjBasics_CanCrouch(ByVal RHS As Boolean)
    mclsObjBasics.CanCrouch = RHS
End Property

Private Property Get clsObjBasics_CrouchFrame() As Integer
    CrouchFrame = mclsObjBasics.CrouchFrame
End Property

Private Property Let clsObjBasics_CrouchFrame(ByVal RHS As Integer)
    mclsObjBasics.CrouchFrame = RHS
End Property

Private Property Let clsObjBasics_JumpVelocity(ByVal RHS As Long)

End Property

Private Property Get clsObjBasics_JumpVelocity() As Long

End Property

Private Function clsObjBasics_MoveLeft(Optional direction As Long) As Long
    
End Function

Private Function clsObjBasics_MoveRight(Optional direction As Long) As Long
    
End Function

Private Property Let clsObjBasics_Raisetime(ByVal RHS As Long)

End Property

Private Property Get clsObjBasics_Raisetime() As Long

End Property

Private Property Let clsObjBasics_Raising(ByVal RHS As Boolean)

End Property

Private Property Get clsObjBasics_Raising() As Boolean

End Property

Private Property Let clsObjBasics_Descending(ByVal RHS As Boolean)

End Property

Private Property Get clsObjBasics_Descending() As Boolean

End Property

Private Property Get clsObjBasics_RScreenLimit() As Long
    clsObjBasics_RScreenLimit = mclsObjBasics.RScreenLimit()
End Property

Private Property Let clsObjBasics_RScreenLimit(ByVal vData As Long)
    mclsObjBasics.RScreenLimit = vData
End Property

Private Property Let clsObjBasics_Fixed(ByVal RHS As Boolean)

End Property

Private Property Get clsObjBasics_Fixed() As Boolean

End Property

Private Property Let clsObjBasics_JumpWhenHitted(ByVal RHS As Boolean)

End Property

Private Property Get clsObjBasics_JumpWhenHitted() As Boolean

End Property

Private Property Let clsObjBasics_MakeJumpWhenHitted(ByVal RHS As Boolean)

End Property

Private Property Get clsObjBasics_MakeJumpWhenHitted() As Boolean

End Property

Private Property Get clsAI_Inited() As Boolean

End Property

Private Property Let clsAI_Inited(ByVal RHS As Boolean)

End Property

Private Sub clsAI_PerformAI(mclsObject As clsObject)

End Sub

Private Property Let clsObjBasics_StartPosOffsetX(ByVal RHS As Long)

End Property

Private Property Get clsObjBasics_StartPosOffsetX() As Long

End Property

Private Property Let clsObjBasics_CanFall(ByVal RHS As Boolean)

End Property

Private Property Get clsObjBasics_CanFall() As Boolean

End Property

Private Property Let clsObjBasics_DieFrame(ByVal RHS As Integer)

End Property

Private Property Get clsObjBasics_DieFrame() As Integer

End Property

Private Property Let clsObjBasics_RemoveWhenDies(ByVal RHS As Boolean)

End Property

Private Property Get clsObjBasics_RemoveWhenDies() As Boolean

End Property

Private Property Let clsObjBasics_Solid(ByVal RHS As Boolean)

End Property

Private Property Get clsObjBasics_Solid() As Boolean

End Property

Private Property Let clsObjBasics_SourceBottom(ByVal RHS As Long)

End Property

Private Property Get clsObjBasics_SourceBottom() As Long

End Property

Private Property Let clsObjBasics_SourceLeft(ByVal RHS As Long)

End Property

Private Property Get clsObjBasics_SourceLeft() As Long

End Property

Private Property Let clsObjBasics_SourceRight(ByVal RHS As Long)

End Property

Private Property Get clsObjBasics_SourceRight() As Long

End Property

Private Property Let clsObjBasics_SourceTop(ByVal RHS As Long)

End Property

Private Property Get clsObjBasics_SourceTop() As Long

End Property

Private Property Let clsObjBasics_Hidden(ByVal RHS As Boolean)

End Property

Private Property Get clsObjBasics_Hidden() As Boolean

End Property

Private Property Let clsObjBasics_Visible(ByVal RHS As Boolean)

End Property

Private Property Get clsObjBasics_Visible() As Boolean

End Property

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'                       Implementaciones propias del personaje
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

Public Sub Animate(ByRef Col As Collection, ptime As Long)
Dim bot As clsObject, CoordY As Long, newPos As Long
Dim x As Single, y As Single
Static dieCounter As Long
Static mtime
If mtime > 30000 Then mtime = 0
mtime = mtime + ptime

'Matar al personaje
If mvarDie Then
    If mvarUser Then
        'Muere el Jugador
        If Not Form1.RollBackPlayer Then
            mvarRunFrame = mclsObjBasics.DieFrame
            CoordY = (mclsObjBasics.PosBottom \ 16) + 1
            If dieCounter = 0 Then dieCounter = CoordY - 5
            If CoordY >= dieCounter Then
                mclsObjBasics.PosTop = mclsObjBasics.PosTop - 4
            ElseIf CoordY > 16 Then
                mclsObjBasics.Visible = False
            ElseIf CoordY <= dieCounter Then
                mclsObjBasics.PosTop = mclsObjBasics.PosTop + 4
                dieCounter = dieCounter + 1
            End If
        Else
            mvarDie = False
        End If
    Else
        If Not mclsObjBasics.Breaking Then
            'Muere un Villano
            mvarRunFrame = mclsObjBasics.DieFrame
            If dieCounter >= mvarDieTiming Then
                If mclsObjBasics.RemoveWhenDies Then
                    mclsObjBasics.Visible = False
                    If LenB(CreateWhenHitted) Or LenB(ChangePlayerTo) Then
                        CreateChar ChangePlayerTo
                    End If
                Else
                    If LenB(CreateWhenHitted) Or LenB(ChangePlayerTo) Then
                        CreateChar ChangePlayerTo
                    End If
                End If
            Else
                If Me.JumpWhenHitted Then
                    mvarRunFrame = mclsObjBasics.DieFrame
                    CoordY = (mclsObjBasics.PosBottom \ 16) + 1
                    If dieCounter = 0 Then dieCounter = CoordY - 5
                    If CoordY >= dieCounter Then
                        mclsObjBasics.PosTop = mclsObjBasics.PosTop - 5
                    ElseIf CoordY > 16 Then
                        mclsObjBasics.Visible = False
                    ElseIf CoordY <= dieCounter Then
                        mclsObjBasics.PosTop = mclsObjBasics.PosTop + 5
                        dieCounter = dieCounter + 1
                    End If
                Else
                    dieCounter = dieCounter + 1
                End If
            End If
        End If
    End If
    Exit Sub
End If

If Not mclsObjBasics.Fixed And Not mclsObjBasics.Solid And Not mvarDie And (InitedAI Or mvarUser Or (Not InitedAI And Not Carried)) Then
' Hacer animacion de caida
If Not mclsObjBasics.Jump And Not mvarJumping Then
    If mclsObjBasics.CanFall Then

        mvartotalJump = 0
        newPos = PosBottom
        If Not OnFloor(newPos) And Form1.picTerrain.BackColor = Form1.picTerrain.Point(PosLeft + 8, PosBottom) Then
'           Si el bloque de abajo no esta ocupado, caer
            If gravity <= 0 Then gravity = mclsObjBasics.JumpVelocity / 12
            
            'If Not mvarUser Then
            '    Debug.Print PosBottom
            'End If
            
            If gravity > 0 And gravity < mclsObjBasics.JumpVelocity Then
                gravity = gravity * 2
            End If
            If Form1.picTerrain.BackColor <> Form1.picTerrain.Point(PosLeft + 8, PosBottom + 1) Then
                PosBottom = PosBottom + 1
                mvarFalling = False
            Else
                PosBottom = PosBottom + (gravity \ 1)
                mvarFalling = True
            End If
            
            'If Form1.picTerrain.BackColor <> Form1.picTerrain.Point(PosLeft + 8, PosBottom) Then
            'newPos = PosBottom
            'For i = 0 To gravity
            '    newPos = newPos + ((gravity - i) \ 1)
            '    If Form1.picTerrain.BackColor = Form1.picTerrain.Point(PosLeft + 8, newPos - 1) Then
            '        Exit For
            '    End If
            '    newPos = PosBottom
            'Next
            'PosBottom = newPos
            
            
            'If gravity > 0 And Not OnFloor(PosBottom) Then
                'For i = 0 To gravity
                '    mclsObjBasics.PosTop = mclsObjBasics.PosTop - 1
                '    If gravity = 0 Then
                '        PosBottom = newPos
                '        mvarFalling = False
                '    End If
                '
                '        PosBottom = newPos
                '        mvarFalling = False
                '        Exit For
                '    End If
                'Next
            'End If
                
            'newPos = PosBottom
            If OnFloor(newPos) And Me.User Then
                'PosBottom = newPos
                mvarFalling = False
            End If
            
            If mvarFalling Then
            If mtime > 500 Then
                mvarRunFrame = mvarRunFrame + 1
                If mvarRunFrame > mvarJumpTo Then
                    mvarRunFrame = mvarJumpFrom
                Else
                    If mvarRunFrame < mvarJumpFrom Then mvarRunFrame = mvarJumpFrom
                End If
                mtime = 0
            Else
                If (mvarRunFrame < mvarJumpFrom Or mvarRunFrame > mvarJumpTo) Then
                    mvarRunFrame = mvarJumpFrom
                End If
            End If
            End If
        Else
            'If mvarUser And mvarFalling Then
            If mvarRunFrame <> Me.ChangeFrom Then
                If (mvarRunFrame < mvarAnimFrom Or mvarRunFrame > mvarAnimTo) And _
                (mvarRunFrame < mMinRunframe Or mvarRunFrame > mMaxRunframe) Then
                    mvarRunFrame = mvarAnimFrom
                End If
            End If
'           Si el bloque de abajo esta ocupado, parase sobre el
            mvarJumping = False
            mvarFalling = False
            'Debug.Print mvarUser
            If newPos = 0 Then
                If OnFloor(newPos) Then
                    mclsObjBasics.PosBottom = (((mclsObjBasics.PosBottom \ 16)) * 16)
                Else
                    Do While Form1.picTerrain.BackColor <> Form1.picTerrain.Point(PosLeft + 8, PosBottom - 1)
                        mclsObjBasics.PosBottom = mclsObjBasics.PosBottom - 1
                        If mclsObjBasics.PosBottom < 0 Then Exit Do
                    Loop
                End If
            Else
                mclsObjBasics.PosBottom = newPos
            End If
            'mclsObjBasics.PosTop = (mclsObjBasics.PosBottom \ 16) * 16 - (mclsObjBasics.SourceBottom - mclsObjBasics.SourceTop) + 1
            fallcnt = 0
        End If
    End If
    'jumpcnt = 0
    'mvartotalJump = 0
' Hacer animacion de salto
Else
    fallcnt = 0
    If Not IsCeiling Then
        If mvartotalJump < mclsObjBasics.JumpSize Or jumpcnt Then
         
            If (mclsObjBasics.Jump = True) Or _
               (mclsObjBasics.Jump = False And mvarJumping = True) Then
                
                If Not mvarJumping And Not mvarFalling Or gravity > 0 Then
                    JumpStart = True
                    gravity = -mclsObjBasics.JumpVelocity / 8
                    endjump = mclsObjBasics.PosTop - mclsObjBasics.JumpSize
                End If
                mvarJumping = True 'Para que siga entrando por este IF
                
                'If Me.User Then
                '    Debug.Print ""
                'End If
                    
                If mclsObjBasics.PosTop <= endjump + mclsObjBasics.JumpVelocity * 2 And gravity < 0 Then
                    gravity = Int(gravity / 2)
                Else
                    If gravity > -mclsObjBasics.JumpVelocity Then
                        gravity = gravity * 2
                    End If
                End If
                
                mvarJumping = True
                If mclsObjBasics.PosTop < endjump Then
                    mclsObjBasics.PosTop = endjump
                    jumpcnt = 0
                    mclsObjBasics.Jump = False
                    mvarJumping = False
                    mvarFalling = True
                Else
                    mclsObjBasics.PosTop = mclsObjBasics.PosTop + gravity
                End If
                'mvartotalJump = mvartotalJump + mclsObjBasics.JumpVelocity - jumpcnt
                'If mvartotalJump <= 0 Then mvartotalJump = 1
            
                'Me.MoveRight 1
                If mvarRunFrame < mvarJumpFrom Then mvarRunFrame = mvarJumpFrom
                If mtime > 500 Then
                    mvarRunFrame = mvarRunFrame + 1
                    If mvarRunFrame > mvarJumpTo Then
                        mvarRunFrame = mvarJumpFrom
                    End If
                    mtime = 0
                End If
                'If jumpcnt + 8 > mclsObjBasics.JumpSize Then
                '    mclsObjBasics.Jump = False
                'End If
            Else
                'If (mvarJumping = True And jumpcnt < 16) Or _
                '   (mclsObjBasics.Jump = True And jumpcnt < mclsObjBasics.JumpSize) Then
                '    jumpcnt = jumpcnt + 1
                '    mclsObjBasics.PosTop = mclsObjBasics.PosTop - 1 'mclsObjBasics.JumpVelocity
                'Else
                '    mvarJumping = False
                '    mvarFalling = True
                'End If
                mclsObjBasics.Jump = False
                mvarJumping = False
                mvarFalling = True
                jumpcnt = 0
                mvartotalJump = 0
            End If
        Else
            mvarJumping = False 'Para que no entre mas por este IF
            mvarFalling = True
            mclsObjBasics.Jump = False
            jumpcnt = 0
            mvartotalJump = 0
        End If
    Else
        mvarJumping = False
        mclsObjBasics.Jump = False
        mvarFalling = True
        jumpcnt = 0
        mvartotalJump = 0
    End If
End If

Dim rct As RECT, rctBot As RECT
If mvarCanHit Then
For Each bot In Col
If bot.InScreen And Not bot.Died And Not bot.Breaking And Not (mvarID = bot.ID) Then
If (mvarUser And Not bot.User) Or (Not mvarUser And bot.User) Or CanHitEnemies Or bot.CanHitEnemies Then

If bot.Carried And Form1.mKeySpeed = 0 And mvarUser Then
    If direction Then
        bot.PosLeft = PosRight + 1
    Else
        bot.PosLeft = PosLeft - (bot.PosRight - bot.PosLeft) - 1
    End If
    bot.Carried = False
    bot.direction = Me.direction
    bot.WasHitted ("ByTop")
End If
If mvarUser And Not bot.InitedAI And bot.Carried Then
    If direction Then
        bot.PosLeft = PosRight - 4
    Else
        bot.PosLeft = PosLeft - (bot.PosRight - bot.PosLeft) + 4
    End If
    bot.PosTop = PosTop + (PosBottom - PosTop) / 2 - (bot.PosBottom - bot.PosTop) / 2
End If

rct.Left = PosLeft: rct.Right = PosRight
rct.Top = PosTop: rct.Bottom = PosBottom
rctBot.Left = bot.PosLeft: rctBot.Right = bot.PosRight
rctBot.Top = bot.PosTop: rctBot.Bottom = bot.PosBottom
If Crouched Then
    rct.Top = (rct.Top + rct.Bottom) / 2
End If
If bot.Crouched Then
    rctBot.Top = (rctBot.Top + rctBot.Bottom) / 2
End If
If Not bot.Hidden Then
'And Not bot.Solid And Not bot.Fixed Then
    If bot.HittedByRight Or bot.HittedByLeft Or (Me.IsFireBall And bot.FireBallCanHit) Then
        If rct.Bottom = rctBot.Bottom Then
        ' Este personaje golpeo a otro de Izquierda a Derecha
            If (rct.Right >= rctBot.Left And rct.Left < rctBot.Left) Then
                If LenB(bot.ChangePlayerTo) = 0 Or mvarUser Then
                    'bot.direction = CBool(rct.left < rctBot.left)
' Golpear si: 1. Un jugador choca a un caparazon (enemigo en modo hibernando)
'             2. El objeto chocado es un jugador
'             3. Mientras que se este moviendo y pueda chocar enemigos
                    If ((mvarUser And bot.Hibernating) Or _
                        bot.User Or _
                       (mvarCanHitEnemies And InitedAI And Not mvarHibernating) Or _
                        LenB(bot.ChangePlayerTo) Or _
                        mvarUser Or bot.MakeGrow) _
                        And Not (bot.User = True And Me.CannotHitUser) _
                        And Not (Me.User = True And bot.CannotHitUser) _
                        Then
                        'If bot.MakeGrow And LenB(GrowTo) Then
                        '    Form1.SetPlayer GrowTo
                        '    bot.MakeGrow = False
                        'End If
                        'Debug.Print bot.User
                        If Not bot.InitedAI And mvarUser And Form1.mKeySpeed <> 0 Then
                            bot.Carried = True
                        ElseIf Not InitedAI And bot.User And Form1.mKeySpeed <> 0 Then
                            Me.Carried = True
                        Else
                            If Me.CanHitEnemies Then
                                bot.WasHitted "ByLeft", True
                            Else
                                bot.WasHitted "ByLeft"
                            End If
    
                            If bot.MakeJumpWhenHitted And mvarUser Then
                                ForceJump = True
                                Jump = True
                            End If
                            If DieWhenHits Then
                                mvarDie = True
                            End If
                        End If
                    End If
                End If
            End If
        ' Este personaje golpeo a otro de Derecha a Izquierda
            If (rct.Left <= rctBot.Right And rct.Right > rctBot.Right) Then
            'If rct.right > rctBot.left And rct.left < rctBot.left Then
                If LenB(bot.ChangePlayerTo) = 0 Or mvarUser Then
                    'bot.direction = CBool(rct.left < rctBot.left)
    ' Golpear si: 1. Un jugador choca a un caparazon (enemigo en modo hibernando)
    '             2. El objeto chocado es un jugador
    '             3. Mientras que se este moviendo y pueda chocar enemigos
                    If ((mvarUser And bot.Hibernating) Or _
                        bot.User Or _
                       (mvarCanHitEnemies And InitedAI And Not mvarHibernating) Or _
                        LenB(bot.ChangePlayerTo) Or _
                        mvarUser Or bot.MakeGrow) _
                        And Not (bot.User = True And Me.CannotHitUser) _
                        And Not (Me.User = True And bot.CannotHitUser) _
                        Then
                        'If bot.MakeGrow Then
                        '    Form1.SetPlayer GrowTo
                        'End If
                        If Not bot.InitedAI And mvarUser And Form1.mKeySpeed <> 0 Then
                            bot.Carried = True
                        ElseIf Not InitedAI And bot.User And Form1.mKeySpeed <> 0 Then
                            Me.Carried = True
                        Else
                            If Me.CanHitEnemies Then
                                bot.WasHitted "ByRight", True
                            Else
                                bot.WasHitted "ByRight"
                            End If

                            If bot.MakeJumpWhenHitted And mvarUser Then
                                ForceJump = True
                                Jump = True
                            End If
                            If DieWhenHits Then
                                mvarDie = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    

    If ((bot.HittedByTop Or bot.HittedByBottom) And Not bot.Carried) Or _
        (Me.IsFireBall And bot.FireBallCanHit) Then ' Si el bot puede ser pisado
         If (rct.Left >= rctBot.Left And rct.Left <= rctBot.Right) Or _
            (rct.Right >= rctBot.Left And rct.Right <= rctBot.Right) Or _
            (rct.Left <= rctBot.Left And rct.Right >= rctBot.Right) Or _
            (rct.Left >= rctBot.Left And rct.Right <= rctBot.Right) Then
'        If (Not (rct.left < rctBot.left And rct.right <= rctBot.left) And _
'            Not (rct.left >= rctBot.right And rct.right > rctBot.right)) Then
   ' Este personaje golpeo a otro de abajo hacia arriba
            'If mvarJumping Or Jump Then
                If bot.HittedByBottom Or (Me.IsFireBall And bot.FireBallCanHit) Then
                     If rct.Top <= rctBot.Bottom And rct.Top > rctBot.Top Then
                        If (LenB(bot.ChangePlayerTo) = 0 Or _
                            mvarUser Or bot.MakeGrow) _
                            And Not (bot.User = True And Me.CannotHitUser) _
                            And Not (Me.User = True And bot.CannotHitUser) _
                            Then
                            If bot.MakeJumpWhenHitted Then
                                PosTop = rctBot.Bottom + JumpVelocity
                            End If
                            'If bot.MakeGrow Then
                            '    Form1.SetPlayer GrowTo
                            'End If
                            If bot.CanBeBreaked Then
                                If mvarCanBreak Then
                                    bot.WasHitted "ByBottom"
                                    bot.Breaking = True
                                End If
                            Else
                                If Me.CanHitEnemies Then
                                    bot.WasHitted "ByBottom", True
                                Else
                                    bot.WasHitted "ByBottom"
                                End If
                            End If
                            'If bot.NextLevel Then
                            '    JumpNextLevel = True
                            'End If
                            If bot.MakeJumpWhenHitted And mvarUser Then
                                ForceJump = True
                                Jump = True
                            End If
                            If DieWhenHits Then
                                mvarDie = True
                            End If
                        End If
                    End If
                End If
            'Else
    ' Este personaje golpeo a otro Pisandolo
                If bot.HittedByTop Or (Me.IsFireBall And bot.FireBallCanHit) Then
                    If rct.Bottom <= rctBot.Bottom And rct.Bottom >= rctBot.Top Then
'                    If rct.bottom > rctBot.top And rct.bottom < rctBot.bottom And _
'                       (rct.left >= rctBot.left And rct.left <= rctBot.right) Or _
'                       (rct.right >= rctBot.left And rct.right <= rctBot.right) Or _
'                       (rct.left <= rctBot.left And rct.right >= rctBot.right) Or _
'                       (rct.left >= rctBot.left And rct.right <= rctBot.right) Then
'                        (Not (rct.left < rctBot.left And rct.right <= rctBot.left) And _
'                         Not (rct.left >= rctBot.right And rct.right > rctBot.right)) Then
                        If ((LenB(bot.ChangePlayerTo) = 0 Or _
                            mvarUser Or bot.MakeGrow)) _
                            And Not (bot.User = True And Me.CannotHitUser) _
                            And Not (Me.User = True And bot.CannotHitUser) _
                            Then
                            'If bot.MakeGrow Then
                            '    Form1.SetPlayer GrowTo
                            'End If
                            'bot.direction = CBool(rct.left < rctBot.left)
                            If (mvarUser And bot.Hibernating) Then
                                If rct.Left < rctBot.Left Then
                                    'bot.direction = True
                                Else
                                    'bot.direction = False
                                End If
                            End If
                            
                            If Me.CanHitEnemies Then
                                bot.WasHitted "ByTop", True
                            Else
                                bot.WasHitted "ByTop"
                            End If
                            'If bot.NextLevel Then
                            '    JumpNextLevel = True
                            'End If
                            If bot.MakeJumpWhenHitted And mvarUser Then
                                ForceJump = True
                            End If
                            If DieWhenHits Then
                                mvarDie = True
                            End If
                        End If
                    End If
                End If
            'End If
        End If
    End If
End If
End If
End If
Next
End If
End If

If mvarFalling Or mvarJumping Or mclsObjBasics.Jump Then
    'If mtime > 1000 Then
    '    If mvarRunFrame + 1 > mvarJumpTo Then
    '        mvarRunFrame = mvarJumpFrom
    '    ElseIf mvarRunFrame < mvarJumpFrom Then
    '        mvarRunFrame = mvarJumpFrom
    '    Else
    '        'mvarRunFrame = mvarRunFrame + 1
    '        mvarRunFrame = mvarJumpFrom
    '    End If
    '    mtime = 0
    'End If
ElseIf Not mvarMoving Then
    If mtime > 1000 Then
        If mvarRunFrame + 1 > mvarAnimTo Then
            mvarRunFrame = mvarAnimFrom
        ElseIf mvarRunFrame < mvarAnimFrom Then
            mvarRunFrame = mvarAnimFrom
        Else
            mvarRunFrame = mvarRunFrame + 1
        End If
        mtime = 0
    End If
ElseIf mvarMoving And Not mvarUser Then
    If mtime > 1000 Then
        If mvarRunFrame + 1 > mMaxRunframe Then
            mvarRunFrame = mMinRunframe
        ElseIf mvarRunFrame < mMinRunframe Then
            mvarRunFrame = mMinRunframe
        Else
            mvarRunFrame = mvarRunFrame + 1
        End If
        mtime = 0
    End If
End If
End Sub

Public Sub WasHitted(str1 As String, Optional Destroy As Boolean = False)
    'If Me.User Then
    '    Debug.Print Me.CharType
    'End If
    If Me.Amnisty <> 0 Then Exit Sub
    If Me.User Then
        Me.Amnisty = 3000
    Else
        Me.Amnisty = 1000
    End If
    If DieWhenHits Then Exit Sub
    'Debug.Print Me.InitedAI
    If mvarInitAIWhenHitted And Destroy = False Then
        If InitedAI Then
            Me.InitedAI = False
            Me.Hibernating = True
            mvarMoving = False
        Else
            Me.InitedAI = True
            Me.Hibernating = False
        End If
    Else
        'Matar al personaje (mostrar animacion)
        mvarDie = True
        Playwav mvarDieSnd
    End If
End Sub

Public Sub Fire()
    Playwav FireSnd
    CreatePlace = "BySide"
    CreateChar , FireBall
End Sub

Public Sub CreateChar(Optional ToChar As String, Optional CreateObj As String)
Dim Enemy As clsObject
    If LenB(CreateObj) = 0 Then
        CreateObj = CreateWhenHitted
    End If
    'Si el personaje crea a otro, por ej una caja ? crear el objeto nuevo y posicionarlo
    If MakeGrow Then
        Form1.SetPlayer ToChar
    ElseIf LenB(CreateObj) Then
        Set Enemy = New clsObject
        For Each vObjProp In colObjProp
            If vObjProp.CharType = CreateObj Then Exit For
        Next
        If Not vObjProp Is Nothing Then
        MoveProp vObjProp, Enemy
        If CreatePlace = "OnTop" Then
            Enemy.PosTop = PosTop - (Enemy.SourceBottom - Enemy.SourceTop)
            Enemy.PosBottom = Enemy.PosTop + (Enemy.SourceBottom - Enemy.SourceTop)
            Enemy.PosLeft = PosLeft - Abs(Enemy.SourceRight - Enemy.SourceLeft) \ 2 + (PosRight - PosLeft) \ 2
            Enemy.PosRight = Enemy.PosLeft + Abs(Enemy.SourceRight - Enemy.SourceLeft)
        ElseIf CreatePlace = "OnPlace" Then
            Enemy.PosTop = PosTop
            Enemy.PosBottom = PosBottom
            Enemy.PosLeft = PosLeft
            Enemy.PosRight = PosRight
        ElseIf CreatePlace = "BySide" Then
            Enemy.PosTop = PosTop
            Enemy.PosBottom = PosBottom
            Enemy.PosLeft = PosLeft + CLng(direction) * -1 * Abs(PosRight - PosLeft) + CLng(Not direction) * Abs(Enemy.SourceRight - Enemy.SourceLeft)
            'Enemy.PosRight = PosRight + Direction * Abs(Enemy.PosRight - Enemy.PosLeft) + Direction * -1 * Abs(PosRight - PosLeft)
            Enemy.direction = direction
        End If
        Enemy.ID = Col.Count
        Enemy.StopMotion
        If CreateObj = FireBall Then
            Enemy.IsFireBall = True
            Enemy.ForceJump = True
        End If
        KeyIdx = KeyIdx + 1
        Enemy.Key = Enemy.CharType & KeyIdx
        Col.Add Enemy, Enemy.Key
        CreateWhenHitted = ""
        Set vObjProp = Nothing
        End If
    ElseIf LenB(ChangePlayerTo) Then
        Form1.SetPlayer ChangePlayerTo
    End If
End Sub

Public Function MoveRight(time As Long, Optional KeyPressed As Boolean = True, Optional vel As Long) As Boolean
Dim inc As Long, newPosLeft As Long, vVar As Long, i As Long
Static mtime As Long
    
    
    If mtime > 30000 Then mtime = 0
    mtime = mtime + time

    If KeyPressed Then
        If vel <> 0 Then
            newPosLeft = mclsObjBasics.PosLeft + vel
        Else
            newPosLeft = mclsObjBasics.PosLeft + mclsObjBasics.Velocity
        End If
        If mvarUser Then
            If mvarInercia < 32 Then
                'mvarInercia = mvarInercia + 2
            End If
        End If
        mvarMoving = True
    Else
        If mvarJumping Or mvarFalling Then
            If vel <> 0 Then
                newPosLeft = mclsObjBasics.PosLeft + vel
            Else
                newPosLeft = mclsObjBasics.PosLeft + mclsObjBasics.Velocity
            End If
            'mvarInercia = mvarInercia - 1
        Else
            If vel <> 0 Then
                newPosLeft = mclsObjBasics.PosLeft + vel
            Else
                newPosLeft = mclsObjBasics.PosLeft + mclsObjBasics.Velocity
            End If
            'mvarInercia = mvarInercia - 2
        End If
        If mvarInercia <= 0 Then
            mvarInercia = 0
            mvarMoving = False
        Else
            mvarMoving = True
        End If
    End If
    
    vVar = IsRightLocked(newPosLeft)
    If vVar = 0 Then
        If Not mclsObjBasics.Jump And Not mvarJumping Then
            If OnFloor(mclsObjBasics.PosLeft + 8) Or _
            Form1.picTerrain.BackColor <> Form1.picTerrain.Point(PosLeft + 8, PosBottom) Then
            For i = 1 To 6
                PosBottom = PosBottom - 1
                vVar = IsRightLocked(newPosLeft)
                If vVar <> 0 Then Exit For
            Next
            If vVar = 0 Then
                PosBottom = PosBottom + 6
                newPosLeft = 0
            Else
                'Debug.Print ""
            End If
            End If
        End If
        newPosLeft = vVar
    Else
        vVar = newPosLeft
    End If
    
    If newPosLeft <> 0 Then
        mclsObjBasics.PosLeft = newPosLeft
        mclsObjBasics.PosRight = newPosLeft + mclsObjBasics.cWidth
        Call mclsObjBasics.MoveRight(vel)
        MoveRight = True
    Else
        mvarInercia = 0
        mvarMoving = False
    End If
    
    If Not mvarJumping And Not mvarFalling And mvarMoving Then
        If mtime > 500 Or (mtime > 300 And Form1.mKeySpeed <> 0) Then
            If mvarRunFrame + 1 > mMaxRunframe Then
                mvarRunFrame = mMinRunframe
            Else
                mvarRunFrame = mvarRunFrame + 1
            End If
            mtime = 0
        End If
    ElseIf Not mvarJumping And Not mvarFalling And Not mvarMoving Then
        If mtime > 500 Or (mtime > 300 And Form1.mKeySpeed <> 0) Then
            If mvarRunFrame + 1 > mvarAnimTo Then
                mvarRunFrame = mvarAnimFrom
            ElseIf mvarRunFrame < mvarAnimFrom Then
                mvarRunFrame = mvarAnimFrom
            Else
                mvarRunFrame = mvarRunFrame + 1
            End If
            mtime = 0
        End If
    End If
End Function

Public Function MoveLeft(time As Long, Optional KeyPressed As Boolean = True, Optional vel As Long) As Boolean
Dim inc As Long, newPosLeft As Long, vVar As Long, i As Long
Static mtime

    If mtime > 30000 Then mtime = 0
    mtime = mtime + time
 
    If mclsObjBasics.PosLeft > cLevel.OffsetX Or Not mvarUser Then
        If KeyPressed Then
            If vel <> 0 Then
                newPosLeft = mclsObjBasics.PosLeft + vel
            Else
                newPosLeft = mclsObjBasics.PosLeft - mclsObjBasics.Velocity
            End If
            If mvarUser Then
                If mvarInercia > -32 Then
                    'mvarInercia = mvarInercia - 2
                End If
            End If
            mvarMoving = True
        Else
            If mvarJumping Or mvarFalling Then
                If vel <> 0 Then
                    newPosLeft = mclsObjBasics.PosLeft + vel
                Else
                    newPosLeft = mclsObjBasics.PosLeft - mclsObjBasics.Velocity
                End If
                'mvarInercia = mvarInercia + 1
            Else
                If vel <> 0 Then
                    newPosLeft = mclsObjBasics.PosLeft + vel
                Else
                    newPosLeft = mclsObjBasics.PosLeft - mclsObjBasics.Velocity
                End If
                'mvarInercia = mvarInercia + 2
                If mvarInercia > 0 Then mvarInercia = 0
            End If
            If mvarInercia >= 0 Then
                mvarInercia = 0
                mvarMoving = False
            Else
                mvarMoving = True
            End If
        End If
    Else
        newPosLeft = 0
    End If
    
    vVar = IsLeftLocked(newPosLeft)
    If vVar = 0 Then
        If Not mclsObjBasics.Jump And Not mvarJumping Then
            If OnFloor(mclsObjBasics.PosLeft + 8) Or _
            Form1.picTerrain.BackColor <> Form1.picTerrain.Point(PosLeft + 8, PosBottom) Then
                For i = 1 To 6
                    PosBottom = PosBottom - 1
                    vVar = IsLeftLocked(newPosLeft)
                    If vVar <> 0 Then Exit For
                Next
                If vVar = 0 Then
                    PosBottom = PosBottom + 6
                End If
            End If
        End If
        newPosLeft = vVar
    Else
        newPosLeft = vVar
    End If
    If newPosLeft Then
        mclsObjBasics.PosLeft = newPosLeft
        mclsObjBasics.PosRight = newPosLeft + mclsObjBasics.cWidth
        Call mclsObjBasics.MoveLeft(vel)
        MoveLeft = True
    Else
        mvarInercia = 0
        mvarMoving = False
    End If
    
    If Not mvarJumping And Not mvarFalling And mvarMoving Then
        If mtime > 500 Or (mtime > 300 And Form1.mKeySpeed <> 0) Then
            If mvarRunFrame + 1 > mMaxRunframe Then
                mvarRunFrame = mMinRunframe
            Else
                mvarRunFrame = mvarRunFrame + 1
            End If
            mtime = 0
        End If
    ElseIf Not mvarJumping And Not mvarFalling And Not mvarMoving Then
        If mtime > 500 Or (mtime > 300 And Form1.mKeySpeed <> 0) Then
            If mvarRunFrame + 1 > mvarAnimTo Then
                mvarRunFrame = mvarAnimFrom
            ElseIf mvarRunFrame < mvarAnimFrom Then
                mvarRunFrame = mvarAnimFrom
            Else
                mvarRunFrame = mvarRunFrame + 1
            End If
            mtime = 0
        End If
    End If
End Function

Public Function IsLeftLocked(newPosLeft As Long) As Long
Dim bcoordX As Long, bcoordY As Long, bcoordYTop As Long
Dim charBottom As String, charTop As String
Dim coordYBottom As Long, coordYTop As Long, CoordX As Long
Dim bot As clsObject

    coordYBottom = (mclsObjBasics.PosBottom - 2) \ 16 'Chequeo el proximo bloque si esta ocupado por una X
    coordYTop = ((mclsObjBasics.PosTop) \ 16)  'Chequeo el proximo bloque si esta ocupado por una X
    CoordX = ((mclsObjBasics.CharPos + newPosLeft) \ 16) + 1
    charBottom = Form1.GetBlock(CoordX, coordYBottom + 1)
    charTop = Form1.GetBlock(CoordX, coordYTop + 1)
    If mvarUser Then
        If Len(charBottom) And InStrB(1, mvarNextLevelChar, charBottom) > 0 Then
            JumpNextLevel = True
        End If
    End If
    'Debug.Print Form1.picTerrain.Point(PosLeft - 1, PosBottom - 1)
    If Form1.picTerrain.BackColor <> Form1.picTerrain.Point(PosLeft + 8, PosBottom - 1) And _
       (Not mvarJumping And Not Jump) Then
        IsLeftLocked = 0
        Exit Function
    End If
    If LenB(charBottom) > 0 And InStrB(1, vFixed, charBottom) Then
        IsLeftLocked = 0
        Exit Function
    ElseIf LenB(charTop) > 0 And InStrB(1, vFixed, charTop) Then
        IsLeftLocked = 0
        Exit Function
    Else
        For Each bot In Col
            If bot.InScreen Then
                If bot.Solid Then
                    If Not bot.Breaking And Not (mvarID = bot.ID) Then
                        bcoordY = (bot.PosBottom - 1) \ 16 'Chequeo el proximo bloque si esta ocupado por una X
                        bcoordYTop = bcoordY + (Abs(bot.PosTop - bot.PosBottom) - 1) \ 16 'Chequeo el proximo bloque si esta ocupado por una X
                        bcoordX = (bot.PosLeft \ 16) + 1
                        If (bcoordY = coordYBottom Or bcoordY = coordYTop) And bcoordX = CoordX Or _
                           ((bcoordYTop = coordYBottom Or bcoordYTop = coordYTop) And bcoordX = CoordX) Then
                        IsLeftLocked = 0
                        Exit Function
                        End If
                    End If
                End If
            End If
        Next
    End If
    IsLeftLocked = newPosLeft
End Function

Public Function IsRightLocked(newPosLeft As Long, Optional vVel) As Long
Dim bcoordX As Long, bcoordY As Long, bcoordYTop As Long
Dim charBottom As String, charTop As String
Dim coordYBottom As Long, coordYTop As Long, CoordX As Long
Dim bot As clsObject

    coordYBottom = (mclsObjBasics.PosBottom - 2) \ 16 'Chequeo el proximo bloque si esta ocupado por una X
    coordYTop = ((mclsObjBasics.PosTop) \ 16)  'Chequeo el proximo bloque si esta ocupado por una X
    CoordX = ((mclsObjBasics.CharPos + newPosLeft + mclsObjBasics.cWidth) \ 16) + 1
    charBottom = Form1.GetBlock(CoordX, coordYBottom + 1)
    charTop = Form1.GetBlock(CoordX, coordYTop + 1)
    If mvarUser Then
        If LenB(charBottom) And InStrB(1, mvarNextLevelChar, charBottom) > 0 Then
            JumpNextLevel = True
        End If
    End If
    If Form1.picTerrain.BackColor <> Form1.picTerrain.Point(PosRight - 8, PosBottom - 1) And _
       (Not mvarJumping And Not Jump) Then
        IsRightLocked = 0
        Exit Function
    End If
    If LenB(charBottom) > 0 And InStrB(1, vFixed, charBottom) Then
        IsRightLocked = 0
        Exit Function
    ElseIf LenB(charTop) > 0 And InStrB(1, vFixed, charTop) Then
        IsRightLocked = 0
        Exit Function
    Else
        CoordX = ((mclsObjBasics.CharPos + newPosLeft + mclsObjBasics.cWidth - JumpVelocity) \ 16) + 1
        For Each bot In Col
            If bot.InScreen Then
                If bot.Solid Then
                    If Not bot.Breaking And Not (mvarID = bot.ID) Then
                        bot.CharType = bot.CharType
                        bcoordY = (bot.PosBottom - 1) \ 16 'Chequeo el proximo bloque si esta ocupado por una X
                        bcoordYTop = bcoordY + (Abs(bot.PosTop - bot.PosBottom) - 1) \ 16 'Chequeo el proximo bloque si esta ocupado por una X
                        bcoordX = (bot.PosLeft \ 16) + 1
                        If ((bcoordY = coordYBottom Or bcoordY = coordYTop) And bcoordX = CoordX) Or _
                           ((bcoordYTop = coordYBottom Or bcoordYTop = coordYTop) And bcoordX = CoordX) Then
                         IsRightLocked = 0
                        Exit Function
                        End If
                    End If
                End If
            End If
        Next
    End If
    IsRightLocked = newPosLeft
End Function

Public Function IsCeiling()
Dim CoordXL As Long, CoordXR As Long, CoordY As Long, i As Long
Dim vChar As String, j As Long, vIsBrekeable As Boolean
Dim bot As clsObject, cbBlock As clsbBlock

    CoordY = ((mclsObjBasics.PosTop - 1) \ 16) + 1 'Chequeo el proximo bloque si esta ocupado por una X
    CoordXL = ((mclsObjBasics.CharPos + mclsObjBasics.PosLeft + 2) \ 16) + 1
    CoordXR = ((mclsObjBasics.CharPos + mclsObjBasics.PosRight - 2) \ 16) + 1
    
    For j = CoordXL To CoordXR
        vChar = Form1.GetBlock(j, CoordY)
        If InStrB(1, vFixed, vChar) Then
            For i = 1 To cBreakeable.Count
                Set cbBlock = cBreakeable(i)
                If cbBlock.Visible And Not cbBlock.StartAnim And InScreen Then
                    If cbBlock.CharType = vChar Then
                        vIsBrekeable = True
                        If (CoordXL <= cbBlock.CoordX And CoordXR >= cbBlock.CoordX) And _
                            CoordY = cbBlock.CoordY Then
                            If CanBreak Then
                                Playwav cbBlock.DieSnd
                                cbBlock.StartAnim = True
                                cbBlock.SetBlock " "
                            End If
                            IsCeiling = True
                            Exit Function
                        End If
                    End If
                End If
            Next
        End If
        vChar = Form1.GetBlock(j, CoordY)
        If InStrB(1, vFixed, vChar) And Not vIsBrekeable Then
            IsCeiling = True
            Exit Function
        End If
    Next
    For Each bot In Col
        If bot.InScreen And bot.Solid And Not bot.Breaking And Not (mvarID = bot.ID) Then
            If PosTop <= bot.PosBottom And PosTop >= bot.PosTop Then
                If (PosLeft > bot.PosLeft And PosLeft < bot.PosRight) Or _
                   (PosRight > bot.PosLeft And PosRight < bot.PosRight) Or _
                   (PosLeft < bot.PosLeft And PosRight > bot.PosRight) Or _
                   (PosLeft >= bot.PosLeft And PosRight <= bot.PosRight) Then
                    PosTop = bot.PosBottom + 1
                     IsCeiling = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function OnFloor(newPos As Long) As Boolean
Dim CoordXL As Long, CoordXR As Long, CoordY As Long
Dim vChar As String
Dim bot As clsObject, i As Long

    'CoordY = (mclsObjBasics.PosBottom \ 16) + 1 'Chequeo el proximo bloque si esta ocupado por una X
    CoordY = ((mclsObjBasics.PosBottom) \ 16) + 1 'Chequeo el proximo bloque si esta ocupado por una X
    CoordXR = ((mclsObjBasics.CharPos + mclsObjBasics.PosRight - 2) \ 16) + 1
    CoordXL = ((mclsObjBasics.CharPos + mclsObjBasics.PosLeft + 2) \ 16) + 1
    
    For i = CoordXL To CoordXR
        vChar = Form1.GetBlock(i, CoordY)
        If vChar = "-1" Then
            mvarDie = True
            Exit Function
        End If
'           (LenB(vChar) > 0 And InStrB(1, vIsFloor, vChar)) Or _'
        If (LenB(vChar) > 0 And InStrB(1, vFixed, vChar)) Or _
            mvarForceJump Then
            'mclsObjBasics.PosBottom = ((mclsObjBasics.PosBottom \ 16) * 16)
            newPos = ((CoordY - 1) * 16) '+ 1
            OnFloor = True
            Exit Function
        End If
    Next
    
    For Each bot In Col
        'Debug.Print Me.User
        If bot.InScreen And bot.Solid And Not bot.Breaking And Not (mvarID = bot.ID) Then
            If PosBottom <= bot.PosBottom And PosBottom >= bot.PosTop - 1 Then
                If (PosLeft > bot.PosLeft And PosLeft < bot.PosRight) Or _
                   (PosRight > bot.PosLeft And PosRight < bot.PosRight) Or _
                   (PosLeft < bot.PosLeft And PosRight > bot.PosRight) Or _
                   (PosLeft >= bot.PosLeft And PosRight <= bot.PosRight) Then
                    OnFloor = True
                    newPos = bot.PosTop
                    Exit Function
                End If
            End If
        End If
    Next
    newPos = 0
End Function

Public Function CheckIfInScreen(pOffsetX As Long, pScreenWidth As Long)
    If PosRight > pOffsetX And PosLeft < pOffsetX + pScreenWidth Then
        CheckIfInScreen = True
        mvarInScreen = True
    Else
        mvarInScreen = False
    End If
End Function
