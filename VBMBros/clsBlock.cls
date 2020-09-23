VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarSourceLeft As Long 'local copy
Private mvarSourceTop As Long 'local copy
Private mvarCharType As String 'local copy
Private mvarSourceRight As Long 'local copy
Private mvarSourceBottom As Long 'local copy
Private mvarNonSolid As Long 'local copy
Private mvarFixed As Long 'local copy
Private mvarJumpNextLevel As Long 'local copy
Private mvarIsFloor As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarCol As Long 'local copy
Private mvarRow As Long 'local copy
Public Property Let Row(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Row = 5
    mvarRow = vData
End Property


Public Property Get Row() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Row
    Row = mvarRow
End Property



Public Property Let Col(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Col = 5
    mvarCol = vData
End Property


Public Property Get Col() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Col
    Col = mvarCol
End Property



Public Property Let IsFloor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IsFloor = 5
    mvarIsFloor = vData
End Property


Public Property Get IsFloor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IsFloor
    IsFloor = mvarIsFloor
End Property



Public Property Let JumpNextLevel(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.JumpNextLevel = 5
    mvarJumpNextLevel = vData
End Property


Public Property Get JumpNextLevel() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.JumpNextLevel
    JumpNextLevel = mvarJumpNextLevel
End Property



Public Property Let Fixed(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Fixed = 5
    mvarFixed = vData
End Property


Public Property Get Fixed() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Fixed
    Fixed = mvarFixed
End Property



Public Property Let NonSolid(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NonSolid = 5
    mvarNonSolid = vData
End Property


Public Property Get NonSolid() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NonSolid
    NonSolid = mvarNonSolid
End Property



Public Property Let SourceBottom(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SourceBottom = 5
    mvarSourceBottom = vData
End Property


Public Property Get SourceBottom() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SourceBottom
    SourceBottom = mvarSourceBottom
End Property



Public Property Let SourceRight(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SourceRight = 5
    mvarSourceRight = vData
End Property


Public Property Get SourceRight() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SourceRight
    SourceRight = mvarSourceRight
End Property



Public Property Let CharType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CharType = 5
    mvarCharType = vData
End Property


Public Property Get CharType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CharType
    CharType = mvarCharType
End Property



Public Property Let SourceTop(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SourceTop = 5
    mvarSourceTop = vData
End Property


Public Property Get SourceTop() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SourceTop
    SourceTop = mvarSourceTop
End Property



Public Property Let SourceLeft(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SourceLeft = 5
    mvarSourceLeft = vData
End Property


Public Property Get SourceLeft() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SourceLeft
    SourceLeft = mvarSourceLeft
End Property



