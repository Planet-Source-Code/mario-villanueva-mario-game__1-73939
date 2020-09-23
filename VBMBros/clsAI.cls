VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarInited As Boolean
Private mvarAI As String
Private mvarCurrentAI As String

Public Property Let AI(ByVal vData As String)
    mvarAI = vData
End Property

Public Property Get AI() As String
    AI = mvarAI
End Property

Public Property Let Inited(ByVal vData As Boolean)
    mvarInited = vData
End Property

Public Property Get Inited() As Boolean
    Inited = mvarInited
End Property

Public Sub PerformAI(mclsObject As clsObject)
Dim CoordX As Long, CoordY As Long, CoordXL As Long, CoordXR As Long
Dim action As String, param2 As String
    If mclsObject.Stand > 0 Then
        mclsObject.Stand = mclsObject.Stand - 1
        Exit Sub
    End If
    If mclsObject.Raising Or mclsObject.Descending Then Exit Sub
    If LenB(mvarCurrentAI) = 0 Then
        If LenB(mvarAI) = 0 Then Exit Sub
        mvarCurrentAI = mvarAI
    End If
    If InStrB(1, mvarCurrentAI, ",") Then
        action = Split(mvarCurrentAI, ",")(0)
    Else
        action = mvarCurrentAI
    End If
    If InStrB(1, action, "=") Then
        param2 = Split(action, "=")(1)
        action = Split(action, "=")(0)
    Else
        action = action
    End If
    Select Case LCase(action)
    Case "stand"
        mclsObject.Stand = Val(param2)
    Case "raise"
        mclsObject.Raising = True
    Case "canhit"
        mclsObject.CanHit = True
    Case "cannothit"
        mclsObject.CanHit = False
    Case "descend"
        mclsObject.Descending = True
    Case "die"
        If mclsObject.Jump = False And _
           mclsObject.Jumping = False And _
           mclsObject.Falling = False Then
            mclsObject.WasHitted "ByTop"
        Else
            mvarAI = "die"
        End If
    Case "jump"
        mclsObject.ForceJump = True
        mclsObject.Jump = True
    Case "ascending"
        If mclsObject.PosTop <= -(mclsObject.PosBottom - mclsObject.PosTop) Then
            mclsObject.PosTop = cLevel.VBlocks * 16
        Else
            mclsObject.PosTop = mclsObject.PosTop - 1
        End If
    Case "descending"
        If mclsObject.PosTop >= cLevel.VBlocks * 16 Then
            mclsObject.PosTop = -(mclsObject.PosBottom - mclsObject.PosTop)
        Else
            mclsObject.PosTop = mclsObject.PosTop + 1
        End If
    Case "horizontaljumpmove"
        If mclsObject.OnFloor(0) Or Form1.picTerrain.BackColor <> Form1.picTerrain.Point(mclsObject.PosLeft + 8, mclsObject.PosBottom) Then
            mclsObject.Jump = True
            mclsObject.ForceJump = False
        End If
        If mclsObject.PosLeft <= 0 Then
            mclsObject.direction = True
        Else
            If mclsObject.direction Then
                If Not mclsObject.MoveRight(50) Then
                    mclsObject.MoveLeft (50)
                End If
            Else
                If Not mclsObject.MoveLeft(50) Then
                    mclsObject.MoveRight (50)
                End If
            End If
        End If
    Case "horizontalmove"
        'If mclsObject.PosLeft <= 0 Then
        '    mclsObject.direction = True
        'Else
            If mclsObject.direction Then
                If Not mclsObject.MoveRight(50) Then
                    mclsObject.MoveLeft (50)
                End If
            Else
                If Not mclsObject.MoveLeft(50) Then
                    mclsObject.MoveRight (50)
                End If
            End If
       ' End If
    End Select
    mvarCurrentAI = Mid$(mvarCurrentAI, Len(action) + 2)
End Sub
