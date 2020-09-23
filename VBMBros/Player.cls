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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarHdc As Long 'copia local
'Private mvarPosCoords As Rect 'copia local
Private SourceCoords As Rect 'copia local
Private mvarVelocity As Long 'copia local
Private mvarDirection As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarPosLeft As Long 'copia local
Private mvarPosRight As Long 'copia local
Private mvarPosTop As Long 'copia local
Private mvarPosBottom As Long 'copia local
Private mvarJump As Boolean 'copia local
Private jumpcnt As Long
Private fallcnt As Long
Private mvarCharPos As Long
Private mvarflip As Boolean
Private RunFrame As Long
'variables locales para almacenar los valores de las propiedades
Private mvarRScreenLimit As Long 'copia local
Public Property Let RScreenLimit(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.RScreenLimit = 5
    mvarRScreenLimit = vData
End Property


Public Property Get RScreenLimit() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.RScreenLimit
    Set RScreenLimit = mvarRScreenLimit
End Property




Friend Property Let PosBottom(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.PosBottom = 5
    mvarPosBottom = vData
End Property


Friend Property Get PosBottom() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.PosBottom
    PosBottom = mvarPosBottom
End Property



Friend Property Let PosTop(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.PosTop = 5
    mvarPosBottom = vData + 32
    mvarPosTop = vData
End Property


Friend Property Get PosTop() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.PosTop
    PosTop = mvarPosTop
End Property



Friend Property Let PosRight(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.PosRight = 5
    mvarPosRight = vData
End Property


Friend Property Get PosRight() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.PosRight
    PosRight = mvarPosRight
End Property

Friend Property Let PosLeft(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.PosLeft = 5
    mvarPosLeft = vData
    mvarPosRight = mvarPosLeft + 16 '(vData - mvarPosLeft)
End Property

Friend Property Get PosLeft() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.PosLeft
    PosLeft = mvarPosLeft
End Property

Friend Property Let Direction(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Direction = 5
    mvarDirection = vData
End Property


Friend Property Get Direction() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Direction
    Direction = mvarDirection
End Property



Friend Property Let Velocity(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Velocity = 5
    mvarVelocity = vData
End Property


Friend Property Get Velocity() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Velocity
    Velocity = mvarVelocity
End Property

Friend Property Let Jump(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Hdc = 5
Dim coordX As Long, coordY As Long
    coordY = Fix(PosBottom / 16) + 1 'Chequeo el proximo bloque si esta ocupado por una X
    coordX = Fix((mvarCharPos + PosLeft + 8) / 16) + 1
    If Form1.GetBlock(coordX, coordY) = "X" Then
        mvarJump = vData
    End If
End Property

Friend Property Get Jump() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Hdc
    Jump = mvarJump
End Property


Friend Property Let hdc(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Hdc = 5
    mvarHdc = vData
End Property

Friend Property Get hdc() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Hdc
    hdc = mvarHdc
End Property

Friend Sub Animate()
Dim coordX As Long, coordY As Long, coordXL As Long, coordXR As Long
If mvarJump = False Then
    coordY = Fix(mvarPosBottom / 16) + 1 'Chequeo el proximo bloque si esta ocupado por una X
    coordXL = Fix((mvarCharPos + mvarPosLeft + 2) / 16) + 1
    coordXR = Fix((mvarCharPos + mvarPosRight - 2) / 16) + 1
    If Form1.GetBlock(coordXL, coordY) <> "X" _
   And Form1.GetBlock(coordXR, coordY) <> "X" Then
        If fallcnt < 4 Then
            fallcnt = fallcnt + 1
            PosTop = mvarPosTop + 1
        Else
            PosTop = mvarPosTop + 3
        End If
        RunFrame = 5
    Else
        PosTop = (coordY - 3) * 16
        If RunFrame = 5 Then RunFrame = 0
    End If
Else
    fallcnt = 0
    coordY = Fix(mvarPosBottom / 16)  'Chequeo el proximo bloque si esta ocupado por una X
    coordXL = Fix((mvarCharPos + mvarPosLeft + 2) / 16) + 1
    coordXR = Fix((mvarCharPos + mvarPosRight - 2) / 16) + 1
    If Form1.GetBlock(coordXL, coordY) <> "X" _
   And Form1.GetBlock(coordXR, coordY) <> "X" Then
        If jumpcnt < 64 Then
            If jumpcnt < 60 Then
                jumpcnt = jumpcnt + 3
                PosTop = mvarPosTop - 3
            Else
                jumpcnt = jumpcnt + 1
                PosTop = mvarPosTop - 1
            End If
        Else
            mvarJump = False
            jumpcnt = 0
        End If
    Else
        mvarJump = False
        jumpcnt = 0
    End If
    RunFrame = 5
End If
End Sub

Public Function MoveRight(time As Long) As Long
Dim coordX As Long, coordY As Long
Static mtime
mtime = mtime + time

    coordY = Fix(mvarPosBottom / 16)  'Chequeo el proximo bloque si esta ocupado por una X
    coordX = Fix((mvarCharPos + mvarPosRight + 2) / 16) + 1
    mvarflip = False
    
    If Form1.GetBlock(coordX, coordY) <> "X" Then
        If mvarPosLeft < mvarRScreenLimit Then
            mvarPosLeft = mvarPosLeft + 2
            mvarPosRight = mvarPosLeft + 16
        Else
            mvarCharPos = mvarCharPos + 2
            MoveRight = mvarCharPos
        End If
    End If
If mtime > 500 Then
    RunFrame = RunFrame + 1
    If RunFrame > 3 Then
        RunFrame = 1
    End If
    mtime = 0
End If
End Function

Public Function MoveLeft(time As Long) As Long
Dim coordX As Long, coordY As Long
Static mtime
mtime = mtime + time
    
    coordY = Fix(mvarPosBottom / 16)  'Chequeo el proximo bloque si esta ocupado por una X
    coordX = Fix((mvarCharPos + mvarPosLeft - 2) / 16) + 1
    mvarflip = True
    
    If Form1.GetBlock(coordX, coordY) <> "X" Then
        If mvarPosLeft > 0 Then
            mvarPosLeft = mvarPosLeft - 2
            mvarPosRight = mvarPosLeft + 16
        Else
            mvarPosLeft = 0
        End If
    End If
If mtime > 500 Then
    RunFrame = RunFrame + 1
    If RunFrame > 3 Then
        RunFrame = 1
    End If
    mtime = 0
End If
End Function

Private Sub Class_Initialize()
    hdc = Form1.picMario.hdc
    SourceCoords.left = 0
    SourceCoords.top = 0
    SourceCoords.right = 16
    SourceCoords.bottom = 32
    mvarPosLeft = 0
    mvarPosTop = 0
    mvarPosRight = 16
    mvarPosBottom = 32
End Sub

Public Sub Draw()
    rctSrc.left = RunFrame * 16 + RunFrame
    rctSrc.right = rctSrc.left + 16
    rctSrc.top = 0
    rctSrc.bottom = 32
    rctDst.left = mvarPosLeft
    rctDst.right = mvarPosRight
    rctDst.top = mvarPosTop
    rctDst.bottom = mvarPosBottom
    TransparentBlt Form1.hdc, rctDst, hdc, rctSrc, vbMagenta, mvarflip
End Sub
