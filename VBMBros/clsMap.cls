VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarOffsetX As Long 'copia local
Private mvarOffsetY As Long 'copia local
Private mvarLoadMap As String 'copia local
Private ArrLines() As String
Private ArrLines2() As String
Private mvarDef As String 'local copy
Private cMapObj As Collection
Private vMapObjProp As clsObjProp
Private vPlayersDef As String

Public Property Let PlayersDef(ByVal vData As String)
    vPlayersDef = vData
End Property

Public Property Get PlayersDef() As String
    PlayersDef = vPlayersDef
End Property

Public Property Let Def(ByVal vData As String)
    mvarDef = vData
End Property


Public Property Get Def() As String
    Def = mvarDef
End Property

Public Function SetBlock(PosX As Long, PosY As Long, pChar As String) As String
    ArrLines(PosY) = left$(ArrLines(PosY), PosX - 1) + " " + Mid$(ArrLines(PosY), PosX + 1)
    SetBlock = Mid$(ArrLines(PosY), PosX, 1)
    If LenB(SetBlock) = 0 Then
        SetBlock = " "
        Exit Function
    End If
End Function

Public Function Block(PosX As Long, PosY As Long) As String
    Block = Mid$(ArrLines(PosY), PosX, 1)
    If LenB(Block) = 0 Then
        Block = ""
        Exit Function
    End If
End Function

Public Function PBlock(PosX As Long, PosY As Long) As String
    PBlock = Mid$(ArrLines2(PosY), PosX, 1)
    If LenB(PBlock) = 0 Then
        PBlock = ""
        Exit Function
    End If
End Function

Public Property Get HBlocks() As String
    HBlocks = Len(ArrLines(0))
End Property

Public Property Let HBlocks(ByVal vData As String)
Dim i As Long
    For i = 1 To UBound(ArrLines)
        ArrLines(i) = Space$(vData)
    Next
End Property

Public Property Get VBlocks() As String
    VBlocks = UBound(ArrLines)
End Property

Public Property Let VBlocks(ByVal vData As String)
    ReDim ArrLines(vData)
End Property

Public Function LoadMap() As String
Dim str1 As String, i As Long, vVar As String
Dim arrObjects() As String, vObjProp As clsObjProp

    Set cMapObj = New Collection
' Load Tiles into the PictureBox
    Form1.picTiles.Picture = LoadPicture(Replace$(GetFromIni("World", "Tiles", Def), "AppPath", App.Path))
    vGameMusic = ""
    vGameMusic = Replace$(GetFromIni("World", "Music", Def), "AppPath", App.Path)
' -------------------------------------------------------------------
' Reads Tile descriptions
' -------------------------------------------------------------------
    vVar = GetFromIni("World", "Objects", Def)
    
    vPlayersDef = GetFromIni("World", "PlayersDef", Def)
    vPlayersDef = Replace$(vPlayersDef, "AppPath", App.Path)
    
    vGifts = GetFromIni("Game", "Gifts", PlayersDef)
    arrGifts = Split(vGifts, ",")
    
    For Each vObjProp In cMapObj
        Set vObjProp = Nothing
    Next
    vIsFloor = ""
    vFixed = ""
    Set cMapObj = Nothing
    Set cMapObj = New Collection
    
    arrObjects = Split(vVar, ",")
    For i = 0 To UBound(arrObjects)
        Call InitMapObjects(arrObjects(i), Def)
    Next
    
    Erase arrObjects
' -------------------------------------------------------------------
' Read Map Data
' -------------------------------------------------------------------
    str1 = ReadBFile(Replace$(GetFromIni("World", "Map", Def), "AppPath", App.Path))
    ArrLines = Split(str1, Chr(13))
    For i = 1 To UBound(ArrLines)
        ArrLines(i) = Replace$(ArrLines(i), Chr(13), "")
        ArrLines(i) = Replace$(ArrLines(i), Chr(10), "")
    Next
    CreateMap
    'Debug.Print GetFromIni("World", "MapPlayers", Def)
    str1 = ReadBFile(Replace$(GetFromIni("World", "MapPlayers", Def), "AppPath", App.Path))
    ArrLines2 = Split(str1, Chr(13))
    For i = 1 To UBound(ArrLines2)
        ArrLines2(i) = Replace$(ArrLines2(i), Chr(13), "")
        ArrLines2(i) = Replace$(ArrLines2(i), Chr(10), "")
    Next

' -------------------------------------------------------------------
' Remove Tile descriptions
' -------------------------------------------------------------------
    For i = cMapObj.Count To 1 Step -1
        Set vObjProp = cMapObj(i)
        If vObjProp.JumpNextLevel Then
            LoadMap = vObjProp.CharType
        End If
        Set vObjProp = Nothing
        cMapObj.Remove (i)
    Next
    Set cMapObj = Nothing
End Function


Public Property Let OffsetY(ByVal vData As Long)
    mvarOffsetY = vData
End Property

Public Property Get OffsetY() As Long
    OffsetY = mvarOffsetY
End Property

Public Property Let OffsetX(ByVal vData As Long)
    mvarOffsetX = vData
End Property


Public Property Get OffsetX() As Long
    OffsetX = mvarOffsetX
End Property

Public Sub Redraw(srleft As Long, srtop As Long, srright As Long, srbottom As Long, _
                  dtleft As Long, dttop As Long, dtright As Long, dtbottom As Long)
                  
    'rctSrc.left = srleft * m_mult - mvarOffsetX
    'rctSrc.top = srtop * m_mult
    'rctSrc.right = srright * m_mult - mvarOffsetX
    'rctSrc.bottom = srbottom * m_mult
    
    'rctDst.left = dtleft * m_mult
    'rctDst.top = dttop * m_mult
    'rctDst.right = dtright * m_multp
    'rctDst.bottom = dtbottom * m_mult
    
    'Paint Form1.hdc, rctDst, Form1.picBack.hdc, rctSrc, vbSrcCopy
    
    rctSrc.left = srleft * m_mult
    rctSrc.top = srtop * m_mult
    rctSrc.right = srright * m_mult
    rctSrc.bottom = srbottom * m_mult
    
    rctDst.left = dtleft * m_mult
    rctDst.top = dttop * m_mult
    rctDst.right = dtright * m_mult
    rctDst.bottom = dtbottom * m_mult

    'Paint Form1.hdc, rctDst, Form1.PicMapMask.hdc, rctSrc, vbMergePaint
    Paint Form1.pForm1.hdc, rctDst, Form1.PicMap.hdc, rctSrc
End Sub

Public Sub CreateMap()
Dim i As Long, j As Long, vVal As String
Dim vObjBreak As clsbBlock, y As Long
Form1.PicMap.Width = HBlocks * 16 * m_mult + 8
Form1.PicMap.Height = VBlocks * 16 * m_mult + 8
Form1.picTerrain.Width = HBlocks * 16 * m_mult + 8
Form1.picTerrain.Height = VBlocks * 16 * m_mult + 8
Form1.PicMapMask.Width = HBlocks * 16 * m_mult + 8
Form1.PicMapMask.Height = VBlocks * 16 * m_mult + 8

'On Error GoTo errFounded
If IsInCol(cMapObj, "Space") <> "" Then
    For Each vObjProp In cMapObj
        If vObjProp.CharType = "Space" Then Exit For
    Next
    rctSrc.left = vObjProp.SourceLeft
    rctSrc.right = vObjProp.SourceRight
    rctSrc.top = vObjProp.SourceTop
    rctSrc.bottom = vObjProp.SourceBottom
    rctDst.left = 0
    rctDst.right = Form1.PicMap.ScaleWidth
    rctDst.top = 0
    rctDst.bottom = Form1.PicMap.ScaleHeight
    StretchPaint Form1.PicMap.hdc, rctDst, Form1.picTiles.hdc, rctSrc
    StretchPaint Form1.picTerrain.hdc, rctDst, Form1.picTiles.hdc, rctSrc
    Form1.picTerrain.BackColor = Form1.picTerrain.Point(0, 0)
    Form1.BackColor = Form1.picTerrain.Point(0, 0)
End If
'On Error GoTo 0

    For i = 1 To VBlocks
        For j = 1 To HBlocks
            vVal = Replace$(Block(j, i), " ", "Space")
                            If vSwitchGifts Then
                                For y = 0 To UBound(arrGifts)
                                    If arrGifts(y) = vVal Then
                                        Call Randomize(Timer)
                                        vVal = CInt(Rnd * UBound(arrGifts))
                                        On Error Resume Next
                                        Mid(ArrLines(i), j, 1) = arrGifts(vVal)
                                        Exit For
                                    End If
                                Next
                            End If
            Select Case vVal
                Case Is = IsInCol(cMapObj, vVal)
                    For Each vObjProp In cMapObj
                        If vObjProp.CharType = vVal Then Exit For
                    Next
                    rctSrc.left = vObjProp.SourceLeft
                    rctSrc.right = vObjProp.SourceRight
                    rctSrc.top = vObjProp.SourceTop
                    rctSrc.bottom = vObjProp.SourceBottom
                Case Else
                    For Each vObjProp In cMapObj
                        If vObjProp.CharType = vVal Then Exit For
                    Next
                    If Not vObjProp Is Nothing Then
                        rctSrc.left = vObjProp.SourceLeft
                        rctSrc.right = vObjProp.SourceRight
                        rctSrc.top = vObjProp.SourceTop
                        rctSrc.bottom = vObjProp.SourceBottom
                    End If
            End Select
            
            If Not vObjProp Is Nothing Then
            rctDst.left = ((j - 1) * 16)
            rctDst.right = (j * 16)
            rctDst.top = (i - 1) * 16
            rctDst.bottom = i * 16
            TransparentBlt Form1.PicMap.hdc, rctDst, Form1.picTiles.hdc, rctSrc, vbMagenta

            For Each vObjProp In cMapObj
                If vObjProp.CharType = vVal Then Exit For
            Next
            
            If vObjProp.IsFloor Then
                TransparentBlt Form1.picTerrain.hdc, rctDst, Form1.picTiles.hdc, rctSrc, vbMagenta
            End If
            
            If vObjProp.NonSolid Then
                Set vObjBreak = New clsbBlock
                vObjBreak.hdc = Form1.picTiles.hdc
                vObjBreak.CharType = vObjProp.CharType
                vObjBreak.SourceLeft = vObjProp.SourceLeft: vObjBreak.SourceRight = vObjProp.SourceRight
                vObjBreak.SourceTop = vObjProp.SourceTop: vObjBreak.SourceBottom = vObjProp.SourceBottom
                vObjBreak.PosLeft = (j - 1) * 16: vObjBreak.PosRight = j * 16
                vObjBreak.PosTop = (i - 1) * 16: vObjBreak.PosBottom = (i) * 16
                vObjBreak.CoordX = j: vObjBreak.CoordY = i
                vObjBreak.DieSnd = vObjProp.DieSnd
                vObjBreak.Visible = True
                vBreakeableIdx = vBreakeableIdx + 1
                vObjBreak.KeyIndex = vBreakeableIdx
                cBreakeable.Add vObjBreak, Str(vBreakeableIdx)
                'On Error GoTo errFounded
                For Each vObjProp In cMapObj
                    If vObjProp.CharType = "Space" Then Exit For
                Next
                'On Error GoTo 0
                rctSrc.left = vObjProp.SourceLeft
                rctSrc.right = vObjProp.SourceRight
                rctSrc.top = vObjProp.SourceTop
                rctSrc.bottom = vObjProp.SourceBottom
                rctDst.left = ((j - 1) * 16)
                rctDst.right = (j * 16)
                rctDst.top = (i - 1) * 16
                rctDst.bottom = i * 16
                TransparentBlt Form1.PicMapMask.hdc, rctDst, Form1.picTiles.hdc, rctSrc, vbMagenta
            End If
            End If
        Next
    Next
    Set vObjProp = Nothing
Exit Sub
errFounded:
    MsgBox "Tile not founded!"
    Set vObjProp = Nothing
End Sub

Public Sub InitMapObjects(vObj As String, vFilePath As String)

    Set vObjProp = New clsObjProp

    vObjProp.CharType = Replace$(GetFromIni(vObj, "CharType", vFilePath), " ", "Space")
    vObjProp.JumpNextLevel = Val(GetFromIni(vObj, "JumpNextLevel", vFilePath))
    vObjProp.SourceLeft = Val(GetFromIni(vObj, "SourceLeft", vFilePath))
    vObjProp.SourceRight = Val(Val(GetFromIni(vObj, "SourceRight", vFilePath)))
    vObjProp.SourceTop = Val(GetFromIni(vObj, "SourceTop", vFilePath))
    vObjProp.SourceBottom = Val(GetFromIni(vObj, "SourceBottom", vFilePath))
    vObjProp.NonSolid = Val(GetFromIni(vObj, "NonSolid", vFilePath))
    vObjProp.IsFloor = Val(GetFromIni(vObj, "IsFloor", vFilePath))
    vObjProp.Fixed = Val(GetFromIni(vObj, "Fixed", vFilePath))
    vObjProp.DieSnd = Replace$(GetFromIni(vObj, "DieSnd", vFilePath), "AppPath", App.Path)
    vObjProp.DieSnd = GetShortPath(vObjProp.DieSnd)
    
    If vObjProp.IsFloor Then
        vIsFloor = vIsFloor & vObjProp.CharType
    End If
    If vObjProp.Fixed Then
        vFixed = vFixed & vObjProp.CharType
    End If
    cMapObj.Add vObjProp
    
End Sub
