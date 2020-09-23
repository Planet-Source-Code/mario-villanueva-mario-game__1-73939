Attribute VB_Name = "Module1"
Option Explicit

Public Const QS_HOTKEY = &H80
Public Const QS_KEY = &H1
Public Const QS_MOUSEBUTTON = &H4
Public Const QS_MOUSEMOVE = &H2
Public Const QS_PAINT = &H20
Public Const QS_POSTMESSAGE = &H8
Public Const QS_SENDMESSAGE = &H40
Public Const QS_TIMER = &H10
Public Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Public Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Public Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Public Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)

Const WS_BORDER = &H800000
Const WS_DLGFRAME = &H400000
Const WS_THICKFRAME = &H40000
Const WS_CAPTION = &HC00000                  ' WS_BORDER Or WS_DLGFRAME
Const WS_EX_CLIENTEDGE = &H200

Private Declare Function AdjustWindowRectEx Lib "user32" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0 'X Size of screen
Public Const SM_CYSCREEN = 1 'Y Size of Screen
Public Const SM_CXVSCROLL = 2 'X Size of arrow in vertical scroll bar.
Public Const SM_CYHSCROLL = 3 'Y Size of arrow in horizontal scroll bar
Public Const SM_CYCAPTION = 4 'Height of windows caption
Public Const SM_CXBORDER = 5 'Width of no-sizable borders
Public Const SM_CYBORDER = 6 'Height of non-sizable borders
Public Const SM_CXDLGFRAME = 7 'Width of dialog box borders
Public Const SM_CYDLGFRAME = 8 'Height of dialog box borders
Public Const SM_CYVTHUMB = 9 'Height of scroll box on horizontal scroll bar
Public Const SM_CXHTHUMB = 10 ' Width of scroll box on horizontal scroll bar
Public Const SM_CXICON = 11 'Width of standard icon
Public Const SM_CYICON = 12 'Height of standard icon
Public Const SM_CXCURSOR = 13 'Width of standard cursor
Public Const SM_CYCURSOR = 14 'Height of standard cursor
Public Const SM_CYMENU = 15 'Height of menu
Public Const SM_CXFULLSCREEN = 16 'Width of client area of maximized window
Public Const SM_CYFULLSCREEN = 17 'Height of client area of maximized window
Public Const SM_CYKANJIWINDOW = 18 'Height of Kanji window
Public Const SM_MOUSEPRESENT = 19 'True is a mouse is present
Public Const SM_CYVSCROLL = 20 'Height of arrow in vertical scroll bar
Public Const SM_CXHSCROLL = 21 'Width of arrow in vertical scroll bar
Public Const SM_DEBUG = 22 'True if deugging version of windows is running
Public Const SM_SWAPBUTTON = 23 'True if left and right buttons are swapped.
Public Const SM_CXMIN = 28 'Minimum width of window
Public Const SM_CYMIN = 29 'Minimum height of window
Public Const SM_CXSIZE = 30 'Width of title bar bitmaps
Public Const SM_CYSIZE = 31 'height of title bar bitmaps
Public Const SM_CXMINTRACK = 34 'Minimum tracking width of window
Public Const SM_CYMINTRACK = 35 'Minimum tracking height of window
Public Const SM_CXDOUBLECLK = 36 'double click width
Public Const SM_CYDOUBLECLK = 37 'double click height
Public Const SM_CXICONSPACING = 38 'width between desktop icons
Public Const SM_CYICONSPACING = 39 'height between desktop icons
Public Const SM_MENUDROPALIGNMENT = 40 'Zero if popup menus are aligned to the left of the memu bar item. True if it is aligned to the right.
Public Const SM_PENWINDOWS = 41 'The handle of the pen windows DLL if loaded.
Public Const SM_DBCSENABLED = 42 'True if double byte characteds are enabled
Public Const SM_CMOUSEBUTTONS = 43 'Number of mouse buttons.
Public Const SM_CMETRICS = 44 'Number of system metrics
Public Const SM_CLEANBOOT = 67 'Windows 95 boot mode. 0 = normal, 1 = safe, 2 = safe with network
Public Const SM_CXMAXIMIZED = 61 'default width of win95 maximised window
Public Const SM_CXMAXTRACK = 59 'maximum width when resizing win95 windows
Public Const SM_CXMENUCHECK = 71 'width of menu checkmark bitmap
Public Const SM_CXMENUSIZE = 54 'width of button on menu bar
Public Const SM_CXMINIMIZED = 57 'width of rectangle into which minimised windows must fit.
Public Const SM_CYMAXIMIZED = 62 'default height of win95 maximised window
Public Const SM_CYMAXTRACK = 60 'maximum width when resizing win95 windows
Public Const SM_CYMENUCHECK = 72 'height of menu checkmark bitmap
Public Const SM_CYMENUSIZE = 55 'height of button on menu bar
Public Const SM_CYMINIMIZED = 58 'height of rectangle into which minimised windows must fit.
Public Const SM_CYSMCAPTION = 51 'height of windows 95 small caption
Public Const SM_MIDEASTENABLED = 74 'Hebrw and Arabic enabled for windows 95
Public Const SM_NETWORK = 63 'bit o is set if a network is present. Const SM_SECURE = 44 'True

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
As String) As Long

Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40

Public gMusicOn As Boolean
Public gEffectsOn As Boolean

'Global Const SND_ASYNC = &H1
'Global Const SND_NODEFAULT = &H2
'Global Const SND_FILENAME = &H20000

Public m_mult As Long
Public rctDst As RECT, rctSrc As RECT

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public colObjProp As Collection
Public vObjProp As clsObjProp
Public Col As Collection
Public cSnd As Collection
Public cBrkFx As Collection
Public cBreakeable As Collection
Public vFixed As String
Public vIsFloor As String
Public vBreakeableIdx As Long
Public vGameMusic As String
Public cLevel As clsLevel
Public gNextlevelChar As String
Public KeyIdx As Long
Public Const ScreenWidth = 320
Public Const ScreenHeight = 240

Public vSwitchGifts As Boolean
Public vGifts As String
Public arrGifts() As String

Public Function cGetInputState()
Dim qsRet As Long
qsRet = GetQueueStatus(QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT)
                   cGetInputState = qsRet
End Function

Public Sub TransparentBlt(dest As Long, rctDest As RECT _
                         , Src As Long, rctSrc As RECT _
                         , TransColor As Long, Optional flip = False)
                        
    Dim workdc As Long, mask1dc As Long
    Dim mask2dc As Long, destdc As Long, srcdc As Long
    Dim dwidth As Long, dheight As Long
    Dim sWidth As Long, sHeight As Long
    Dim maskDC As Long      'DC for the mask
    Dim tempDC As Long      'DC for temporary data
    Dim hMaskBmp As Long    'Bitmap for mask
    Dim hTempBmp As Long    'Bitmap for temporary data
    Dim oldMaskBmp As Long    'Bitmap for mask
    Dim oldTempBmp As Long    'Bitmap for temporary data
    Dim tmp As Long
    
    With rctDest
        dwidth = Abs(.Right - .Left)
        dheight = Abs(.Top - .Bottom)
    End With
    
    With rctSrc
        sWidth = Abs(.Right - .Left)
        sHeight = Abs(.Top - .Bottom)
    End With

   If dwidth = 0 Or dheight = 0 Or _
      sWidth = 0 Or sHeight = 0 Then
      Exit Sub
   End If

    'First, create some DC's. These are our gateways to associated
    'bitmaps in RAM
    maskDC = CreateCompatibleDC(destdc)
    tempDC = CreateCompatibleDC(destdc)
    
    'Then, we need the bitmaps. Note that we create a monochrome
    'bitmap here!
    'This is a trick we use for creating a mask fast enough.
    hMaskBmp = CreateBitmap(sWidth, sHeight, 1, 1, ByVal 0&)
    hTempBmp = CreateCompatibleBitmap(dest, sWidth, sHeight)

    'Then we can assign the bitmaps to the DCs
    oldMaskBmp = SelectObject(maskDC, hMaskBmp)
    oldTempBmp = SelectObject(tempDC, hTempBmp)
    
    'Now we can create a mask. First, we set the background color
    'to the transparent color; then we copy the image into the
    'monochrome bitmap.
    'When we are done, we reset the background color of the
    'original source.
    TransColor = SetBkColor(Src, TransColor)
    BitBlt maskDC, 0, 0, sWidth, sHeight, Src, rctSrc.Left, rctSrc.Top, vbSrcCopy
    TransColor = SetBkColor(Src, TransColor)

    'The first we do with the mask is to MergePaint it into the
    'destination.
    'This will punch a WHITE hole in the background exactly were
    'we want the graphics to be painted in.
    BitBlt tempDC, 0, 0, sWidth, sHeight, maskDC, 0, 0, vbSrcCopy
    If flip Then
        If m_mult = 1 Then
            StretchBlt dest, rctDest.Left + sWidth, rctDest.Top, -sWidth, sHeight, tempDC, 0, 0, sWidth, sHeight, vbMergePaint
        Else
            StretchBlt dest, rctDest.Left * m_mult, rctDest.Top * m_mult, dwidth * m_mult, dheight * m_mult, tempDC, rctSrc.Right - rctSrc.Left - 1, 0, -sWidth, sHeight, vbMergePaint
        End If
    Else
        If m_mult = 1 Then
            BitBlt dest, rctDest.Left, rctDest.Top, sWidth, sHeight, tempDC, 0, 0, vbMergePaint
        Else
            StretchBlt dest, rctDest.Left * m_mult, rctDest.Top * m_mult, dwidth * m_mult, dheight * m_mult, tempDC, 0, 0, sWidth, sHeight, vbMergePaint
        End If
    End If
    'Now we delete the transparent part of our source image. To do
    'this, we must invert the mask and MergePaint it into the
    'source image. The transparent area will now appear as WHITE.
    BitBlt maskDC, 0, 0, sWidth, sHeight, maskDC, 0, 0, vbNotSrcCopy
    BitBlt tempDC, 0, 0, sWidth, sHeight, Src, rctSrc.Left, rctSrc.Top, vbSrcCopy
    BitBlt tempDC, 0, 0, sWidth, sHeight, maskDC, 0, 0, vbMergePaint

    'Both target and source are clean. All we have to do is to AND
    'them together!
    If flip Then
        If m_mult = 1 Then
            StretchBlt dest, rctDest.Left + sWidth, rctDest.Top, -sWidth, sHeight, tempDC, 0, 0, sWidth, sHeight, vbSrcAnd
        Else
            StretchBlt dest, rctDest.Left * m_mult, rctDest.Top * m_mult, dwidth * m_mult, dheight * m_mult, tempDC, rctSrc.Right - rctSrc.Left - 1, 0, -sWidth, sHeight, vbSrcAnd
        End If
    Else
        If m_mult = 1 Then
            BitBlt dest, rctDest.Left, rctDest.Top, sWidth, sHeight, tempDC, 0, 0, vbSrcAnd
        Else
            StretchBlt dest, rctDest.Left * m_mult, rctDest.Top * m_mult, dwidth * m_mult, dheight * m_mult, tempDC, 0, 0, sWidth, sHeight, vbSrcAnd
        End If
    End If
    'Now all we have to do is to clean up after us and free system
    'resources..
    DeleteObject (SelectObject(maskDC, oldMaskBmp))
    DeleteObject (SelectObject(tempDC, oldTempBmp))
    
    DeleteDC (maskDC)
    DeleteDC (tempDC)
End Sub

Public Sub Paint(dest As Long, rctDest As RECT _
               , Src As Long, rctSrc As RECT, Optional dwRop = vbSrcCopy)
                        
    Dim dwidth As Long, dheight As Long
    Dim sWidth As Long, sHeight As Long

    
    With rctDest
        dwidth = Abs(.Right - .Left)
        dheight = Abs(.Top - .Bottom)
    End With

   If dwidth = 0 Or dheight = 0 Then Exit Sub
   
    'BitBlt dest, rctDest.left, rctDest.top, dwidth, dheight, Src, rctSrc.left, rctSrc.top, dwRop
    StretchBlt dest, rctDest.Left, rctDest.Top, dwidth, dheight, Src, rctSrc.Left, rctSrc.Top, rctSrc.Right - rctSrc.Left, rctSrc.Bottom - rctSrc.Top, dwRop
End Sub

Public Sub StretchPaint(dest As Long, rctDest As RECT _
                        , Src As Long, rctSrc As RECT)
    Dim dwidth As Long, dheight As Long
    Dim sWidth As Long, sHeight As Long
    
    If Abs(rctDest.Right - rctDest.Left) = Abs(rctSrc.Right - rctSrc.Left) And _
       Abs(rctDest.Bottom - rctDest.Top) = Abs(rctSrc.Bottom - rctSrc.Top) Then
        Paint dest, rctDest, Src, rctSrc, vbSrcCopy
        Exit Sub
    End If
    
    With rctDest
        dwidth = Abs(.Right - .Left)
        dheight = Abs(.Top - .Bottom)
    End With
    
    With rctSrc
        sWidth = Abs(.Right - .Left)
        sHeight = Abs(.Top - .Bottom)
    End With

   If dwidth = 0 Or dheight = 0 Or _
      sWidth = 0 Or sHeight = 0 Then
      Exit Sub
   End If
    StretchBlt dest, rctDest.Left * m_mult, rctDest.Top * m_mult, dwidth * m_mult, dheight * m_mult, Src, rctSrc.Left, rctSrc.Top, sWidth, sHeight, vbSrcCopy
End Sub

Public Function ReadBFile(vFile As String) As String
Dim txtVar As String, i As Integer
i = FreeFile
On Error GoTo errHandler
Open vFile For Binary As #i
    txtVar = Space$(LOF(i))
    Get #i, , txtVar
Close #i
ReadBFile = txtVar
Exit Function
errHandler:
    MsgBox Err.Description
End Function

Public Sub MoveProp(pObjProp As clsObjProp, pObj As clsObject)
    If Not pObj Is Nothing Then
    pObj.hdc = pObjProp.hdc
    pObj.CanHit = pObjProp.CanHit
    'pObj.FireBall = pObjProp.FireBall
    pObj.CanCrouch = pObjProp.CanCrouch
    pObj.CrouchFrame = pObjProp.CrouchFrame
    pObj.HittedByTop = pObjProp.HittedByTop
    pObj.FireBallCanHit = pObjProp.FireBallCanHit
    pObj.HittedByLeft = pObjProp.HittedByLeft
    pObj.HittedByRight = pObjProp.HittedByRight
    pObj.HittedByBottom = pObjProp.HittedByBottom
    pObj.SourceLeft = pObjProp.SourceLeft
    pObj.SourceRight = pObjProp.SourceRight
    pObj.SourceTop = pObjProp.SourceTop
    pObj.SourceBottom = pObjProp.SourceBottom
    pObj.MinRunframe = pObjProp.MinRunframe
    pObj.Velocity = pObjProp.Velocity
    pObj.JumpVelocity = pObjProp.JumpVelocity
    pObj.MaxRunframe = pObjProp.MaxRunframe
    pObj.InitedAI = pObjProp.InitedAI
    pObj.PosLeft = pObjProp.PosLeft
    pObj.PosTop = pObjProp.PosTop
    pObj.PosRight = pObjProp.PosRight
    pObj.PosBottom = pObjProp.PosBottom
    pObj.CreateWhenHitted = pObjProp.CreateWhenHitted 'Ex "M"
    pObj.CreatePlace = pObjProp.CreatePlace 'Ex OnTop
    pObj.Visible = pObjProp.Visible
    pObj.User = pObjProp.User
    pObj.CanFall = pObjProp.CanFall
    pObj.StartPosOffsetX = pObjProp.StartPosOffsetX
    pObj.ChangeFrom = pObjProp.ChangeFrom
    pObj.JumpFrom = pObjProp.JumpFrom
    pObj.JumpTo = pObjProp.JumpTo
    pObj.JumpSize = pObjProp.JumpSize
    pObj.AnimFrom = pObjProp.AnimFrom
    pObj.AnimTo = pObjProp.AnimTo
    pObj.DieFrame = pObjProp.DieFrame
    pObj.DieTiming = pObjProp.DieTiming
    pObj.RemoveWhenDies = pObjProp.RemoveWhenDies
    pObj.MakeJumpWhenHitted = pObjProp.MakeJumpWhenHitted
    pObj.JumpWhenHitted = pObjProp.JumpWhenHitted
    pObj.Solid = pObjProp.Solid
    pObj.Fixed = pObjProp.Fixed
    pObj.InitAIWhenHitted = pObjProp.InitAIWhenHitted
    pObj.CharType = pObjProp.CharType
    pObj.ChangePlayerTo = pObjProp.ChangePlayerTo
    pObj.AI = pObjProp.AI
    pObj.NextLevel = pObjProp.NextLevel
    pObj.Hibernating = pObjProp.Hibernating
    pObj.CanHitEnemies = pObjProp.CanHitEnemies
    pObj.CannotHitUser = pObjProp.CannotHitUser
    pObj.CanBeBreaked = pObjProp.CanBeBreaked
    pObj.CanBreak = pObjProp.CanBreak
    pObj.direction = pObjProp.direction
    pObj.UserSelection = pObjProp.UserSelection
    pObj.GrowTo = pObjProp.GrowTo
    pObj.MakeGrow = pObjProp.MakeGrow
    pObj.IsFloor = pObjProp.IsFloor
    pObj.DieWhenHits = pObjProp.DieWhenHits
    pObj.JumpSnd = pObjProp.JumpSnd
    pObj.DieSnd = pObjProp.DieSnd
    pObj.FireSnd = pObjProp.FireSnd
    pObj.FireBall = pObjProp.FireBall
    pObj.Raising = pObjProp.Raising
    pObj.Descending = pObjProp.Descending
    pObj.Raisetime = pObjProp.Raisetime
    End If
End Sub

Public Sub BltMask(dest As Long, rctDest As RECT _
                   , Src As Long, rctSrc As RECT _
                   , TransColor As Long, mask As Long)
                        
    Dim workdc As Long, mask1dc As Long
    Dim mask2dc As Long, destdc As Long, srcdc As Long
    Dim dwidth As Long, dheight As Long
    Dim sWidth As Long, sHeight As Long
    Dim maskDC As Long      'DC for the mask
    Dim tempDC As Long      'DC for temporary data
    Dim hMaskBmp As Long    'Bitmap for mask
    Dim hTempBmp As Long    'Bitmap for temporary data
    Dim oldMaskBmp As Long    'Bitmap for mask
    Dim oldTempBmp As Long    'Bitmap for temporary data
    Dim tmp As Long
    
    With rctDest
        dwidth = Abs(.Right - .Left)
        dheight = Abs(.Top - .Bottom)
    End With
    
    With rctSrc
        sWidth = Abs(.Right - .Left)
        sHeight = Abs(.Top - .Bottom)
    End With

   If dwidth = 0 Or dheight = 0 Or _
      sWidth = 0 Or sHeight = 0 Then
      Exit Sub
   End If

    'First, create some DC's. These are our gateways to associated
    'bitmaps in RAM
    maskDC = CreateCompatibleDC(destdc)
    tempDC = CreateCompatibleDC(destdc)
    
    'Then, we need the bitmaps. Note that we create a monochrome
    'bitmap here!
    'This is a trick we use for creating a mask fast enough.
    hMaskBmp = CreateBitmap(dwidth, sHeight, 1, 1, ByVal 0&)
    hTempBmp = CreateCompatibleBitmap(dest, sWidth, sHeight)

    'Then we can assign the bitmaps to the DCs
    oldMaskBmp = SelectObject(maskDC, hMaskBmp)
    oldTempBmp = SelectObject(tempDC, hTempBmp)
    
    'Now we can create a mask. First, we set the background color
    'to the transparent color; then we copy the image into the
    'monochrome bitmap.
    'When we are done, we reset the background color of the
    'original source.
    TransColor = SetBkColor(Src, TransColor)
    BitBlt maskDC, 0, 0, sWidth, sHeight, Src, rctSrc.Left, rctSrc.Top, vbSrcCopy
    TransColor = SetBkColor(Src, TransColor)

    'The first we do with the mask is to MergePaint it into the
    'destination.
    'This will punch a WHITE hole in the background exactly were
    'we want the graphics to be painted in.
    BitBlt tempDC, 0, 0, sWidth, sHeight, maskDC, 0, 0, vbSrcCopy
    BitBlt mask, 0, 0, sWidth, sHeight, tempDC, 0, 0, vbSrcCopy
    'Hacer un bitblt de Mask usando vbMergePaint
    'If flip Then
    '    If m_mult = 1 Then
    '        StretchBlt dest, rctDest.left + dwidth, rctDest.top, -dwidth, dheight, tempDC, 0, 0, sWidth, sHeight, vbMergePaint
    '    Else
    '        StretchBlt dest, rctDest.left * m_mult, rctDest.top * m_mult, dwidth * m_mult, dheight * m_mult, tempDC, rctSrc.right - rctSrc.left - 1, 0, -sWidth, sHeight, vbMergePaint
    '    End If
    'Else
    '    If m_mult = 1 Then
    '        BitBlt dest, rctDest.left, rctDest.top, dwidth, dheight, tempDC, 0, 0, vbMergePaint
    '    Else
    '        StretchBlt dest, rctDest.left * m_mult, rctDest.top * m_mult, dwidth * m_mult, dheight * m_mult, tempDC, 0, 0, sWidth, sHeight, vbMergePaint
    '    End If
    'End If
    'Now we delete the transparent part of our source image. To do
    'this, we must invert the mask and MergePaint it into the
    'source image. The transparent area will now appear as WHITE.
    BitBlt maskDC, 0, 0, sWidth, sHeight, maskDC, 0, 0, vbNotSrcCopy
    BitBlt tempDC, 0, 0, sWidth, sHeight, Src, rctSrc.Left, rctSrc.Top, vbSrcCopy
    BitBlt tempDC, 0, 0, sWidth, sHeight, maskDC, 0, 0, vbMergePaint
    BitBlt dest, 0, 0, sWidth, sHeight, tempDC, 0, 0, vbSrcCopy
    'Hacer un Bitblt en el lugar final usando vbSrcAnd
    
    'Both target and source are clean. All we have to do is to AND
    'them together!
    'If flip Then
    '    If m_mult = 1 Then
    '        StretchBlt dest, rctDest.left + dwidth, rctDest.top, -sWidth, sHeight, tempDC, 0, 0, sWidth, sHeight, vbSrcAnd
    '    Else
    '        StretchBlt dest, rctDest.left * m_mult, rctDest.top * m_mult, dwidth * m_mult, dheight * m_mult, tempDC, rctSrc.right - rctSrc.left - 1, 0, -sWidth, sHeight, vbSrcAnd
    '    End If
    'Else
    '    If m_mult = 1 Then
    '        BitBlt dest, rctDest.left, rctDest.top, dwidth, dheight, tempDC, 0, 0, vbSrcAnd
    '    Else
    '        StretchBlt dest, rctDest.left * m_mult, rctDest.top * m_mult, dwidth * m_mult, dheight * m_mult, tempDC, 0, 0, sWidth, sHeight, vbSrcAnd
    '    End If
    'End If
    'Now all we have to do is to clean up after us and free system
    'resources..
    DeleteObject (SelectObject(maskDC, oldMaskBmp))
    DeleteObject (SelectObject(tempDC, oldTempBmp))
    
    DeleteDC (maskDC)
    DeleteDC (tempDC)
End Sub

Public Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    '*** DESCRIPTION:Reads from an *.INI file strFileName (full path & file name)
    '*** RETURNS:The string stored in [strSectionHeader], line beginning strVariableName=
    '*** NOTE: Requires declaration of API call GetPrivateProfileString
    'Initialise variable
    Dim strReturn As String
    'Blank the return string
    strReturn = String(1024, Chr(0))
    'Get requested information, trimming the
    '     returned string
    GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function

Public Function IsInCol(pObj As Collection, char As String) As String
Dim ObjProp As clsObjProp, i As Long
On Error GoTo errFound
    For i = 1 To pObj.Count
        If pObj(i).CharType = char Then
            IsInCol = char
            Exit Function
        End If
    Next
    Exit Function
errFound:
On Error GoTo 0
    IsInCol = ""
End Function

Public Function IsKey(pObj As Collection, Key As String) As Boolean
Dim i As Long
On Error GoTo errFound
    If pObj.Count > 0 Then
    For i = 1 To pObj.Count
        If pObj(i) = Key Then
            IsKey = True
        End If
    Next
    End If
    Exit Function
errFound:
On Error GoTo 0
    
    Err.Raise Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub PlayMIDI(MIDIFile As String)
    Dim SafeFile As String
    If gMusicOn = False Or MIDIFile = "" Then Exit Sub
    MIDIFile = GetShortPath(MIDIFile)
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub

Public Sub StopMIDI(MIDIFile As String)
    Dim SafeFile As String
    'If MIDIFile = "" Then Exit Sub
    MIDIFile = GetShortPath(MIDIFile)
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub

Public Sub Playwav(WAVFile As String)
    Dim ret As Long, BrkFx As Variant, sStatus As String * 255, nReturn As Long
    Static i As Long
    If i > 5 Then i = 0
    If gEffectsOn = False Or WAVFile = "" Then Exit Sub
    'Si hay alguno abierto cerrarlo
    For Each BrkFx In cBrkFx
        nReturn = mciSendString("Status " & BrkFx & " Mode", sStatus, 255, 0)
        If Left$(sStatus, InStr(sStatus, Chr$(0)) - 1) = "stopped" Then
            Call mciSendString("Close " & BrkFx, 0&, 0, 0)
            cBrkFx.Remove BrkFx
        End If
    Next
    Do While i < 6
        i = i + 1
        nReturn = mciSendString("Status " & WAVFile & i & " Mode", sStatus, 255, 0)
        If Left$(sStatus, InStr(sStatus, Chr$(0)) - 1) = "stopped" Or _
           Left$(sStatus, InStr(sStatus, Chr$(0)) - 1) = "closed" Or _
           Left$(sStatus, InStr(sStatus, Chr$(0)) - 1) = "" Then
            Call mciSendString("Open " & WAVFile & " type waveaudio Alias " & WAVFile & i, 0&, 0, 0)
            Call mciSendString("play " & WAVFile & i & " from 0", 0&, 0, 0)
            cBrkFx.Add WAVFile & i, WAVFile & i
            Exit Sub
        End If
    Loop
    
    
End Sub

Public Sub Playwav3(WAVFile As String)
    Dim ret As Long
    If gEffectsOn = False Or WAVFile = "" Then Exit Sub
    Call mciSendString("play " & WAVFile & " from 0", 0&, 0, 0)
End Sub

Public Function GetShortPath(strFileName As String) As String
    Dim lngRes As Long, strPath As String
    'Create a buffer
    strPath = String$(165, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(strFileName, strPath, 164)
    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPath, lngRes)
End Function

Public Function HideTaskBar() As Boolean
Dim lRet As Long
    lRet = FindWindow("Shell_traywnd", "")
    If lRet > 0 Then
        lRet = SetWindowPos(lRet, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
        HideTaskBar = lRet > 0
    End If
End Function

Public Function ShowTaskBar() As Boolean
Dim lRet As Long
lRet = FindWindow("Shell_traywnd", "")
If lRet > 0 Then
    lRet = SetWindowPos(lRet, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    ShowTaskBar = lRet > 0
End If
End Function


