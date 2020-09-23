VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function EnumDisplaySettings Lib "user32" _
Alias "EnumDisplaySettingsA" _
(ByVal lpszDeviceName As Long, _
ByVal iModeNum As Long, _
lpDevMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" _
Alias "ChangeDisplaySettingsA" _
(lpDevMode As Any, ByVal dwFlags As Long) As Long

Private Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved As Long) As Long

 Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                                      (ByVal hwnd As Long, ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long
 Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                       ByVal hWndInsertAfter As Long, _
                                       ByVal x As Long, ByVal y As Long, _
                                       ByVal cx As Long, ByVal cy As Long, _
                                       ByVal wFlags As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

 Private Declare Function GetWindowLong _
     Lib "user32" _
     Alias "GetWindowLongA" ( _
         ByVal hwnd As Long, _
         ByVal nIndex As Long) As Long
         
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, _
    ByVal y As Long) As Long
    
Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_FORCE = 4
Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H4
Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Private Const SWP_SHOWWINDOW = &H40
Private Const WS_POPUP = &H80000000
Private Const HWND_TOP = 0
Private Const HWND_NOTOPMOST = -2
Private Const GWL_STYLE = (-16)
Private Const HWND_TOPMOST = -1

'variables locales para almacenar los valores de las propiedades
Private mvarsWidth As Long 'copia local
Private mvarsHeight As Long 'copia local
Private oldWidth As Long 'copia local
Private oldHeight As Long 'copia local
Private oldBitsPerPixel As Integer
'variables locales para almacenar los valores de las propiedades
Private mvarResChanged As Boolean 'copia local

Public Property Let ResChanged(ByVal vData As Boolean)
    mvarResChanged = vData
End Property

Public Property Get ResChanged() As Boolean
    ResChanged = mvarResChanged
End Property

Public Sub ReturnRes()
    Call mChangeRes(oldWidth, oldHeight, oldBitsPerPixel, True)
    ShowTaskBar
    Form1.cScreenWidth = Module1.ScreenWidth
    Form1.cScreenHeight = Module1.ScreenHeight
    Form1.Move 0, 0, Form1.ScaleX(Form1.cScreenWidth, vbPixels, vbTwips), _
                     Form1.ScaleY(Form1.cScreenHeight, vbPixels, vbTwips)
    Form1.BorderStyle = 1
    Form1.Caption = Form1.Caption
    mvarResChanged = False
End Sub

Public Function ChangeRes(ByVal nWidth As Long, ByVal nHeight As Long) As Long
Dim DevM As DEVMODE, erg As String, an As Long
If mvarResChanged Then Exit Function

ChangeRes = mChangeRes(nWidth, nHeight, 32)
If ChangeRes = 0 Then
    HideTaskBar
    mvarResChanged = True
    Form1.BorderStyle = 0
    Form1.Caption = Form1.Caption
    Form1.Move 0, 0, Form1.ScaleX(Form1.cScreenWidth, vbPixels, vbTwips), _
                     Form1.ScaleY(Form1.cScreenHeight, vbPixels, vbTwips)
Else
    mvarResChanged = False
End If
SetCursorPos nWidth, nHeight
Form1.SetFocus
End Function

Public Function mChangeRes(ByVal nWidth As Long, ByVal nHeight As Long, nBitsPerPixel As Integer, Optional ReturnRes As Boolean = False) As Long
Dim i As Long
Const WS_BORDER = &H800000
Const GWL_EXSTYLE = (-20)
Const SWP_FRAMECHANGED = &H20
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Const SWP_NOSIZE = &H1
Const WS_EX_TOOLWINDOW = &H80
Const GWL_STYLE = (-16)
Const WS_CAPTION = &HC00000
Const ENUM_CURRENT_SETTINGS As Long = -1&

Dim DevM As DEVMODE, erg As String, an As Long
'On Error Resume Next
Form1.Left = 0
Form1.Top = 0
Form1.Width = Form1.ScaleX(nWidth, vbPixels, vbTwips)
Form1.Height = Form1.ScaleY(nHeight, vbPixels, vbTwips)
'Form1.BorderStyle = 0
'DoEvents
'Get the info into DevM
erg = EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, DevM)
'We don't change the colordepth, because a
'rebot will be necessary
If ReturnRes = False Then
    oldBitsPerPixel = DevM.dmBitsPerPel
    oldWidth = DevM.dmPelsWidth
    oldHeight = DevM.dmPelsHeight
End If

DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
DevM.dmBitsPerPel = nBitsPerPixel
DevM.dmPelsWidth = nWidth 'ScreenWidth
DevM.dmPelsHeight = nHeight 'ScreenHeight
'Now change the display and check if possible

erg = ChangeDisplaySettings(DevM, CDS_TEST)

'Check if succesfull

Select Case erg
    Case DISP_CHANGE_RESTART
        an = MsgBox("You've to reboot", vbYesNo + vbSystemModal, "Info")
        If an = vbYes Then
            erg = ExitWindowsEx(EWX_REBOOT, 0&)
        End If
    Case DISP_CHANGE_SUCCESSFUL
        erg = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
    Case Else
        'MsgBox "Mode not supported", vbOKOnly + vbSystemModal, "Error"
        mChangeRes = 1
End Select
SendKeys vbKeySpace
Do While GetQueueStatus(255)
    DoEvents
Loop
End Function

Public Property Let sHeight(ByVal vData As Long)
    mvarsHeight = vData
End Property

Public Property Get sHeight() As Long
    sHeight = mvarsHeight
End Property

Public Property Let sWidth(ByVal vData As Long)
    mvarsWidth = vData
End Property

Public Property Get sWidth() As Long
    sWidth = mvarsWidth
End Property

