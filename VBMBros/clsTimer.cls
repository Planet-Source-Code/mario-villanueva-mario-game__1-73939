VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_timer As Long
Private mvarCaptureFrames As Long 'copia local
Private mvarFrameCount As Long 'copia local
Private m_elapsedTime As Long

Public Function fps() As Long
Dim vTimeperFrame As Long
    vTimeperFrame = timeGetTime()
    If vTimeperFrame < 0 Then
        vTimeperFrame = ULong(vTimeperFrame) - 2147483647
    End If
    vTimeperFrame = vTimeperFrame - m_timer
    If mvarFrameCount = mvarCaptureFrames Then
        If m_elapsedTime + vTimeperFrame > 0 Then
            fps = Fix(mvarCaptureFrames * 1000 / (m_elapsedTime + vTimeperFrame))
        End If
    End If
End Function

Public Function ElapsedTime() As Long
    ElapsedTime = m_elapsedTime + TimePerFrame
End Function

Public Property Let FrameCount(ByVal vData As Long)
    mvarFrameCount = vData
End Property

Public Property Get FrameCount() As Long
    FrameCount = mvarFrameCount
End Property

Public Property Let CaptureFrames(ByVal vData As Long)
    mvarCaptureFrames = vData
End Property

Public Property Get CaptureFrames() As Long
    CaptureFrames = mvarCaptureFrames
End Property

Public Sub StartFrame()
Dim vTimeperFrame As Long

    If mvarFrameCount = mvarCaptureFrames Then
        m_elapsedTime = 0
        mvarFrameCount = 0
    Else
        vTimeperFrame = timeGetTime() '- m_timer
        If vTimeperFrame < 0 Then
            vTimeperFrame = ULong(vTimeperFrame) - 2147483647
        End If
        vTimeperFrame = vTimeperFrame - m_timer
        m_elapsedTime = m_elapsedTime + vTimeperFrame
    End If
    mvarFrameCount = mvarFrameCount + 1
    m_timer = timeGetTime()
    If m_timer < 0 Then
        m_timer = ULong(m_timer) - 2147483647
    End If
End Sub

Public Function TimePerFrame() As Long
    TimePerFrame = timeGetTime()
    If TimePerFrame < 0 Then
        TimePerFrame = ULong(TimePerFrame) - 2147483647
    End If
    TimePerFrame = TimePerFrame - m_timer
End Function

Private Function ULong(ByVal Int32 As Long) As Currency
Const OFFSET_4 As Currency = 4294967296@
  If Int32 < 0 Then
    ULong = CCur(Int32) + OFFSET_4
  Else
    ULong = Int32
  End If
End Function
