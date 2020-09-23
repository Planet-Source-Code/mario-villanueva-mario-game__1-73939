VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu New 
         Caption         =   "New Game"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuFullScreen 
         Caption         =   "FullScreen"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuMusicOn 
         Caption         =   "MusicOn"
      End
      Begin VB.Menu mnuEffectsOn 
         Caption         =   "EffectsOn"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuEffectsOn_Click()
    mnuEffectsOn.Checked = Not mnuEffectsOn.Checked
    gEffectsOn = mnuEffectsOn.Checked
End Sub

Private Sub mnuFullscreen_Click()
    mnuFullScreen.Checked = Not mnuFullScreen.Checked
End Sub

Private Sub mnuMusicOn_Click()
    mnuMusicOn.Checked = Not mnuMusicOn.Checked
    gMusicOn = mnuMusicOn.Checked
End Sub

Private Sub New_Click()
    Form1.cScreenWidth = Module1.ScreenWidth
    Form1.cScreenHeight = Module1.ScreenHeight
    Form1.ScaleWidth = Form1.cScreenWidth
    Form1.ScaleHeight = Form1.cScreenHeight
    Form1.Show
    Form1.Start
End Sub
