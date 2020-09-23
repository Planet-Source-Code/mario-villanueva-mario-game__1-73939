VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Propiedades"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtIsFloor 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtJumpNextLevel 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtNonSolid 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtFixed 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "IsFloor"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "JumpNextLevel"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "NonSolid"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Fixed"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.actualBlock.Fixed = Form2.txtFixed
    Form1.actualBlock.NonSolid = Form2.txtNonSolid
    Form1.actualBlock.JumpNextLevel = Form2.txtJumpNextLevel
    Form1.actualBlock.IsFloor = Form2.txtIsFloor
End Sub
