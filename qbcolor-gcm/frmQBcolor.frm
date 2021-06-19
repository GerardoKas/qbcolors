VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdelante 
      Caption         =   "adelante"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAtras 
      Caption         =   "atras"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblcolor 
      Caption         =   "QbCOLOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub qb(c As Integer)
Me.BackColor = QBColor(i)
End Sub

Private Sub cmdAdelante_Click()
i = i + 1
If i > 15 Then i = 0
qb (i)

lblcolor = i
End Sub

Private Sub cmdAtras_Click()
i = i - 1
If i < 0 Then i = 15
qb (i)
lblcolor = i
End Sub
