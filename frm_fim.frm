VERSION 5.00
Begin VB.Form frm_fim 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   14835
   ClientLeft      =   -60
   ClientTop       =   0
   ClientWidth     =   19050
   LinkTopic       =   "Form2"
   ScaleHeight     =   14835
   ScaleWidth      =   19050
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   2760
   End
   Begin VB.Label lbl_votou 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "VOTO CONCLUÍDO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   555
      Left            =   14400
      TabIndex        =   1
      Top             =   12000
      Width           =   4305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "FIM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   200.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4485
      Left            =   6083
      TabIndex        =   0
      Top             =   5175
      Width           =   6885
   End
End
Attribute VB_Name = "frm_fim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Contador As Integer

Private Sub Form_Load()
            Contador = 0
            Timer1.Enabled = True
          
End Sub
Private Sub Timer1_Timer()
            Contador = Contador + 1
            If Contador = 2 Then
            Unload Me
            End If
End Sub
