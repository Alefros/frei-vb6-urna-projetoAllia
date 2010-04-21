VERSION 5.00
Begin VB.Form frm_gvoto 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Voto"
   ClientHeight    =   9450
   ClientLeft      =   885
   ClientTop       =   795
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_ladrao 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      MaxLength       =   2
      TabIndex        =   0
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmd_branco 
      BackColor       =   &H8000000E&
      Caption         =   "BRANCO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      MaskColor       =   &H8000000E&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton cmd_corrigir 
      BackColor       =   &H000000FF&
      Caption         =   "CORRIGE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   5160
      MaskColor       =   &H8000000E&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton cmd_confirmar 
      BackColor       =   &H00008000&
      Caption         =   "CONFIRMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7200
      MaskColor       =   &H8000000E&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label lbl_laranja 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2280
      TabIndex        =   15
      Top             =   8040
      Width           =   105
   End
   Begin VB.Label lbl_verde 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2280
      TabIndex        =   14
      Top             =   7440
      Width           =   105
   End
   Begin VB.Label lbl_tecla 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   240
      TabIndex        =   13
      Top             =   6960
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   0
      Top             =   6720
      Width           =   13455
   End
   Begin VB.Label lbl_partido2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label lbl_partido 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label lbl_vicepresidente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Label lbl_vice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label lbl_candidato 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label lbl_presidente 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   105
   End
   Begin VB.Label lbl_numero 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lbl_mestre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "GOVERNADOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label lbl_seuvotopara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   3135
   End
   Begin VB.Image img_ladrao 
      Height          =   5415
      Left            =   7440
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5295
   End
End
Attribute VB_Name = "frm_gvoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
