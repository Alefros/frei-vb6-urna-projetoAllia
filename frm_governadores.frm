VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_governadores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Governadores"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   6360
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   5520
      TabIndex        =   29
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txt_numero 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2520
      TabIndex        =   28
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txt_municipio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   9
      Left            =   1440
      TabIndex        =   27
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Frame lbl_dadoseleitorais 
      Caption         =   "Dados Eleitorais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   9375
      Begin VB.TextBox txt_vice 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   16
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txt_nomepolitico 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   15
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txt_nomereal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   4215
      End
      Begin VB.ComboBox cbo_ano 
         Height          =   315
         Index           =   6
         Left            =   4200
         TabIndex        =   13
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Procurar"
         Height          =   375
         Left            =   6720
         TabIndex        =   12
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ok"
         Height          =   375
         Left            =   8400
         TabIndex        =   11
         Top             =   3120
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   8
         ItemData        =   "frm_governadores.frx":0000
         Left            =   4200
         List            =   "frm_governadores.frx":0002
         TabIndex        =   10
         Top             =   2640
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtp_nascimento 
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   29949953
         CurrentDate     =   40179
      End
      Begin VB.Image img_candidato 
         Height          =   2775
         Left            =   6720
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Nascimento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lbl_vice 
         Caption         =   "Vice:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lbl_nomepolitico 
         Caption         =   "Nome Politico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lbl_nomereal 
         Caption         =   "Nome Completo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbl_numerodopartido 
         Caption         =   "Número do Partido:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   22
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lbl_numerodocandidato 
         Caption         =   "Número do Candidato:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lbl_ano 
         Caption         =   "Ano:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   20
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "UF:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Município:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Auxiliares e Navegação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   5895
      Begin VB.CommandButton cmd_ultimo 
         Caption         =   "Último"
         Height          =   495
         Left            =   4440
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_anterior 
         Caption         =   "Anterior"
         Height          =   495
         Left            =   1560
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_proximo 
         Caption         =   "Próximo"
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_primeiro 
         Caption         =   "Primeiro"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_alterar 
         Caption         =   "Alterar"
         Height          =   495
         Left            =   4440
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_excluir 
         Caption         =   "Excluir"
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_incluir 
         Caption         =   "Incluir"
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_novo 
         Caption         =   "Novo"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_governadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub gravar()
            If status <> "Alteradas" Then tab_pcd.AddNew
            tab_pcd!numero_candidato = txt_numero.Text
            tab_pcd!nome_completo = txt_nomereal.Text
            tab_pcd!nome_politico = txt_nomepolitico.Text
            tab_pcd!data_nascimento = dtp_nascimento
            tab_pcd!vice = txt_vice.Text
            tab_pcd!imagem = img_candidato
            tab_pcd.Update
          
End Sub

Private Sub cmd_alterar_Click()
            status = "Alteradas"
            Call box_1
            Call gravar
End Sub

Private Sub cmd_incluir_Click()
            status = "Incluidas"
            Call box_1
            Call gravar
End Sub

Private Sub Form_Load()
            Call OpenBD
            tab_pcd.Open "Governadores", bdvb, adOpenKeyset, adLockOptimistic
End Sub

