VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_localizacoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro e Controle de Localizações"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   1815
      Left            =   6120
      TabIndex        =   20
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3201
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frm_localizacoes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Tab 3"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Tab 4"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cadastro e Controle de Localizações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txt_area 
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txt_regiao 
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txt_municipio 
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txt_estado 
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txt_pais 
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Áreas:"
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Região:"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Município:"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Estado:"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "País:"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   615
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
      Top             =   3120
      Width           =   5895
      Begin VB.CommandButton cmd_ultimo 
         Caption         =   "Último"
         Height          =   495
         Left            =   4440
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_proximo 
         Caption         =   "Próximo"
         Height          =   495
         Left            =   3000
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_anterior 
         Caption         =   "Anterior"
         Height          =   495
         Left            =   1560
         TabIndex        =   17
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
Attribute VB_Name = "frm_localizacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub localizacoes()
            If status <> "Alteradas" Then tab_loca.AddNew
            tab_loca!pais = txt_pais.Text
            tab_loca!municipio = txt_municipio.Text
            tab_loca!regiao = txt_regiao.Text
            tab_loca!area = txt_area.Text
            tab_loca.Update
End Sub

Private Sub cmd_novo_Click()
            txt_pais = Clear
            txt_estado = Clear
            txt_municipio = Clear
            txt_regiao = Clear
            txt_area = Clear
End Sub
