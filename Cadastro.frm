VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_cadastro 
   Caption         =   "Cadastro de Candidatos"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   9690
   Begin VB.CommandButton cmd_ultimo 
      Caption         =   "Último"
      Height          =   495
      Left            =   4560
      TabIndex        =   50
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmd_anterior 
      Caption         =   "Anterior"
      Height          =   495
      Left            =   3120
      TabIndex        =   49
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmd_proximo 
      Caption         =   "Próximo"
      Height          =   495
      Left            =   1680
      TabIndex        =   48
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmd_primeiro 
      Caption         =   "Primeiro"
      Height          =   495
      Left            =   240
      TabIndex        =   47
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "Alterar"
      Height          =   495
      Left            =   4560
      TabIndex        =   46
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   3120
      TabIndex        =   45
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmd_incluir 
      Caption         =   "Incluir"
      Height          =   495
      Left            =   1680
      TabIndex        =   44
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "Novo"
      Height          =   495
      Left            =   240
      TabIndex        =   43
      Top             =   5160
      Width           =   1335
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
      Height          =   1455
      Left            =   120
      TabIndex        =   42
      Top             =   4920
      Width           =   5895
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
      Height          =   285
      Left            =   1680
      TabIndex        =   29
      Top             =   3120
      Width           =   4215
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
      Height          =   1935
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   9375
      Begin VB.ComboBox cbo_area 
         Height          =   315
         Left            =   5400
         TabIndex        =   40
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cbo_ano 
         Height          =   315
         Left            =   8040
         TabIndex        =   38
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cbo_cargo 
         Height          =   315
         ItemData        =   "Cadastro.frx":0000
         Left            =   840
         List            =   "Cadastro.frx":000A
         TabIndex        =   36
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txt_cpco 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   35
         Top             =   960
         Width           =   3855
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
         Height          =   285
         Left            =   4800
         TabIndex        =   34
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox cbo_coligacao 
         Height          =   315
         Left            =   1200
         TabIndex        =   33
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox cbo_sigla 
         Height          =   315
         Left            =   1680
         TabIndex        =   32
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cbo_situacao 
         Height          =   315
         Left            =   6480
         TabIndex        =   31
         Top             =   600
         Width           =   2775
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
         Height          =   285
         Left            =   8040
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl_areadepoder 
         Caption         =   "Área de Poder:"
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
         Left            =   3840
         TabIndex        =   39
         Top             =   1320
         Width           =   1455
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
         Height          =   255
         Left            =   7560
         TabIndex        =   37
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lbl_nomepolitico 
         Caption         =   "Nome Político:"
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
         TabIndex        =   28
         Top             =   240
         Width           =   1335
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
         Height          =   255
         Left            =   5880
         TabIndex        =   27
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lbl_situacao 
         Caption         =   "Situação:"
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
         Left            =   5520
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lbl_sigladopartido 
         Caption         =   "Sigla do Partido:"
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
         Top             =   600
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
         Left            =   3000
         TabIndex        =   24
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lbl_nomedacoligacao 
         Caption         =   "Coligação:"
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
         TabIndex        =   23
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbl_composicaodacoligacao 
         Caption         =   "Composição da Coligação:"
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
         Left            =   2880
         TabIndex        =   22
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lbl_cargo 
         Caption         =   "Cargo:"
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
         TabIndex        =   21
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.Frame fra_dadospessoais 
      Caption         =   "Dados Pessoais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.CommandButton cmd_imagem 
         Caption         =   "Inserir imagem"
         Height          =   375
         Left            =   4080
         TabIndex        =   41
         Top             =   2160
         Width           =   2055
      End
      Begin VB.ComboBox cbo_ocupacao 
         Height          =   315
         Left            =   1320
         TabIndex        =   19
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox cbo_instrucao 
         Height          =   315
         Left            =   4680
         TabIndex        =   18
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox cbo_civil 
         Height          =   315
         Left            =   1320
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox cbo_nacionalidade 
         Height          =   315
         Left            =   3000
         TabIndex        =   16
         Top             =   1440
         Width           =   3135
      End
      Begin VB.ComboBox cbo_uf 
         Height          =   315
         Left            =   480
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox cbo_municipio 
         Height          =   315
         Left            =   2520
         TabIndex        =   13
         Top             =   1080
         Width           =   3615
      End
      Begin VB.OptionButton opt_feminino 
         Alignment       =   1  'Right Justify
         Caption         =   "Feminino"
         Height          =   255
         Left            =   5160
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton opt_masculino 
         Caption         =   "Masculino"
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txt_nomecompleto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   4455
      End
      Begin MSMask.MaskEdBox msk_nascimento 
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.Image img_candidato 
         Height          =   1935
         Left            =   6840
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lbl_uf 
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
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label lbl_municipiodenascimento 
         Caption         =   "Município de Nascimento:"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lbl_nacionalidade 
         Caption         =   "Nacionalidade:"
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
         Left            =   1560
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lbl_graudeinstrucao 
         Caption         =   "Grau de Instrução:"
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
         Left            =   3000
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lbl_ocupacao 
         Caption         =   "Ocupação:"
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
         TabIndex        =   5
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lbl_estadocivil 
         Caption         =   "Estado Civil:"
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
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl_datadenascimento 
         Caption         =   "Data de Nascimento:"
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
         TabIndex        =   3
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lbl_sexo 
         Caption         =   "Sexo:"
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
         Left            =   3360
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lbl_nomecompleto 
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
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
            frm_abrir.Show
End Sub

Private Sub cmd_incluir_Click()
            If status <> "Alteradas" Then tab_Vend.AddNew
            If cbo_cargo.Text = Presidentes Then
            tab_pcd!numero_candidato = txt_numero
            tab_pcd!nome_completo = txt_nomecompleto
            tab_pcd!nome_politico = txt_nomepolitico
            tab_pcd!data_nascimento = msk_nascimento
            tab_pcd!cod_mun_nasc = cbo_municipio
            tab_pcd!vice =
            tab_pcd!imagem =
            tab_pcd!
            tab_pcd.Update
            Call Habilitar_Mascara
End Sub
