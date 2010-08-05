VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_candidatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Candidatos"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cod_dialog 
      Left            =   8880
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   1680
      TabIndex        =   5
      Top             =   3480
      Width           =   1935
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
      Height          =   1095
      Left            =   0
      TabIndex        =   16
      Top             =   4080
      Width           =   9375
      Begin VB.CommandButton cmd_novo 
         Caption         =   "Novo"
         Height          =   495
         Left            =   5040
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_incluir 
         Caption         =   "Incluir"
         Height          =   495
         Left            =   6120
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_excluir 
         Caption         =   "Excluir"
         Height          =   495
         Left            =   7200
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_alterar 
         Caption         =   "Alterar"
         Height          =   495
         Left            =   8280
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_primeiro 
         Caption         =   "Primeiro"
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_proximo 
         Caption         =   "Próximo"
         Height          =   495
         Left            =   1200
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_anterior 
         Caption         =   "Anterior"
         Height          =   495
         Left            =   2280
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_ultimo 
         Caption         =   "Último"
         Height          =   495
         Left            =   3360
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame lbl_dadoseleitorais 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   9375
      Begin VB.ComboBox cbo_esc 
         Height          =   315
         Left            =   4800
         TabIndex        =   33
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txt_ano 
         Height          =   405
         Left            =   4560
         TabIndex        =   31
         Top             =   2880
         Width           =   1215
      End
      Begin VB.ComboBox cbo_cargo 
         Height          =   315
         ItemData        =   "frm_candidatos.frx":0000
         Left            =   1680
         List            =   "frm_candidatos.frx":000A
         TabIndex        =   30
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txt_numero 
         Height          =   375
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   29
         Top             =   2880
         Width           =   1575
      End
      Begin VB.ComboBox cbo_partido 
         Height          =   315
         Left            =   4560
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txt_imagem 
         Height          =   375
         Left            =   6600
         TabIndex        =   27
         Top             =   2640
         Width           =   2295
      End
      Begin VB.ComboBox cbo_uf 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         ItemData        =   "frm_candidatos.frx":0026
         Left            =   4560
         List            =   "frm_candidatos.frx":0028
         TabIndex        =   6
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_procurar 
         Caption         =   "Procurar"
         Height          =   375
         Left            =   7200
         TabIndex        =   15
         Top             =   3120
         Width           =   975
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
         TabIndex        =   0
         Top             =   240
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
         TabIndex        =   1
         Top             =   840
         Width           =   4215
      End
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
         TabIndex        =   4
         Top             =   2400
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker dtp_nascimento 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   85000193
         CurrentDate     =   40179
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Escolaridade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3360
         TabIndex        =   32
         Top             =   1560
         Width           =   1260
      End
      Begin VB.Label Label4 
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
         TabIndex        =   28
         Top             =   1560
         Width           =   975
      End
      Begin VB.Image img_candidato 
         Height          =   2175
         Left            =   6840
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1935
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
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3480
         Width           =   975
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
         TabIndex        =   25
         Top             =   3480
         Width           =   735
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
         Left            =   3720
         TabIndex        =   14
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lbl_numerodocandidato 
         Caption         =   "Número:"
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
         TabIndex        =   13
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label lbl_numerodopartido 
         Caption         =   "Partido:"
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
         TabIndex        =   12
         Top             =   2040
         Width           =   735
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
         TabIndex        =   11
         Top             =   360
         Width           =   1575
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
         TabIndex        =   10
         Top             =   960
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
         TabIndex        =   9
         Top             =   2520
         Width           =   615
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
         TabIndex        =   8
         Top             =   2040
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_candidatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l_codp As Integer
Dim l_partido As Integer
Dim imagem As String
Dim way As String


Private Sub cbo_partido_Click()
            
            l_partido = cbo_partido.ItemData(cbo_partido.ListIndex)
                If cbo_partido = "BRANCO" Then
                    cbo_partido = Empty
                    txt_candidato = Clear
                    MsgBox ("este não é um partido valido!selecione outro partido..."), vbExclamation, "Allia"
                    Exit Sub
                ElseIf cbo_partido = "NULO" Then
                    cbo_partido = Empty
                    txt_candidato = Clear
                    MsgBox ("este não é um partido valido!selecione outro partido..."), vbExclamation, "Allia"
                    Exit Sub
                Else
                    tab_par.Close
                    tab_par.Open "select * from Partidos where sigla ='" & cbo_partido & "'"   '* = seleção de todos os campos da entidade/ seleciona todos os campos da entidade ufs, onde o nome for igual ao que é digitado na caixinha de texto ;  '2ª instrução em SQL''
                        If tab_par.RecordCount <> 0 Then
                        txt_candidato = tab_par!legenda
                        End If
                End If
        
    
            
End Sub





Private Sub cmd_alterar_Click()
            status = "Alteradas"
            Call box_1
            Call gravar
End Sub

Private Sub cmd_excluir_Click()
              status = " Excluídas"
            If MsgBox("Deseja realmente excluir este Presidente", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
               tab_pcd.Open "Select * From Presidentes Where codigo = " & l_codp
               If tab_pcd.RecordCount = 1 Then
                  conexao.Execute "Delete From Presidentes Where codigo = " & l_codp
               End If
               Call box_1
            End If
End Sub

Private Sub cmd_incluir_Click()
            status = " incluídas"
            
            If txt_nomereal = Empty Or txt_nomepolitico = Empty Or txt_vice = Empty Or txt_candidato = Empty Then
               MsgBox "Atenção, é necessário completar todos os campos para cadastrar um presidente", vbExclamation, "Allia"
               Exit Sub
            Else
              If cbo_cargo.Text = "" Then
               MsgBox "Atenção, é necessário completar todos os campos para cadastrar um presidente", vbExclamation, "Allia"
               ElseIf cbo_partido.Text = "" Then
                MsgBox "Atenção, é necessário completar todos os campos para cadastrar um presidente", vbExclamation, "Allia"
                ElseIf txt_vice.Text = "" Then
                 MsgBox "Atenção, é necessário completar todos os campos para cadastrar um presidente", vbExclamation, "Allia"
                 ElseIf txt_numero.Text = "" Then
                  MsgBox "Atenção, é necessário completar todos os campos para cadastrar um presidente", vbExclamation, "Allia"
                  ElseIf txt_ano = "" Then
                   MsgBox "Atenção, é necessário completar todos os campos para cadastrar um presidente", vbExclamation, "Allia"
                   ElseIf txt_municipio = "" Then
                    MsgBox "Atenção, é necessário completar todos os campos para cadastrar um presidente", vbExclamation, "Allia"
                    ElseIf cbo_uf.Text = "" Then
                     MsgBox "Atenção, é necessário completar todos os campos para cadastrar um presidente", vbExclamation, "Allia"
                     End If
                 
                 
            
            
                Call box_1
End Sub

Private Sub gravar()
           
            

            
            imagem = cod_dialog.FileTitle
            If status <> "Alteradas" Then tab_pcd.AddNew
            tab_pcd!nome_completo = txt_nomereal.Text
            tab_pcd!nome_politico = txt_nomepolitico.Text
            'tab_pcd!data_nascimento = dtp_nascimento
            tab_pcd!numero_candidato = txt_candidato
            tab_pcd!vice = txt_vice.Text
            tab_pcd!imagem = txt_imagem.Text
            tab_pcd.Update
            
End Sub

Private Sub cmd_procurar_Click()
            Call Selecionar
End Sub

Private Sub Command2_Click()

End Sub



Private Sub Form_Load()
'            Call OpenBD
             Call OpenUrna
             Do While tab_par.EOF = False
                      cbo_partido.AddItem tab_par!sigla
                      cbo_partido.ItemData(cbo_partido.NewIndex) = tab_par!legenda
                      tab_par.MoveNext
            Loop
             
            
End Sub

Private Sub Selecionar()
            cod_dialog.ShowOpen
            imagem = cod_dialog.FileTitle
            txt_imagem = cod_dialog.FileTitle
            img_candidato = LoadPicture(imagem)
End Sub


Private Sub txt_candidato_LostFocus()
            'tab_par.Close
            'On Error Resume Next
            'tab_par.Open "select * from Partidos where legenda ='" & txt_candidato   '* = seleção de todos os campos da entidade/ seleciona todos os campos da entidade ufs, onde o nome for igual ao que é digitado na caixinha de texto ;  '2ª instrução em SQL''
                        'If tab_par.RecordCount <> 0 Then
                        'cbo_partido = tab_par!sigla
                        
                        'End If
End Sub
