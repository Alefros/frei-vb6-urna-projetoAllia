VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_releitores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro e Controle de Eleitores"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Cadastro e Controle de Eleitores"
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
      TabIndex        =   5
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmd_procurar 
         Caption         =   "Procurar"
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txt_nome 
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
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   4455
      End
      Begin MSMask.MaskEdBox msk_rg 
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   840
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "99.999.999-9"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Nome:"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "RG:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comandos"
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
      Top             =   1560
      Width           =   5415
      Begin VB.CommandButton cmd_votar 
         Caption         =   "VOTAR"
         Height          =   495
         Left            =   1920
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmd_alterar 
         Caption         =   "Alterar"
         Height          =   495
         Left            =   4080
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_excluir 
         Caption         =   "Excluir"
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_incluir 
         Caption         =   "Incluir"
         Height          =   495
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_novo 
         Caption         =   "Novo"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_releitores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_alterar_Click()
           status = " alteradas"
           msk_rg.PromptInclude = False
           If txt_nome = Empty Then
               MsgBox "Não é possível alterar uma informação em branco, por favor procure seus dados pelo RG.", vbExclamation, "Allia"
                Exit Sub
            End If
           If msk_rg = Empty Then
               MsgBox "Não é possível alterar uma informação em branco, por favor procure seus dados pelo RG.", vbExclamation, "Allia"
                Exit Sub
            End If
            If tab_ele.State = adStateOpen Then tab_ele.Close
            
            tab_ele.Open "select * from Alunos where RG = '" & msk_rg & "'"
            If tab_ele.RecordCount <> 0 Then
            tab_ele.Update
            tab_ele!nome = txt_nome
            tab_ele!RG = msk_rg
            tab_ele.Update
            Else
            MsgBox "Eleitor não cadastrado.", vbExclamation, "Allia"
            Exit Sub
            End If
            Call box_1
            msk_rg.PromptInclude = True
            Call cmd_novo_Click
End Sub

Private Sub cmd_excluir_Click()
           status = " excluidas"
           msk_rg.PromptInclude = False
           If txt_nome = Empty Then
               MsgBox "Não é possível excluir uma informação em branco, por favor procure seus dados pelo RG.", vbExclamation, "Allia"
                Exit Sub
            End If
           If msk_rg = Empty Then
               MsgBox "Não é possível excluir uma informação em branco, por favor procure seus dados pelo RG.", vbExclamation, "Allia"
                Exit Sub
            End If
            
            If tab_ele.State = adStateOpen Then tab_ele.Close
            tab_ele.Open "select * from Alunos where RG = '" & msk_rg & "'"
            If tab_ele.RecordCount <> 0 Then
            If MsgBox("Deseja realmente excluir este eleitor?", vbQuestion + vbYesNo) = vbYes Then
            bdvb.Execute "delete * from Alunos where RG = '" & msg_rg & "'"
            Call cmd_novo_Click
            End If
            Else
            MsgBox "Eleitor não cadastrado.", vbExclamation, "Allia"
            Exit Sub
            End If
            Call box_1
            msk_rg.PromptInclude = True
            Call cmd_novo_Click
        
End Sub

Private Sub cmd_incluir_Click()
            status = " incluidas"
            msk_rg.PromptInclude = False
            If txt_nome = Empty And msk_rg = Empty Then
               MsgBox "Não é possível incluir uma informação em branco, por favor preencha os campos necessários.", vbExclamation, "Allia"
               Exit Sub
            End If
            If tab_ele.State = adStateOpen Then tab_ele.Close
            
               tab_ele.Open "Select * From Alunos where Nome = '" & txt_nome & "'"
               If tab_ele.RecordCount <> 0 Then
                  MsgBox "Aluno já cadastrado, por favor cadastre um aluno com outro nome.", vbExclamation, "Allia"
                  txt_nome = Clear
                  txt_nome.SetFocus
               Exit Sub
               ElseIf tab_ele.RecordCount = 0 Then
                      bdvb.Execute "Insert into Alunos(Nome,RG) values('" & txt_nome & "','" & msk_rg & "')"
               End If
               msk_rg.PromptInclude = True
               Call box_1
               Call cmd_novo_Click
End Sub

Private Sub cmd_novo_Click()
            msk_rg.PromptInclude = False
            msk_rg = Empty
            txt_nome = Empty
            txt_nome.SetFocus
            msk_rg.PromptInclude = True
End Sub

Private Sub cmd_procurar_Click()
            
            msk_rg.PromptInclude = False
            
            If msk_rg = Empty Then MsgBox "Não é possível encontrar uma informação em branco, por favor insira uma informação.", vbExclamation, "Allia"
            
            If tab_ele.State = adStateOpen Then tab_ele.Close
            
            tab_ele.Open "select * from Alunos where RG = '" & msk_rg & "'"
            
            If tab_ele.RecordCount <> 0 Then
            txt_nome = tab_ele!nome
            msk_rg = tab_ele!RG
            Else
            MsgBox "Eleitor não cadastrado.", vbExclamation, "Allia"
            Exit Sub
            End If
            msk_rg.PromptInclude = True
            
End Sub


Private Sub cmd_votar_Click()
            
            msk_rg.PromptInclude = False
            If tab_ele.State = adStateOpen Then tab_ele.Close
            tab_ele.Open "select * from Alunos where RG = '" & msk_rg & "'"
            If tab_ele.RecordCount <> 0 Then
            If tab_ele!votou = "0" Then
            tab_ele.Update
            tab_ele!votou = "1"
            tab_ele.Update
            Call cmd_novo_Click
            If tab_ele.RecordCount <> 0 Then
            frm_gvoto.Show
            Exit Sub
            ElseIf tab_ele.RecordCount = 0 Then
            MsgBox "Eleitor não cadastrado.", vbExclamation, "Allia"
            Exit Sub
            End If
            ElseIf tab_ele!votou = "1" Then
            MsgBox "Eleitor já votou.", vbExclamation, "Allia"
            Exit Sub
            End If
            End If
            
End Sub

Private Sub Form_Load()
            Call OpenUrna
End Sub

