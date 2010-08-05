VERSION 5.00
Begin VB.Form frm_coli 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Coligações"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   2
      Top             =   960
      Width           =   5895
      Begin VB.CommandButton cmd_novo 
         Caption         =   "Novo"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_incluir 
         Caption         =   "Incluir"
         Height          =   495
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_excluir 
         Caption         =   "Excluir"
         Height          =   495
         Left            =   3000
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_alterar 
         Caption         =   "Alterar"
         Height          =   495
         Left            =   4440
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_primeiro 
         Caption         =   "Primeiro"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_proximo 
         Caption         =   "Próximo"
         Height          =   495
         Left            =   3000
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_anterior 
         Caption         =   "Anterior"
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_ultimo 
         Caption         =   "Último"
         Height          =   495
         Left            =   4440
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Coligações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txt_coligacao 
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "Coligação:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_coli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l_coli As Integer

Private Sub cmd_alterar_Click()
            Call box_1
            status = "Alteradas"
            If txt_coligacao = Empty Then
               MsgBox "Desculpe, mas é necessário selecionar uma informação para depois alterar-la.", vbExclamation, "Allia"
               txt_coligacao.SetFocus
            Exit Sub
            End If
            If MsgBox("Você deseja realmente alterar esta coligação?", vbQuestion + vbYesNo + vbDefaultButton2, "Allia") = vbYes Then
               If tab_coli = adStateOpen Then
                  tab_coli.Close
                  l_coli = tab_coli!cod_coli
                  bdvb.Execute "Update Coligacoes set coligacao = '" & txt_coligacao & "' where cod_coli = '" & l_coli
               End If
            End If
End Sub

Private Sub cmd_anterior_Click()
            On Error GoTo A
            'If tab_coli.BOF = True Then tab_coli.MoveLast
            tab_coli.MovePrevious
            txt_coligacao = tab_coli!coligacao
A:
End Sub

Private Sub cmd_excluir_Click()
            If txt_coligacao = Empty Then
               MsgBox "Desculpe, mas é necessário selecionar uma informação para depois excluir", vbExclamation, "Allia"
               txt_coligacao.SetFocus
            Exit Sub
            End If
            If MsgBox("Você deseja realmente excluir esta coligação?", vbQuestion + vbYesNo + vbDefaultButton2, "Allia") = vbYes Then
                  bdvb.Execute "delete from Coligacoes where coligacao = '" & txt_coligacao & "'"
                  txt_coligacao = Clear
                  txt_coligacao.SetFocus
                  Exit Sub
               End If
            
End Sub
Private Sub coli()
            If status <> "Alteradas" Then tab_coli.AddNew
            tab_coli!coligacao = txt_coligacao.Text
            tab_coli.Update
End Sub
Private Sub cmd_incluir_Click()
            Call box_1
            status = "Gravadas"
            Call coli
              'If tab_coli.State = adStateOpen Then
              'tab_coli.Close
              'tab_coli.Open "Select * from Coligacoes where coligacao = '" & txt_coligacao & "'"
              'If tab_coli.RecordCount <> 0 Then
                 'MsgBox "Coligação já cadastrada, por favor tente outra.", vbExclamation, "Allia"
              'Exit Sub
              'ElseIf tab_coli.RecordCount = 0 Then
                      'bdvb.Execute "Insert Into Coligacoes(coligacao) values ('" & txt_coligacao & "')"
              'End If
              'End If
             txt_coligacao = Clear
End Sub

Private Sub cmd_novo_Click()
            txt_coligacao = Clear
End Sub

Private Sub cmd_primeiro_Click()
            On Error GoTo A
            tab_coli.MoveFirst
            txt_coligacao = tab_coli!coligacao
A:

End Sub

Private Sub cmd_proximo_Click()
            On Error GoTo A
            If tab_coli.EOF = True Then tab_coli.MoveFirst
            tab_coli.MoveNext
            txt_coligacao = tab_coli!coligacao
A:
End Sub

Private Sub cmd_ultimo_Click()
            On Error GoTo A
            tab_coli.MoveLast
            txt_coligacao = tab_coli!coligacao
A:
End Sub

Private Sub Form_Load()
            Call OpenBD
            Call OpenUrna

End Sub
