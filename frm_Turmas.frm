VERSION 5.00
Begin VB.Form frm_Turmas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Turmas"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmd_Alterar 
      Caption         =   "Alterar"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame2 
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
      Height          =   1335
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   4935
      Begin VB.CommandButton cmdultimo 
         Caption         =   "Último"
         Height          =   375
         Left            =   3720
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdanterior 
         Caption         =   "Anterior"
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdproximo 
         Caption         =   "Próximo"
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_Primeiro 
         Caption         =   "Primeiro"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_Novo 
         Caption         =   "Novo"
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_Excluir 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Cmd_Gravar 
         Caption         =   "Gravar"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cadastro de turmas"
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
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      Begin VB.TextBox Txt_Turma 
         Height          =   405
         Left            =   4200
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Txt_Curso 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         TabIndex        =   2
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox Txt_Codigo 
         Height          =   405
         Left            =   840
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Turma:"
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Curso:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   495
      End
   End
End
Attribute VB_Name = "Frm_Turmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_turma As Integer
Private Sub cmd_alterar_Click()
            status = "Alteradas"
            Call box_1
            Call turma
End Sub

Private Sub cmd_excluir_Click()
            status = " Excluídas"
            Call Fechar
            If MsgBox("Deseja realmente excluir este Regitro", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
               tab_turma.Open "Select * From Turmas Where cod_turma = " & L_turma
               If tab_turma.RecordCount = 1 Then
                  bdvb.Execute "Delete From Turmas Where cod_turma = " & L_turma
               End If
               End If
               Call box_1
End Sub

Private Sub Fechar()
            If tab_turma.State = adStateOpen Then tab_turma.Close
End Sub

Private Sub Cmd_Gravar_Click()
            status = "Incluídas"
            If tab_turma.State = adStateOpen Then
            tab_turma.Close
            tab_turma.Open "select * from Turmas where cod_turma = '" & Txt_Codigo & "'"
            If tab_turma.RecordCount <> 0 Then
            MsgBox "Código ja cadastrado", vbCritical, "Allia"
            Exit Sub
            End If
            End If
            Call box_1
            Call turma
End Sub
Private Sub turma()
            If status <> "Alteradas" Then tab_turma.AddNew
            tab_turma!turma = Txt_Turma.Text
            tab_turma!curso = Txt_Curso.Text
            tab_turma!cod_turma = Txt_Codigo.Text
            tab_turma.Update
End Sub

Private Sub cmd_novo_Click()
            Txt_Turma = Clear
            Txt_Curso = Clear
            Txt_Codigo = Clear
            Txt_Codigo.SetFocus
End Sub

Private Sub cmd_primeiro_Click()
            Txt_Codigo = tab_turma!cod_turma
            Txt_Curso = tab_turma!curso
            Txt_Turma = tab_turma!turma
            tab_turma.MoveFirst
End Sub

Private Sub cmdanterior_Click()
            If tab_turma.BOF = True Then
            tab_turma.MoveLast
            End If
            Txt_Codigo = tab_turma!cod_turma
            Txt_Curso = tab_turma!curso
            Txt_Turma = tab_turma!turma
            tab_turma.MovePrevious
End Sub

Private Sub cmdproximo_Click()
            If tab_turma.EOF = True Then
            tab_turma.MoveFirst
            End If
            Txt_Codigo = tab_turma!cod_turma
            Txt_Curso = tab_turma!curso
            Txt_Turma = tab_turma!turma
            tab_turma.MoveNext
End Sub

Private Sub cmdultimo_Click()
            Txt_Codigo = tab_turma!cod_turma
            Txt_Curso = tab_turma!curso
            Txt_Turma = tab_turma!turma
            tab_turma.MoveLast
End Sub

Private Sub Form_Load()
'            Call OpenBD
            tab_turma.Open "Turmas", caminho, adOpenKeyset, adLockOptimistic
End Sub

