VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_partidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Partidos"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid msflex 
      Height          =   2415
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   3
      FormatString    =   ""
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
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   6495
      Begin VB.CommandButton cmd_alterar 
         Caption         =   "Alterar"
         Height          =   495
         Left            =   4440
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_excluir 
         Caption         =   "Excluir"
         Height          =   495
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_incluir 
         Caption         =   "Incluir"
         Height          =   495
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_novo 
         Caption         =   "Novo"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Partidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txt_sigla 
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
         Left            =   1200
         TabIndex        =   12
         Top             =   720
         Width           =   1095
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
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txt_legenda 
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
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Sigla"
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
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Nome"
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
         Left            =   2040
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Legenda"
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
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_partidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim L_Colunas, L_Linha As Long
Dim L_Buscar As Integer
Dim L_CodColi As Integer
Dim l_coli As Integer
Dim L_par As Integer

Private Sub alterar()
            tab_par.Update
            tab_par!legenda = txt_legenda
            tab_par!sigla = txt_sigla
            tab_par!Partido = txt_nome
            tab_par.Update
End Sub

Private Sub cmd_alterar_Click()
            On Error GoTo penis2:
            status = " alteradas"
            Call Fechar
             tab_par.Open "Select * From Partidos where Partido = '" & txt_nome & "'"
               
              If tab_par.RecordCount <> 0 Then
                  tab_pcd.Open "select * from Presidentes where legenda = " & txt_legenda
                  
                  If tab_pcd.RecordCount = 0 Then
                  tab_gcd.Open "select * from Governadores where legenda = " & txt_legenda
                  If tab_gcd.RecordCount = 0 Then
                        
            Call alterar
            Call box_1
            Call carrega_lista
            Exit Sub
            End If
            End If
            End If
penis2:
            MsgBox "Favor excluir ou alterar os candidatos deste partido antes de excluir o mesmo.", vbExclamation, "Allia"
            
End Sub

Private Sub cmd_excluir_Click()
            
            On Error GoTo penis3
            status = " Excluídas"
            Call Fechar
            If MsgBox("Deseja realmente excluir estas informações?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
               tab_par.Open "Select * From Partidos where Partido = '" & txt_nome & "'"
               
               If tab_par.RecordCount <> 0 Then
                  tab_pcd.Open "select * from Presidentes where legenda = " & txt_legenda
                  
                  If tab_pcd.RecordCount = 0 Then
                  tab_gcd.Open "select * from Governadores where legenda = " & txt_legenda
                  If tab_gcd.RecordCount = 0 Then
                  bdvb.Execute "Delete From Partidos where Partido = '" & txt_nome & "'"
                  Call box_1
                  Call carrega_lista
                  Call Novo
                    Exit Sub
                    End If
                    End If
                    End If
                    End If

penis3:
             MsgBox "Favor excluir ou alterar os candidatos deste partido antes de excluir o mesmo.", vbExclamation, "Allia"

End Sub
Private Sub Fechar()
            If tab_par.State = adStateOpen Then tab_par.Close
            If tab_pcd.State = adStateOpen Then tab_pcd.Close
            If tab_gcd.State = adStateOpen Then tab_gcd.Close
End Sub

Private Sub cmd_incluir_Click()
            Call Partido
            Call carrega_lista
End Sub

Private Sub cmd_novo_Click()
            Call Novo
End Sub
Private Sub carrega_lista()
            If tab_par.State = adStateOpen Then tab_par.Close
               tab_par.Open "Partidos", bdvb, adOpenKeyset, adLockOptimistic
               If tab_par.BOF = False Or tab_par.EOF = False Then
                  tab_par.MoveFirst
                  msflex.Rows = 2
                  msflex.Clear
                  msflex.FormatString = "LEGENDA   | SIGLA    | PARTIDO                                                       "
                  Do Until tab_par.EOF
                     msflex.TextMatrix(msflex.Rows - 1, 0) = tab_par!legenda
                     msflex.TextMatrix(msflex.Rows - 1, 1) = tab_par!sigla
                     msflex.TextMatrix(msflex.Rows - 1, 2) = tab_par!Partido
                     msflex.Rows = msflex.Rows + 1
                     tab_par.MoveNext
               Loop
               msflex.Rows = msflex.Rows - 1
               Else
               msflex.Rows = 2
               msflex.Clear
               msflex.FormatString = "LEGENDA   | SIGLA      | PARTIDO                                                      "
            End If
End Sub
Private Sub Form_Load()
            Call OpenBD
            Call OpenUrna
            Call carrega_lista

End Sub

Private Sub Partido()
            If tab_par.State = adStateOpen Then tab_par.Close
            tab_par.Open "select * from Partidos where legenda = " & txt_legenda
            If tab_par.RecordCount = 0 Then
            tab_par.AddNew
            tab_par!Partido = txt_nome.Text
            tab_par!legenda = txt_legenda.Text
            tab_par!sigla = txt_sigla.Text
            tab_par.Update
            Call Novo
            Exit Sub
            End If
            MsgBox "Partido já cadastrado.", vbExclamation, "Allia"
End Sub

Private Sub Novo()
            txt_nome.Text = Clear
            txt_legenda.Text = Clear
            txt_sigla.Text = Clear
            txt_nome.SetFocus
            
End Sub

Private Sub msflex_Click()
            On Error GoTo Penis
            L_Linha = msflex.Row
            l_CodCli = msflex.TextMatrix(L_Linha, 0)
            If tab_par.State = adStateOpen Then tab_par.Close
            tab_par.Open "Select * From Partidos Where legenda = " & l_CodCli
            Call Mostrar
Penis:
End Sub
Private Sub Mostrar()
            txt_nome.Text = tab_par!Partido
            txt_legenda.Text = tab_par!legenda
            txt_sigla.Text = tab_par!sigla
          
End Sub
