VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_registros 
   Caption         =   "Registro"
   ClientHeight    =   2115
   ClientLeft      =   2010
   ClientTop       =   2460
   ClientWidth     =   3315
   LinkTopic       =   "Form2"
   ScaleHeight     =   2115
   ScaleWidth      =   3315
   Begin TabDlg.SSTab SSTab1 
      Height          =   1335
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2355
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "PARTIDOS"
      TabPicture(0)   =   "Registro.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_nome_partido"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_chapa_partido"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_descricao_partido"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt_nome_partido"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txt_chapa_partido"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txt_descricao_partido"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "POLITICOS"
      TabPicture(1)   =   "Registro.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_partido_candidato"
      Tab(1).Control(1)=   "lbl_chapa_candidato"
      Tab(1).Control(2)=   "lbl_nome_candidato"
      Tab(1).Control(3)=   "cbo_partido"
      Tab(1).Control(4)=   "txt_chapa_candidato"
      Tab(1).Control(5)=   "txt_nome_candidato"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txt_nome_candidato 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74280
         TabIndex        =   17
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txt_chapa_candidato 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72480
         MaxLength       =   5
         TabIndex        =   16
         Top             =   840
         Width           =   615
      End
      Begin VB.ComboBox cbo_partido 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74280
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txt_descricao_partido 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txt_chapa_partido 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txt_nome_partido 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   600
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lbl_nome_candidato 
         Appearance      =   0  'Flat
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lbl_chapa_candidato 
         Appearance      =   0  'Flat
         Caption         =   "Chapa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73080
         TabIndex        =   19
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl_partido_candidato 
         Appearance      =   0  'Flat
         Caption         =   "Partido"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl_descricao_partido 
         Appearance      =   0  'Flat
         Caption         =   "Descrição"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lbl_chapa_partido 
         Appearance      =   0  'Flat
         Caption         =   "Chapa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lbl_nome_partido 
         Appearance      =   0  'Flat
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.CommandButton cmd_ultimo 
      Caption         =   "Último"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmd_proximo 
      Caption         =   "Próximo"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmd_anterior 
      Caption         =   "Anterior"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmd_primeiro 
      Caption         =   "Primeiro"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmd_excluir 
      Caption         =   "Excluir"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmd_alterar 
      Caption         =   "Alterar"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmd_incluir 
      Caption         =   "Incluir"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmd_novo 
      Caption         =   "Novo"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   735
   End
   Begin VB.Menu mnu_modo 
      Caption         =   "Modo"
      Begin VB.Menu mnu_inicio 
         Caption         =   "Inicio"
         Index           =   1
      End
      Begin VB.Menu mnu_registro 
         Caption         =   "Registro"
         Index           =   2
      End
      Begin VB.Menu mnu_voto 
         Caption         =   "Voto"
         Index           =   3
      End
   End
   Begin VB.Menu mnu_sair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "frm_registros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l_codptd As Integer
Private Sub cmd_primeiro_Click()
            If SSTab1.Tab = 0 Then
                tab_ptd.MoveFirst
                txt_nome_partido = tab_ptd!nome
                txt_chapa_partido = tab_ptd!chapa_partido
                txt_descricao_partido = tab_ptd!Descricao
                ElseIf SSTab1.Tab = 1 Then
                    tab_cdt.MoveFirst
                    txt_nome_candidato = tab_cdt!nome
                    cbo_partido = tab_cdt!chapa_partido
                    txt_chapa_candidato = tab_cdt!Chapa_Candidato
            End If
End Sub
Private Sub cmd_anterior_Click()
            If SSTab1.Tab = 0 Then
                tab_ptd.MovePrevious
                If tab_ptd.BOF = True Then
                    tab_ptd.MoveLast
                End If
                txt_nome_partido = tab_ptd!nome
                txt_chapa_partido = tab_ptd!chapa_partido
                txt_descricao_partido = tab_ptd!Descricao
                ElseIf SSTab1.Tab = 1 Then
                    tab_cdt.MovePrevious
                    If tab_cdt.BOF = True Then
                        tab_cdt.MoveLast
                    End If
                    txt_nome_candidato = tab_cdt!nome
                    cbo_partido = tab_cdt!chapa_partido
                    txt_chapa_candidato = tab_cdt!Chapa_Candidato
            End If
End Sub
Private Sub cmd_proximo_Click()
            If SSTab1.Tab = 0 Then
                tab_ptd.MoveNext
                If tab_ptd.EOF = True Then
                    tab_ptd.MoveFirst
                End If
                txt_nome_partido = tab_ptd!nome
                txt_chapa_partido = tab_ptd!chapa_partido
                txt_descricao_partido = tab_ptd!Descricao
                ElseIf SSTab1.Tab = 1 Then
                    tab_cdt.MoveNext
                    If tab_cdt.EOF = True Then
                        tab_cdt.MoveFirst
                    End If
                    txt_nome_candidato = tab_cdt!nome
                    cbo_partido = tab_cdt!chapa_partido
                    txt_chapa_candidato = tab_cdt!Chapa_Candidato
            End If
End Sub
Private Sub cmd_ultimo_Click()
            If SSTab1.Tab = 0 Then
                tab_ptd.MoveLast
                txt_nome_partido = tab_ptd!nome
                txt_chapa_partido = tab_ptd!chapa_partido
                txt_descricao_partido = tab_ptd!Descricao
                ElseIf SSTab1.Tab = 1 Then
                    tab_cdt.MoveLast
                    txt_nome_candidato = tab_cdt!nome
                    cbo_partido = tab_cdt!chapa_partido
                    txt_chapa_candidato = tab_cdt!Chapa_Candidato
            End If
End Sub
Private Sub cmd_novo_Click()
            Call limpar
End Sub
Private Sub cmd_incluir_Click()
            Status = "incluídas"
            If SSTab1.Tab = 0 Then
                If tab_ptd.State = adStateOpen Then 'determina o estado, a situação em que se encontra
                    tab_ptd.Close 'tabela UFS já está aberta em Form_Load e para manipular o BD ela precisa ser fechada para tab_ufs.Open (linha abaixo) ter efeito. "Não é possível abrir algo que já está aberto".
                    tab_ptd.Open "Select * from Partidos where Nome = '" & txt_nome_partido & "'" 'o asterisco serve para selecionar todos os atributos da entidade (geral)que vc vai selecionar
                    If tab_ptd.RecordCount <> 0 Then 'RecordCount conta registros
                        MsgBox "Atenção! Este partido já foi cadastrado, por favor, verificar.", vbExclamation
                        Exit Sub
                        ElseIf tab_ptd.RecordCount = 0 Then
                        bdvb.Execute "Insert into Partidos(Chapa_Partido, Nome, Descricao) values('" & txt_chapa_partido & "', '" & txt_nome_partido & "', '" & txt_descricao_partido & "')"
                    End If
                End If
                ElseIf SSTab1.Tab = 1 Then
                    If tab_cdt.State = adStateOpen Then 'determina o estado, a situação em que se encontra
                        tab_cdt.Close 'tabela UFS já está aberta em Form_Load e para manipular o BD ela precisa ser fechada para tab_ufs.Open (linha abaixo) ter efeito. "Não é possível abrir algo que já está aberto".
                        tab_cdt.Open "Select * from Politicos where Nome = '" & txt_nome_candidato & "'" 'o asterisco serve para selecionar todos os atributos da entidade (geral)que vc vai selecionar
                        If tab_cdt.RecordCount <> 0 Then 'RecordCount conta registros
                            MsgBox "Atenção! Este politico já foi cadastrado, por favor, verificar.", vbExclamation
                            Exit Sub
                            ElseIf tab_cdt.RecordCount = 0 Then
                            bdvb.Execute "Insert into Politicos(Chapa_Candidato, Nome, Chapa_Partido) values('" & txt_chapa_candidato & "', '" & txt_nome_candidato & "', '" & l_codptd & "')"
                        End If
                    End If
            End If
            Call box_1
            Call limpar
End Sub
Private Sub Form_Load()
            Call OpenBD
            tab_ptd.Open "Partidos", bdvb, adOpenKeyset, adLockOptimistic 'OpenKeyset abre uma janela de ferramentas
            tab_cdt.Open "Politicos", bdvb, adOpenKeyset, adLockOptimistic
            
            Do While tab_ptd.EOF = False
                     cbo_partido.AddItem tab_ptd!nome
                     cbo_partido.ItemData(cbo_partido.NewIndex) = tab_ptd!chapa_partido 'faz com que a combo_box salve no BD o codigo da UF e não a própria UF
                     tab_ptd.MoveNext
            Loop

End Sub

Private Sub mnu_inicio_Click(Index As Integer)
            frm_inicio.Show
            frm_registros.Hide
            frm_voto.Hide
End Sub
Private Sub mnu_registro_Click(Index As Integer)
            frm_registros.Show
            frm_inicio.Hide
            frm_voto.Hide
End Sub
Private Sub mnu_voto_Click(Index As Integer)
            frm_inicio.Hide
            frm_voto.Show
            frm_registros.Hide
End Sub
Private Sub mnu_sair_Click()
            End
End Sub

Private Sub txt_nome_partido_Keypress(KeyAscii As Integer)
            KeyAscii = Asc(UCase$(Chr$(KeyAscii))) ' para letras maiusculas
            Select Case KeyAscii
                Case vbKeyDelete
                Case vbKeyBack
                Case 65 To 90 'somente letras
                Case Else
                    Beep
                    KeyAscii = 0
            End Select
            If KeyAscii = 46 Then 'não aparecer ponto (.)
                KeyAscii = 0
            End If
End Sub
Private Sub txt_chapa_partido_Keypress(KeyAscii As Integer)
            Select Case KeyAscii
                Case vbKeyDelete
                Case vbKeyBack
                Case 48 To 57 'somente numeros
                Case Else
                    Beep
                    KeyAscii = 0
            End Select
            If KeyAscii = 46 Then 'não aparecer ponto (.)
                KeyAscii = 0
            End If
End Sub
Private Sub txt_descricao_partido_KeyPress(KeyAscii As Integer)
            KeyAscii = Asc(UCase$(Chr$(KeyAscii))) ' para letras maiusculas
            Select Case KeyAscii
                Case vbKeyDelete
                Case vbKeyBack
                Case 32
                Case 65 To 90 'somente letras
                Case Else
                    Beep
                    KeyAscii = 0
            End Select
            If KeyAscii = 46 Then 'não aparecer ponto (.)
                KeyAscii = 0
            End If
End Sub
Private Sub txt_nome_candidato_KeyPress(KeyAscii As Integer)
            KeyAscii = Asc(UCase$(Chr$(KeyAscii))) ' para letras maiusculas
            Select Case KeyAscii
                Case vbKeyDelete
                Case vbKeyBack
                Case 32
                Case 65 To 90 'somente letras
                Case Else
                    Beep
                    KeyAscii = 0
            End Select
            If KeyAscii = 46 Then 'não aparecer ponto (.)
                KeyAscii = 0
            End If
End Sub
Private Sub cbo_partido_KeyPress(KeyAscii As Integer)
            KeyAscii = 0
End Sub
Private Sub cbo_partido_Click()
            l_codptd = cbo_partido.ItemData(cbo_partido.ListIndex)
End Sub
Private Sub txt_chapa_candidato_Keypress(KeyAscii As Integer)
            Select Case KeyAscii
                Case vbKeyDelete
                Case vbKeyBack
                Case 48 To 57 'somente numeros
                Case Else
                    Beep
                    KeyAscii = 0
            End Select
            If KeyAscii = 46 Then 'não aparecer ponto (.)
                KeyAscii = 0
            End If
End Sub
Private Sub limpar()
                txt_nome_partido = Clear
                txt_chapa_partido = Clear
                txt_descricao_partido = Clear
                txt_nome_candidato = Clear
                cbo_partido = Clear
                txt_chapa_candidato = Clear
End Sub
