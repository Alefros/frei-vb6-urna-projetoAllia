VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmloca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro e Controle de Localizações"
   ClientHeight    =   3315
   ClientLeft      =   4575
   ClientTop       =   3360
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
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
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   4935
      Begin VB.CommandButton cmd_alterar 
         Caption         =   "Alterar"
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
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_incluir 
         Caption         =   "Incluir"
         Enabled         =   0   'False
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
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_novo 
         Caption         =   "Novo"
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_excluir 
         BackColor       =   &H8000000B&
         Caption         =   "Excluir"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_primeiro 
         Caption         =   "Primeiro"
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
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmd_proximo 
         Caption         =   "Próximo"
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
         Left            =   2520
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmd_ultimo 
         Caption         =   "Último"
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
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmd_anterior 
         Caption         =   "Anterior"
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
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3625
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Países"
      TabPicture(0)   =   "frmloca.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txt_pais"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Estados"
      TabPicture(1)   =   "frmloca.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Combo1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   -73680
         TabIndex        =   13
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   -73680
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txt_pais 
         Alignment       =   2  'Center
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
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         Left            =   -74760
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "País"
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
         Left            =   -74880
         TabIndex        =   14
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmloca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l_uf As Integer

Dim l_cid As Integer

Dim l_brr As Integer

Dim l_coduf As Integer

Dim l_codcid As Integer

Dim l_codbrr As Integer




Private Sub cbo_cidade_Click()
            l_codcid = cbo_cidade.ItemData(cbo_cidade.ListIndex)
End Sub

Private Sub cbo_cidade_KeyPress(KeyAscii As Integer)
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
End Sub

Private Sub cbo_estado_Change()
            l_coduf = cbo_estados.ItemData(cbo_estados.ListIndex)
End Sub

Private Sub cbo_estados_Change()
            l_coduf = cbo_estados.ItemData(cbo_estados.ListIndex)
End Sub

Private Sub cbo_pais_Click()
            l_coduf = cbo_pais.ItemData(cbo_pais.ListIndex)
End Sub

Private Sub cbo_pais_KeyPress(KeyAscii As Integer)
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
End Sub

Private Sub cmd_alterar_Click()
            status = "alteradas"
            If SSTab1.Tab = 0 Then
                If txt_pais = Empty Then
                MsgBox "Não é possível alterar Nome", vbExclamation, "Projeto ALIIA"
                Exit Sub
                End If
               l_uf = tab_pais!cod_pais
               If tab_pais.State = adStateOpen Then 'determina o estado, a situação em que se encontra
                    tab_pais.Close
                    tab_pais.Open "Select * from Paises where pais = '" & txt_pais & "'" 'o asterisco serve para selecionar todos os atributos da entidade (geral)que vc vai selecionar
                    If tab_pais.RecordCount <> 0 Then 'RecordCount conta registros
                        MsgBox "Atenção! Esta unidade já foi cadastrada, por favor, verificar.", vbExclamation
                        Exit Sub
                        ElseIf tab_pais.RecordCount = 0 Then
                            tab_pais.Close
                            tab_pais.Open "Update Paises set pais = '" & txt_pais & "' where cod_pais = " & l_uf
                            
                    End If
                End If
            
            ElseIf SSTab1.Tab = 1 Then
                    l_uf = tab_regioes!cod_reg
                    If tab_regioes.State = adStateOpen Then 'determina o estado, a situação em que se encontra
                        tab_regioes.Close
                        tab_regioes.Open "Select * from Regioes where regiao = '" & txt_regiao & "'" 'o asterisco serve para selecionar todos os atributos da entidade (geral)que vc vai selecionar
                        If tab_regioes.RecordCount <> 0 Then 'RecordCount conta registros
                            MsgBox "Atenção! Esta cidade já foi cadastrada, por favor, verificar.", vbExclamation
                            Exit Sub
                            ElseIf tab_regioes.RecordCount = 0 Then
                            tab_regioes.Close
                            tab_regioes.Open "Update Regioes set regiao = '" & txt_regiao & "' where cod_reg = " & l_uf
                        End If
                    End If
                ElseIf SSTab1.Tab = 2 Then
                        l_uf = tab_estados!cod_uf
                        If tab_estados.State = adStateOpen Then 'determina o estado, a situação em que se encontra
                            tab_estados.Close
                            tab_estados.Open "Select * from Estados where Nome = '" & txt_estado & "'" 'o asterisco serve para selecionar todos os atributos da entidade (geral)que vc vai selecionar
                            If tab_estados.RecordCount <> 0 Then 'RecordCount conta registros
                                MsgBox "Atenção! Esta cidade já foi cadastrada, por favor, verificar.", vbExclamation
                                Exit Sub
                                ElseIf tab_estados.RecordCount = 0 Then
                                tab_estados.Close
                                tab_estados.Open "Update Estados set uf = '" & txt_estado & "' where cod_uf = " & l_uf
                            End If
                        End If
                       ElseIf SSTab1.Tab = 3 Then
                                l_uf = tab_municipios!cod_mun
                                If tab_municipios.State = adStateOpen Then 'determina o estado, a situação em que se encontra
                                    tab_municipios.Close
                                    tab_municipios.Open "Select * from Municipios where nome = '" & txt_municipio & "'" 'o asterisco serve para selecionar todos os atributos da entidade (geral)que vc vai selecionar
                                    If tab_municipios.RecordCount <> 0 Then 'RecordCount conta registros
                                        MsgBox "Atenção! Este Municipio já foi cadastrado, por favor, verificar.", vbExclamation
                                        Exit Sub
                                        ElseIf tab_municipios.RecordCount = 0 Then
                                            tab_municipios.Close
                                            tab_municipios.Open "Update Municipios set nome = '" & txt_municipio & "' where cod_mun = " & l_uf
                                    End If
                                End If
            End If
            Call box_1
            Call limpar
End Sub
Private Sub cmd_excluir_Click()
            status = "excluídas"
            If tab_pais.State = adStateOpen Then tab_pais.Close
            
            If SSTab1.Tab = 0 Then
            tab_pais.Open "select * from Paises where pais = '" & txt_pais & "'"
            If tab_pais.RecordCount = 0 Then
            MsgBox "Pais não existente.", vbInformation
            Exit Sub
            ElseIf tab_pais.RecordCount > 0 Then
               If MsgBox("Deseja realmente excluir este País?", vbQuestion + vbYesNo + vbDefaultButton2, "Allia") = vbYes Then  'criar uma caixa de colocar foco no segundo botão
                  bdvb.Execute "delete from Paises where pais = '" & txt_pais & "'"
                  'txt_pais.Text = Clear
               Else
               Exit Sub
               End If
               End If
               
               ElseIf SSTab1.Tab = 1 Then
                      If MsgBox("Deseja realmente excluir esta cidade?", vbQuestion + vbYesNo + vbDefaultButton2, "Allia") = vbYes Then
                         bdvb.Execute "delete from Regioes where regiao = '" & txt_regiao & "'"
                      End If
                   
                   ElseIf SSTab1.Tab = 2 Then
                          If MsgBox("Deseja realmente excluir este bairro?", vbQuestion + vbYesNo + vbDefaultButton2, "vendas") = vbYes Then
                             bdvb.Execute "delete from Estados where uf = '" & txt_estado & "'"
                          End If
                       
                       ElseIf SSTab1.Tab = 3 Then
                              If MsgBox("Deseja realmente excluir este endereço?", vbQuestion + vbYesNo + vbDefaultButton2, "vendas") = vbYes Then
                                  bdvb.Execute "delete from Municipios where nome = '" & txt_municipio & "'"
                              End If
            End If
            Call box_1
            Call limpar
End Sub
Private Sub cmd_incluir_Click()
            status = "Gravadas"
            If SSTab1.Tab = 0 Then
               If tab_pais.State = adStateOpen Then tab_pais.Close '
                    tab_pais.Open "select * from Paises where pais = '" & txt_pais & "'"
                    If tab_pais.RecordCount <> 0 Then 'RecordCount conta registros
                        MsgBox "Atenção! Esta unidade já foi cadastrada, por favor, verificar.", vbExclamation
                        txt_pais = ""
                        Exit Sub
                    ElseIf tab_pais.RecordCount = 0 Then
                        bdvb.Execute "Insert into Paises(pais) values('" & txt_pais & "')"
                    End If
                
            ElseIf SSTab1.Tab = 1 Then
                If tab_regioes.State = adStateOpen Then 'determina o estado, a situação em que se encontra
                    tab_regioes.Close 'tabela UFS já está aberta em Form_Load e para manipular o BD ela precisa ser fechada para tab_pais.Open (linha abaixo) ter efeito. "Não é possível abrir algo que já está aberto".
                    tab_regioes.Open "Select * from Regioes where regiao = '" & txt_regiao & "'" 'o asterisco serve para selecionar todos os atributos da entidade (geral)que vc vai selecionar
                    If tab_regioes.RecordCount <> 0 Then 'RecordCount conta registros
                        MsgBox "Atenção! Esta cidade já foi cadastrada, por favor, verificar.", vbExclamation
                        Exit Sub
                    ElseIf tab_regioes.RecordCount = 0 Then
                        bdvb.Execute "Insert into Regioes(regiao) values('" & txt_regiao & "')"
                    End If
                End If
                ElseIf SSTab1.Tab = 2 Then
                    If tab_estados.State = adStateOpen Then 'determina o estado, a situação em que se encontra
                        tab_estados.Close 'tabela UFS já está aberta em Form_Load e para manipular o BD ela precisa ser fechada para tab_pais.Open (linha abaixo) ter efeito. "Não é possível abrir algo que já está aberto".
                        tab_estados.Open "Select * from Estados where uf = '" & txt_estado & "'" 'o asterisco serve para selecionar todos os atributos da entidade (geral)que vc vai selecionar
                        If tab_estados.RecordCount <> 0 Then 'RecordCount conta registros
                            MsgBox "Atenção! Este Estado já foi cadastrado, por favor, verificar.", vbExclamation
                        Exit Sub
                        ElseIf tab_estados.RecordCount = 0 Then
                            bdvb.Execute "Insert into Estados(uf) values('" & txt_estado & "')"
                        End If
                    End If
                    ElseIf SSTab1.Tab = 3 Then
                           If tab_municipios.State = adStateOpen Then 'determina o estado, a situação em que se encontra
                                tab_municipios.Close 'tabela UFS já está aberta em Form_Load e para manipular o BD ela precisa ser fechada para tab_pais.Open (linha abaixo) ter efeito. "Não é possível abrir algo que já está aberto".
                                tab_municipios.Open "Select * from Municipios where nome = '" & txt_municipio & "'" 'o asterisco serve para selecionar todos os atributos da entidade (geral)que vc vai selecionar
                                If tab_municipios.RecordCount <> 0 Then 'RecordCount conta registros
                                    MsgBox "Atenção! Este Município já foi cadastrado, por favor, verificar.", vbExclamation
                                    Exit Sub
                                ElseIf tab_municipios.RecordCount = 0 Then
                                    bdvb.Execute "Insert into Municipios (nome) values('" & txt_municipio & "')"
                           End If
                    End If
            End If
            Call box_1
            Call limpar
            'md_incluir.Enabled = False
End Sub
Private Sub cmd_novo_Click()
            Call limpar
End Sub
Private Sub limpar()
            If SSTab1.Tab = 0 Then 'refere-se a aba Paises
                txt_pais = Clear
                txt_pais.SetFocus
            ElseIf SSTab1.Tab = 1 Then 'refere-se a aba Cidades
                    txt_regiao = Clear
                    txt_regiao.Text = Clear
                ElseIf SSTab1.Tab = 2 Then 'refere-se a aba Bairros
                        txt_bairro = Clear
                        cbo_cidade = Clear
                    ElseIf SSTab1.Tab = 3 Then 'refere-se a aba Localizacoes
                            
                          
            End If
End Sub
Private Sub cmd_primeiro_Click()
           On Error Resume Next
            If SSTab1.Tab = 0 Then
            tab_pais.MoveFirst
            txt_pais = tab_pais!pais

            ElseIf SSTab1.Tab = 1 Then 'refere-se a aba Cidades
                    tab_regioes.MoveFirst
                    txt_regiao = tab_regioes!regiao
                    l_uf = tab_regioes!cod_reg 'codigo vai pra variavel l_uf
                    If tab_pais.State = adStateOpen Then
                        tab_pais.Close
                        tab_pais.Open "Select * from UFS where Codigo = " & l_uf
                        If tab_pais.RecordCount <> 0 Then
                            cbo_pais.Text = tab_pais!nome 'transforma o codigo em letras
                        End If
                    End If
                
                ElseIf SSTab1.Tab = 2 Then 'refere-se a aba Bairros
                        tab_estados.MoveFirst
                        txt_bairro = tab_estados!nome
                        l_cid = tab_estados!cod_cid 'codigo vai pra variavel l_uf
                        If tab_regioes.State = adStateOpen Then
                            tab_regioes.Close
                            tab_regioes.Open "Select * from Cidades where Codigo = " & l_cid
                            If tab_regioes.RecordCount <> 0 Then
                                cbo_cidade.Text = tab_regioes!nome 'transforma o codigo em letras
                            End If
                        End If
                    
                    ElseIf SSTab1.Tab = 3 Then 'refere-se a aba Localizacoes
                            tab_municipios.MoveFirst
                            txt_municipio = tab_municipios!nome
                            
            End If
            
            
End Sub
Private Sub cmd_proximo_Click()
            On Error Resume Next
            If SSTab1.Tab = 0 Then
                If tab_pais.EOF = True Then tab_pais.MoveFirst
                tab_pais.MoveNext
                txt_pais = tab_pais!pais

            ElseIf SSTab1.Tab = 1 Then
                    tab_regioes.MoveNext
                    If tab_regioes.EOF = True Then
                       tab_regioes.MoveFirst
                    End If
                    txt_regiao = tab_regioes!regiao
                    l_uf = tab_regioes!cod_uf
                    If tab_pais.State = adStateOpen Then
                        tab_pais.Close
                        tab_pais.Open "Select * from Estados where Codigo = " & l_uf
                        If tab_pais.RecordCount <> 0 Then
                            cbo_pais.Text = tab_pais!nome
                        End If
                    End If
                
                ElseIf SSTab1.Tab = 2 Then '
                       tab_estados.MoveNext
                       If tab_estados.EOF = True Then
                          tab_estados.MoveFirst
                       End If
                       txt_bairro = tab_estados!nome
                        
                        ElseIf SSTab1.Tab = 3 Then 'refere-se a aba Localizacoes
                           tab_municipios.MoveNext
                           If tab_municipios.EOF = True Then
                              tab_municipios.MoveFirst
                           End If
                           txt_logr = tab_municipios!logradouro
            End If
            
End Sub
Private Sub cmd_anterior_Click()
            On Error Resume Next
            If SSTab1.Tab = 0 Then
                If tab_pais.BOF = True Then tab_pais.MoveLast
                tab_pais.MovePrevious
                txt_pais = tab_pais!pais
                

            ElseIf SSTab1.Tab = 1 Then 'refere-se a aba Cidades
                    tab_regioes.MovePrevious
                    If tab_regioes.BOF = True Then
                        tab_regioes.MoveLast
                    End If
                    txt_regiao = tab_regioes!regiao
                    l_uf = tab_regioes!cod_reg 'codigo vai pra variavel l_uf
                    If tab_pais.State = adStateOpen Then
                        tab_pais.Close
                        tab_pais.Open "Select * from Regioes where cod_reg = " & l_uf
                        If tab_pais.RecordCount <> 0 Then
                            'cbo_pais.Text = tab_pais!nome 'transforma o codigo em letras
                        End If
                    End If
                ElseIf SSTab1.Tab = 2 Then 'refere-se a aba Bairros
                        tab_estados.MovePrevious
                        If tab_estados.BOF = True Then
                           tab_estados.MoveLast
                        End If
                        txt_bairro = tab_estados!nome
                        l_cid = tab_estados!cod_cid 'codigo vai pra variavel l_cid
                            If tab_regioes.State = adStateOpen Then
                            tab_regioes.Close
                            tab_regioes.Open "Select * from Cidades where Codigo = " & l_cid
                                If tab_regioes.RecordCount <> 0 Then
                                    cbo_cidade.Text = tab_regioes!nome 'transforma o codigo em letras
                                End If
                        End If
                    ElseIf SSTab1.Tab = 3 Then 'refere-se a aba Localizacoes
                            tab_municipios.MovePrevious
                            If tab_municipios.BOF = True Then
                               tab_municipios.MoveLast
                            End If
                            txt_municipio = tab_municipios!nome
            End If
End Sub
Private Sub cmd_ultimo_Click()
            On Error Resume Next
            If SSTab1.Tab = 0 Then
            tab_pais.MoveLast
            txt_pais = tab_pais!pais

            ElseIf SSTab1.Tab = 1 Then 'refere-se a aba Cidades
                    tab_regioes.MoveLast
                    txt_regiao = tab_regioes!regiao
                    l_uf = tab_regioes!cod_reg 'codigo vai pra variavel l_uf
                    If tab_pais.State = adStateOpen Then
                        tab_pais.Close
                        tab_pais.Open "Select * from UFS where Codigo = " & l_uf
                        If tab_pais.RecordCount <> 0 Then
                            cbo_pais.Text = tab_pais!nome 'transforma o codigo em letras
                        End If
                    End If
                
                ElseIf SSTab1.Tab = 2 Then 'refere-se a aba Bairros
                        tab_estados.MoveLast
                        txt_estado = tab_estados!nome
                    
                    ElseIf SSTab1.Tab = 3 Then 'refere-se a aba Localizacoes
                             tab_municipios.MoveLast
                             txt_logr = tab_municipios!logradouro
            End If
            
            
End Sub
Private Sub Form_Load()
           ' Call OpenBD
             Call OpenUrna
                Do While tab_pais.EOF = False
                         cbo_pais.AddItem tab_pais!pais
                         cbo_pais.ItemData(cbo_pais.NewIndex) = tab_pais!cod_pais 'faz com que a combo_box salve no BD o codigo da UF e não a própria UF
                         tab_pais.MoveNext
                Loop

End Sub






Private Sub Label6_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub txt_pais_Change()
            cmd_incluir.Enabled = True
End Sub

Private Sub txt_pais_KeyPress(KeyAscii As Integer)
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
Private Sub txt_pais_LostFocus()
            txt_pais = UCase(txt_pais) 'na txtbox pode aparecer letras minusculas, mas no BD só cadastra maiúscula
End Sub
