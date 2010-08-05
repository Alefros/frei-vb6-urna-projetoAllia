VERSION 5.00
Begin VB.Form frm_gvoto 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   ClientHeight    =   11190
   ClientLeft      =   855
   ClientTop       =   435
   ClientWidth     =   15105
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11190
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_ladrao 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   0
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmd_branco 
      BackColor       =   &H8000000E&
      Caption         =   "BRANCO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      MaskColor       =   &H8000000E&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton cmd_corrigir 
      BackColor       =   &H000080FF&
      Caption         =   "CORRIGE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   5160
      MaskColor       =   &H8000000E&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton cmd_confirmar 
      BackColor       =   &H00008000&
      Caption         =   "CONFIRMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7200
      MaskColor       =   &H8000000E&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label lbl_outro 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   4320
      Width           =   5775
   End
   Begin VB.Label lbl_laranja 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2280
      TabIndex        =   15
      Top             =   8040
      Width           =   5505
   End
   Begin VB.Label lbl_verde 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2280
      TabIndex        =   14
      Top             =   7440
      Width           =   6105
   End
   Begin VB.Label lbl_tecla 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   240
      TabIndex        =   13
      Top             =   6960
      Width           =   4545
   End
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   0
      Top             =   6720
      Width           =   14775
   End
   Begin VB.Label lbl_partido2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label lbl_partido 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label lbl_vicegovernador 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Label lbl_vice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label lbl_candidato 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Label lbl_governador 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   2265
   End
   Begin VB.Label lbl_numero 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lbl_mestre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "GOVERNADOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label lbl_seuvotopara 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   3135
   End
   Begin VB.Image img_ladrao 
      Height          =   4455
      Left            =   8520
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4215
   End
End
Attribute VB_Name = "frm_gvoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim way As String

Dim img As String

Dim gvoto As Integer


Private Sub Labels()
            img = tab_gcd!imagem
            img_ladrao.Picture = LoadPicture(way & img)
            lbl_seuvotopara = "SEU VOTO PARA"
            lbl_numero = "Número:"
            lbl_governador = "Governador:"
            lbl_vice = "Vice-governador:"
            lbl_partido = "Partido:"
            
End Sub
Private Sub labels2()
            lbl_tecla = "Aperte a tecla:"
            lbl_verde = "VERDE para CONFIRMAR"
            lbl_laranja = "LARANJA para CORRIGIR"
End Sub
Private Sub limpar()
               
            If txt_ladrao = "-1" Or "-2" Then
            lbl_seuvotopara.Visible = True
            lbl_numero.Visible = True
            lbl_governador.Visible = True
            lbl_candidato.Visible = True
            lbl_vice.Visible = True
            lbl_partido.Visible = True
            lbl_partido2.Visible = True
            txt_ladrao.Visible = True
               End If
               If tab_gcd.State = adStateOpen Then
               tab_gcd.Close
               End If
               txt_ladrao = Clear
           
                img_ladrao.Picture = LoadPicture(Empty)
                lbl_seuvotopara = Clear
                lbl_numero = Clear
                lbl_governador = Clear
                lbl_candidato = Clear
                lbl_vice = Clear
                lbl_vicegovernador = Clear
                lbl_partido = Clear
                lbl_partido2 = Clear
                lbl_tecla = Clear
                lbl_verde = Clear
                lbl_laranja = Clear
                txt_ladrao.SetFocus
                lbl_outro = Clear
                If tab_gcd.State = adStateClose Then
                   tab_gcd.Open
                End If

            
                        
End Sub

Private Sub cmd_branco_Click()
                      
             If txt_ladrao = "" Then
              GoTo A
               Else
               Exit Sub
               End If
A:
            lbl_outro.Visible = True
            lbl_outro = "VOTO EM BRANCO"
            lbl_seuvotopara.Visible = False
            lbl_numero.Visible = False
            lbl_governador.Visible = False
            lbl_candidato.Visible = False
            lbl_vice.Visible = False
            lbl_partido.Visible = False
            lbl_partido2.Visible = False
            Call labels2
                     
            If lbl_outro = "VOTO EM BRANCO" Then
               txt_ladrao = "-2"
               txt_ladrao.Visible = False
               gvoto = tab_gcd!qtde_votos
               gvoto = pvoto + 1
               
            End If
            

End Sub

Private Sub cmd_confirmar_Click()

             If Len(txt_ladrao) <> 2 Then
             Exit Sub
              End If
            gvoto = tab_gcd!qtde_votos
            gvoto = gvoto + 1
                   
            If lbl_outro = "VOTO NULO" Then
               txt_ladrao = "-1"
               txt_ladrao.Enabled = False
               gvoto = tab_gcd!qtde_votos
               gvoto = gvoto + 1
            End If
            
            If tab_gcd.State = adStateOpen Then
                tab_gcd.Close
                tab_gcd.Open "Select * from Governadores where nome_politico = '" & lbl_candidato & "'"
                    
                    If tab_gcd.RecordCount <> 0 Then
                       tab_gcd.Close
                       tab_gcd.Open "Update Governadores set qtde_votos = '" & gvoto & "' where nome_politico = '" & lbl_candidato & "'"
                                            
                       ElseIf tab_gcd.RecordCount = 0 Then
                              tab_gcd.Close
                              tab_gcd.Open "Update Governadores set qtde_votos = '" & gvoto & "' where nome_politico = '" & lbl_presidente & "'"
                    End If

            End If
            
            Unload Me
            frm_pvoto.Show
End Sub

Private Sub cmd_corrigir_Click()
            On Error Resume Next
            Call limpar
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
            Select Case KeyCode
            Case vbKeySubtract
            Call cmd_confirmar_Click
            Case vbKeyMultiply
            Call cmd_corrigir_Click
            Case vbKeyDivide
            Call cmd_branco_Click
            End Select
            
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
            If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
            Call OpenUrna
            way = App.Path & "\Imagens\"
End Sub


Private Sub txt_ladrao_Change()
            
            If lbl_outro = "VOTO EM BRANCO" Then Exit Sub
            
            If Len(txt_ladrao) = 2 Then
            GoTo B
            Else
            Exit Sub
            End If
            
            
            
B:          If tab_gcd.State = adStateOpen Then
               tab_gcd.Close
                   tab_gcd.Open "Select * from Governadores where numero_candidato = " & txt_ladrao
                   If tab_gcd.RecordCount <> 0 Then
                        Call Labels
                        Call labels2
                        lbl_candidato = tab_gcd!nome_politico
                        lbl_vicegovernador = tab_gcd!vice
                        lbl_outro.Visible = False
                        If tab_par.State = adStateOpen Then tab_par.Close
                          tab_par.Open "select * from Partidos where legenda = " & txt_ladrao
                          If tab_par.RecordCount <> 0 Then
                             lbl_partido2 = tab_par!sigla
                          End If
                          ElseIf tab_gcd.RecordCount = 0 Then
                               lbl_outro.Visible = True
                               Call labels2
                               lbl_candidato = "Número Incorreto"
                               lbl_outro = "VOTO NULO"
             
             End If
             End If
           
                   
        
End Sub

Private Sub txt_ladrao_KeyPress(KeyAscii As Integer)
            Select Case KeyAscii
            Case 48 To 57
            Case Else
            Beep
            KeyAscii = 0
            End Select
         
            If KeyAscii = 46 Then
               KeyAscii = 0
            End If

End Sub




