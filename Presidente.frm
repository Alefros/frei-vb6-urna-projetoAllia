VERSION 5.00
Begin VB.Form frm_pvoto 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Voto"
   ClientHeight    =   9450
   ClientLeft      =   840
   ClientTop       =   420
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   13395
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
      Left            =   1800
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
      BackColor       =   &H000000FF&
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
      Width           =   105
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
      Width           =   105
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
      Width           =   105
   End
   Begin VB.Shape Shape1 
      Height          =   15
      Left            =   0
      Top             =   6720
      Width           =   13455
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
   Begin VB.Label lbl_vicepresidente 
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
      Left            =   3720
      TabIndex        =   8
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label lbl_presidente 
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
      Width           =   1545
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
      Caption         =   "PRESIDENTE"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   1200
      Width           =   4095
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
      Height          =   5415
      Left            =   7440
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5295
   End
   Begin VB.Menu mnupopup 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnupop 
         Caption         =   "Pop"
      End
      Begin VB.Menu mnuup 
         Caption         =   "Up"
      End
   End
End
Attribute VB_Name = "frm_pvoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Way As String

Dim img As String

Dim pvoto As Integer

Dim l_pres As Integer

Private Sub Labels()
            img = tab_pcd!imagem
            img_ladrao.Picture = LoadPicture(Way & img)
            lbl_seuvotopara = "SEU VOTO PARA"
            lbl_numero = "Número:"
            lbl_presidente = "Presidente:"
            lbl_vice = "Vice-presidente:"
            lbl_partido = "Partido:"
            lbl_partido2 = "PT"
            lbl_tecla = "Aperte a tecla:"
            lbl_verde = "VERDE para CONFIRMAR"
            lbl_laranja = "LARANJA para CORRIGIR"
End Sub

Private Sub limpar()
            txt_ladrao = Clear
            img_ladrao.Picture = LoadPicture(Empty)
            lbl_seuvotopara = "SEU VOTO PARA"
            lbl_numero = Clear
            lbl_presidente = Clear
            lbl_candidato = Clear
            lbl_vice = Clear
            lbl_vicepresidente = Clear
            lbl_partido = Clear
            lbl_partido2 = Clear
            lbl_tecla = Clear
            lbl_verde = Clear
            lbl_laranja = Clear
End Sub

Private Sub cmd_confirmar_Click()
            
            Dim num As String
            num = "Numero Incorreto"
            pvoto = tab_pcd!qtde_votos
            pvoto = pvoto + 1
            l_pres = txt_ladrao
            If tab_pcd.State = adStateOpen Then
                tab_pcd.Close
                tab_pcd.Open "Select * from Presidentes where nome = '" & lbl_candidato & "'"
                    If tab_pcd.RecordCount <> 0 Then
                        
                        tab_pcd.Close
                        tab_pcd.Open "Update Presidentes set qtde_votos = '" & pvoto & "' where nome = '" & lbl_candidato & "'"
                                            
                        ElseIf tab_pcd.RecordCount = 0 Then
                            tab_pcd.Close
                            tab_pcd.Open "Update Presidentes set qtde_votos = '" & pvoto & "' where nome = '" & lbl_presidente & "'"
                    End If
                    
            End If
                            
            frm_pvoto.Hide
            frm_gvoto.Show
            
            
End Sub

Private Sub cmd_corrigir_Click()
            'Call limpar
End Sub



Private Sub Form_Load()
            Call OpenBD
            tab_pcd.Open "Presidentes", bdvb, adOpenKeyset, adLockOptimistic
            
            Way = "D:\Cleiton\estudo\Urna2\Imagens\"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As _
Integer, X As Single, Y As Single)
'Mostra o menu popup _'se o botão direito for clicado
If Button = 2 Then PopupMenu mnupopup
End Sub



Private Sub txt_ladrao_Change()
            If tab_pcd.State = adStateOpen Then
                   tab_pcd.Close
                    tab_pcd.Open "Select * from Presidentes where cod_pres = " & txt_ladrao
                    If tab_pcd.RecordCount <> 0 Then
                        Call Labels
                        lbl_candidato = tab_pcd!nome
                        lbl_vicepresidente = tab_pcd!vice
                        lbl_partido2 = tab_pcd!chapa_partido
                        ElseIf tab_pcd.RecordCount = 0 Then
                            'txt_ladrao.MaxLength = 2 Then
                                lbl_presidente = "Número Incorreto"
                    End If
            End If
            
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
