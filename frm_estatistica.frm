VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frm_estatistica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estatística"
   ClientHeight    =   12450
   ClientLeft      =   210
   ClientTop       =   765
   ClientWidth     =   18465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12450
   ScaleWidth      =   18465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_escolha 
      Caption         =   "Escolha de Gráfico"
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
      Top             =   10800
      Width           =   3735
      Begin VB.ComboBox cbo_graf 
         Height          =   315
         ItemData        =   "frm_estatistica.frx":0000
         Left            =   1920
         List            =   "frm_estatistica.frx":0010
         TabIndex        =   5
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmd_analisar 
         Caption         =   "Analisar"
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cbo_cargo 
         Height          =   315
         ItemData        =   "frm_estatistica.frx":002F
         Left            =   1920
         List            =   "frm_estatistica.frx":0039
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Selecione o cargo "
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Gráfico"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame fra_graf 
      Caption         =   "Amostra de Gráficos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   18255
      Begin MSChart20Lib.MSChart msc_graf 
         Height          =   10095
         Left            =   120
         OleObjectBlob   =   "frm_estatistica.frx":0055
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   17895
      End
   End
End
Attribute VB_Name = "frm_estatistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public contador As Integer

Private Sub cmd_analisar_Click()
            
            contador = 1
            
            If tab_pcd.State = adStateOpen Then tab_pcd.Close
            tab_pcd.Open "select * from Presidentes order by qtde_votos desc"
            If tab_gcd.State = adStateOpen Then tab_gcd.Close
            tab_gcd.Open "select * from Governadores order by qtde_votos desc"
            
            If cbo_graf.Text = "" Then Exit Sub
            If cbo_cargo.Text = "" Then Exit Sub
            
            
            If cbo_graf.Text = "Área" Then msc_graf.chartType = VtChChartType2dArea
            
            If cbo_graf.Text = "Barra" Then msc_graf.chartType = VtChChartType2dBar
            
            If cbo_graf.Text = "Linha" Then msc_graf.chartType = VtChChartType2dLine
            
            If cbo_graf.Text = "Torta" Then msc_graf.chartType = VtChChartType2dPie
                  
           If cbo_cargo.Text = "Presidente" Then GoTo A
           
           If cbo_cargo.Text = "Governador" Then GoTo B
           
            
            
A:
            On Error Resume Next
            If contador > tab_pcd.RecordCount Then
            GoTo C
            End If
            msc_graf.TitleText = "OS 5 PRESIDENTES MAIS VOTADOS"
            msc_graf.Row = contador
            msc_graf.RowLabel = tab_pcd!nome_politico
            msc_graf.Data = tab_pcd!qtde_votos
            tab_pcd.MoveNext
            contador = contador + 1
            GoTo A
            
B:
            On Error Resume Next
            If contador > tab_gcd.RecordCount Then
            GoTo C
            End If
            msc_graf.TitleText = "OS 5 GOVERNADORES MAIS VOTADOS"
            msc_graf.Row = contador
            msc_graf.RowLabel = tab_gcd!nome_politico
            msc_graf.Data = tab_gcd!qtde_votos
            tab_gcd.MoveNext
            contador = contador + 1
            GoTo B
                        
            
C:
            msc_graf.Visible = True
            Exit Sub
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
            If KeyAscii = 27 Then Unload Me
            If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
            Call OpenUrna
End Sub


