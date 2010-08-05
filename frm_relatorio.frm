VERSION 5.00
Begin VB.Form frm_relatorio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visualização de Relatórios"
   ClientHeight    =   2790
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Visualização e Impressão"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   4215
      Begin VB.PictureBox cr 
         Height          =   480
         Left            =   3240
         ScaleHeight     =   420
         ScaleWidth      =   660
         TabIndex        =   6
         Top             =   360
         Width           =   720
      End
      Begin VB.CommandButton cmd_visualizar 
         Caption         =   "Visualizar"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmd_imprimir 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controle de Relatórios"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox cbo_relatorio 
         Height          =   315
         ItemData        =   "frm_relatorio.frx":0000
         Left            =   1920
         List            =   "frm_relatorio.frx":0002
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Escolha o Relatório:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_relatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_imprimir_Click()
            If cbo_relatorio = Empty Then
                MsgBox "Desculpe-me, mas é necessário selecionar um relatório antes de imprimir-lo", vbExclamation, "Allia"
                ElseIf cbo_relatorio = Presidente Then
                    cr.ReportFileName = "F:\Projeto Allia\Presidente.rpt"
                    cr.RetrieveDataFiles
                    cr.Action = 1
                    cr.Destination = crptToPrinter
            End If
               
            
End Sub

Private Sub cmd_visualizar_Click()
           If cbo_relatorio = Empty Then
                MsgBox "Desculpe-me, mas é necessário selecionar um relatório antes de imprimir-lo", vbExclamation, "Allia"
                ElseIf cbo_relatorio = Presidente Then
                    cr.ReportFileName = "F:\Projeto Allia\Presidente.rpt"
                    cr.RetrieveDataFiles
                    cr.Action = 1
                    cr.Destination = crptToWindow
            End If
            
            
               If cbo_relatorio.Text = "Zerésima de Presidentes" Then
               cr.Destination = 0
               cr.ReportFileName = App.Path & "\zeresimap.rpt"
               cr.RetrieveDataFiles
               cr.Destination = crptToWindow
               cr.Action = 1
            End If
            
            If cbo_relatorio.Text = "Sumula de Presidentes" Then
               cr.Destination = 0
               cr.ReportFileName = App.Path & "\sumulap.rpt"
               cr.RetrieveDataFiles
               cr.Destination = crptToWindow
               cr.Action = 1
            End If
            
            If cbo_relatorio.Text = "Zerésima de Governadores" Then
               cr.Destination = 0
               cr.ReportFileName = App.Path & "\zeresimag.rpt"
               cr.RetrieveDataFiles
               cr.Destination = crptToWindow
               cr.Action = 1
            End If
            
             If cbo_relatorio.Text = "Sumula de Presidentes" Then
               cr.Destination = 0
               cr.ReportFileName = App.Path & "\sumulag.rpt"
               cr.RetrieveDataFiles
               cr.Destination = crptToWindow
               cr.Action = 1
            End If
End Sub
