VERSION 5.00
Begin VB.Form frm_confg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurações"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbo_cargo 
      Height          =   315
      ItemData        =   "frm_confg.frx":0000
      Left            =   1560
      List            =   "frm_confg.frx":000A
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmd_votos 
      Caption         =   "ZERAR VOTOS"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CARGO"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   960
      Width           =   570
   End
End
Attribute VB_Name = "frm_confg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public contador As Integer

Private Sub cmd_votos_Click()
           contador = 1
           Dim merda As String
           merda = "0"
           
          ' If MsgBox("Deseja mesmo excluir os votos?", vbQuestion + vbYesNo, "Allia") = vbYes Then
           If cbo_cargo.Text = "" Then Exit Sub
           If cbo_cargo.Text = "Presidentes" Then GoTo A
           If cbo_cargo.Text = "Governadores" Then GoTo B
           
A:
           If tab_pcd.State = adStateOpen Then tab_pcd.Close
           tab_pcd.Open
           If contador > tab_pcd.RecordCount Then GoTo C
           tab_pcd.Update
           tab_pcd!qtde_votos = merda
           tab_pcd.Update
           tab_pcd.MoveNext
           contador = contador + 1
           GoTo A

B:
           If tab_gcd.State = adStateOpen Then tab_gcd.Close
           tab_pcd.Open
           If contador > tab_gcd.RecordCount Then GoTo C
           tab_gcd.Update
           tab_gcd!qtde_votos = merda
           tab_gcd.Update
           tab_gcd.MoveNext
           contador = contador + 1
           GoTo B
           
C:
           MsgBox "Informações excluidas com sucesso!", vbInformation, "Allia"
           'End If
           
End Sub

Private Sub Form_Load()
            Call OpenUrna
End Sub
