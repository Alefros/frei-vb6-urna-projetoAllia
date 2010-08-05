VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frm_eleitores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de eleitores"
   ClientHeight    =   2325
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fr_celeitores 
      Caption         =   "Cadastro de Eleitores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.TextBox txt_nome 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmd_cadastrar 
         Caption         =   "Cadastrar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   960
         TabIndex        =   1
         Top             =   1440
         Width           =   1335
      End
      Begin MSMask.MaskEdBox msk_rg 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   12
         Mask            =   "99.999.999-9"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "RG"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   420
      End
   End
End
Attribute VB_Name = "frm_eleitores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Sub cmd_cadastrar_Click()
                msk_rg.PromptInclude = False
                If tab_ele.State = adStateOpen Then
                tab_ele.Close
                End If
                tab_ele.Open "Select * from Eleitores where RG = '" & msk_rg & "'"
                    If tab_ele.RecordCount = 0 Then
                    bdvb.Execute "Insert into Eleitores(RG, Nome)values('" & msk_rg & "', '" & txt_nome & "')"
                    MsgBox "Eleitor cadastrado com sucesso."
                    Unload Me
                    frm_pvoto.Show
                    
                    Else
                    MsgBox "Este eleitor já votou."
                    msk_rg = Clear
                    txt_nome = Clear
                    msk_rg.PromptInclude = True
                   
                   End If
           

End Sub

Private Sub Form_Load()
           ' Call OpenBD
            Call OpenUrna
End Sub

Private Sub msk_rg_Change()
            msk_rg.PromptInclude = False
            If msk_rg = Empty Then
            cmd_cadastrar.Enabled = False
            Else
            cmd_cadastrar.Enabled = True
            End If
            msk_rg.PromptInclude = True
End Sub

Private Sub Sair_Click()
            Unload Me
End Sub

Private Sub txt_nome_Change()
            If txt_nome = Empty Then
            cmd_cadastrar.Enabled = False
            Else
            Exit Sub
            End If
End Sub
