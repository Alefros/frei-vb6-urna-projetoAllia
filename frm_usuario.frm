VERSION 5.00
Begin VB.Form frm_login 
   Caption         =   "Login"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_senha 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txt_usuario 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmd_entrar 
      Caption         =   "Entrar"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lbl_usuario 
      AutoSize        =   -1  'True
      Caption         =   "Usuário"
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   465
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_entrar_Click()
            If txt_usuario = "" Then
            MsgBox "Preencha todos os campos.", vbInformation, "Login"
            Exit Sub
            End If
            If txt_senha = "" Then
            MsgBox "Preencha todos os campos.", vbInformation, "Login"
            Exit Sub
            End If
            If tab_usu.State = adStateOpen Then
            tab_usu.Close
            End If
            tab_usu.Open "select * from Usuários where Nome = '" & txt_usuario & "'"
            If tab_usu.RecordCount <> 0 Then
            If txt_senha = tab_usu!Senha Then
            Unload Me
            mdi_urna.Show
            ElseIf txt_senha <> tab_usu!Senha Then
            MsgBox "Nome de usuário ou senha inválido.", vbInformation, "Login"
            End If
            ElseIf tab_usu.RecordCount = 0 Then
            MsgBox "Nome de usuário ou senha inválido.", vbInformation, "Login"
            End If

                                   
End Sub

Private Sub Form_Load()
            Call OpenBD
            tab_usu.Open "Usuários", bdvb, adOpenKeyset, adLockOptimistic
            
            'Skin1.LoadSkin "D:\Cleiton\estudo\Urna\Skins+++\SKins+++\Corona.skn"
            'Skin1.ApplySkin Me.hWnd

End Sub

