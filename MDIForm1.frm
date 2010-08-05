VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdi_urna 
   BackColor       =   &H00FFFFFF&
   Caption         =   "~ Allia"
   ClientHeight    =   12990
   ClientLeft      =   -105
   ClientTop       =   1290
   ClientWidth     =   15510
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      Picture         =   "MDIForm1.frx":0000
      ScaleHeight     =   915
      ScaleWidth      =   15450
      TabIndex        =   2
      Top             =   10695
      Visible         =   0   'False
      Width           =   15510
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12600
      Top             =   8640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":167F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":31D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":3F25
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":45C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":57BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":5B6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6665
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":6F35
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7A15
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8767
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":94B9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb_padrao 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15510
      _ExtentX        =   27358
      _ExtentY        =   1905
      ButtonWidth     =   2408
      ButtonHeight    =   1852
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "    Localizações  "
            Object.ToolTipText     =   "Cadastro de e controle de localizações."
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Partidos"
            Object.ToolTipText     =   "Cadastro e controle de partidos."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Candidatos"
            Object.ToolTipText     =   "Cadastro e controle de governadores."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eleitores"
            Object.ToolTipText     =   "Cadastro de eleitores."
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Urna"
            Object.ToolTipText     =   "Iniciar votação."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Resultados"
            Object.ToolTipText     =   "Ver resultados, candidatos vencedores e gráficos."
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Relatórios"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
      MousePointer    =   1
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   9615
      Left            =   0
      ScaleHeight     =   9555
      ScaleWidth      =   15450
      TabIndex        =   0
      Top             =   1080
      Width           =   15510
      Begin VB.Image Image1 
         Height          =   16200
         Left            =   -120
         Picture         =   "MDIForm1.frx":B00B
         Top             =   0
         Width           =   28800
      End
   End
End
Attribute VB_Name = "mdi_urna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ApplyChanges()
    Dim LngNew As Long
   
    LngNew = CreatePatternBrush(Picture2.Picture.Handle)
    ChangeTBBack tlb_padrao, LngNew, enuTB_FLAT
   
    InvalidateRect 0&, 0&, False
End Sub


Private Sub MDIForm_Resize()
    Picture1.Cls
    mdi_urna.Picture = LoadPicture("")
    Picture1.Visible = True
    Picture1.AutoRedraw = True
    Picture1.BackColor = &H8000000C
    Picture1.Height = Me.Height
    'Para centralizar a imagem no fundo
    'Image1.Top = Picture1.Height / 2 - Image1.Height / 2
    'Image1.Left = Picture1.Width / 2 - Image1.Width / 2
    'ou expandir a imagem por todo o fundo
    Image1.Stretch = True
    Image1.Top = 0
    Image1.Left = 0
    Image1.Height = Picture1.Height
    Image1.Width = Picture1.Width

    Picture1.PaintPicture Image1, Image1.Left, Image1.Top, Image1.Width, Image1.Height
    mdi_urna.Picture = Picture1.Image
    Picture1.Visible = False
End Sub

Private Sub MDIForm_Load()
            Call ApplyChanges
            Call OpenBD
End Sub

Private Sub tlb_padrao_ButtonClick(ByVal Button As MSComctlLib.Button)
            Select Case Button.Index
            Case 1
                frmloca.Show
            Case 3
                frm_partidos.Show
            Case 5
                frm_candidatos.Show
            Case 7
                frm_releitores.Show
            Case 9
                frm_gvoto.Show
            Case 11
                frm_estatistica.Show
            Case 13
                frm_relatorio.Show
            Case 15
                End
            End Select
End Sub
