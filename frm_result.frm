VERSION 5.00
Begin VB.Form frm_result 
   BorderStyle     =   0  'None
   Caption         =   "Resultados"
   ClientHeight    =   8655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   13035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frm_result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dias As Integer
Dim linha As Integer
Dim mes As Integer
Dim freq(8) As Long
Dim Freq_ac(8) As Long
Dim Pm(8) As Double
Dim pm_X As Double
Dim Fatu(8, 2) As Long

Private Sub Command1_Click()
            
            
            Dim lamina As Double
            cont1 = 1
            While cont1 <> Empty
            If Pm(cont1) <> 0 Then
               pm_X = pm_X + ((Pm(cont1) - media) ^ 2)
            End If
            cont1 = cont1 + 1
            If cont1 >= 8 Then cont1 = Empty

Wend

            graf1.chartType = VtChChartType3dBar
            graf1.ShowLegend = False
            graf1.Title = "Votagens presidentes"
            cont1 = 1
            While cont1 <> Empty
            If Fatu(cont1, 1) = 0 Then GoTo fuck
            msc.RowCount = cont1
            msc.Row = cont1
            msc.RowLabel = Fatu(cont1, 1)
            msc.Data = freq(cont1)
fuck:
            cont1 = cont1 + 1
            If cont1 > 8 Then cont1 = Clear
Wend

End Sub


