Attribute VB_Name = "Module1"
Option Explicit

Global bdvb As New ADODB.Connection

Global tab_ele As New ADODB.Recordset

Global tab_pcd As New ADODB.Recordset

Global tab_gcd As New ADODB.Recordset

Global tab_par As New ADODB.Recordset

Global tab_coli As New ADODB.Recordset

Global tab_pais As New ADODB.Recordset

Global tab_estados As New ADODB.Recordset

Global tab_municipios As New ADODB.Recordset
                                                
Global tab_regioes As New ADODB.Recordset
                                                                                            
Global tab_areas As New ADODB.Recordset

Global caminho As String

Global status As String
Global tab_loca As New ADODB.Recordset
Global tab_estado As New ADODB.Recordset

Global tab_turma As New ADODB.Recordset

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
    ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" ( _
                ByVal hwnd As Long, ByVal nindex As Long, ByVal dwnewlong As Long) As Long

Public Declare Function InvalidateRect Lib "user32" _
                (ByVal hwnd As Long, lpRect As Long, ByVal bErase As Long) As Long

Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

Public Enum enuTBType
    enuTB_FLAT = 1
    enuTB_STANDARD = 2
End Enum

Private Const GCL_HBRBACKGROUND = (-10)

Public Sub ChangeTBBack(TB As Object, PNewBack As Long, pType As enuTBType)
Dim lTBWnd      As Long

    Select Case pType
        
        Case enuTB_FLAT
            DeleteObject SetClassLong(TB.hwnd, GCL_HBRBACKGROUND, PNewBack)
    End Select
End Sub

Function OpenBD()
         caminho = "Provider = microsoft.jet.oledb.4.0; data source = " & App.Path & "\Urna.mdb"
         If bdvb.State = adStateOpen Then
         bdvb.Close
         End If
         bdvb.Open (caminho)
End Function

Function box_1()
         MsgBox "Informações" & status & " com sucesso! ", vbInformation, "Allia"
End Function

Public Sub OpenUrna()
           If tab_pcd.State = adStateOpen Then tab_pcd.Close
           If tab_gcd.State = adStateOpen Then tab_gcd.Close
           If tab_ele.State = adStateOpen Then tab_ele.Close
           If tab_par.State = adStateOpen Then tab_par.Close
           If tab_estados.State = adStateOpen Then tab_estados.Close
           If tab_regioes.State = adStateOpen Then tab_regioes.Close
           If tab_pais.State = adStateOpen Then tab_pais.Close
           
         
           tab_estados.Open "Estados", bdvb, adOpenKeyset, adLockOptimistic
           tab_ele.Open "Alunos", bdvb, adOpenKeyset, adLockOptimistic
           tab_regioes.Open "Regioes", bdvb, adOpenKeyset, adLockOptimistic
           tab_pais.Open "Paises", bdvb, adOpenKeyset, adLockBatchOptimistic
           tab_pcd.Open "Presidentes", bdvb, adOpenKeyset, adLockOptimistic
           tab_gcd.Open "Governadores", bdvb, adOpenKeyset, adLockOptimistic
           tab_par.Open "Partidos", bdvb, adOpenKeyset, adLockOptimistic
End Sub
