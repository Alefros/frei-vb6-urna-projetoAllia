Attribute VB_Name = "Module1"
Global bdvb As New ADODB.Connection

Global tab_ptd As New ADODB.Recordset

Global tab_pcd As New ADODB.Recordset

Global tab_gcd As New ADODB.Recordset

Global tab_scd As New ADODB.Recordset

Global tab_fcd As New ADODB.Recordset

Global tab_ecd As New ADODB.Recordset

Global caminho As String

Function OpenBD()
        caminho = "Provider = microsoft.jet.oledb.4.0; data source = D:\Cleiton\estudo\Urna2\urna.mdb"
        bdvb.Open caminho
End Function

