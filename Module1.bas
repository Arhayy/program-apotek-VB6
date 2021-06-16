Attribute VB_Name = "Koneksi"
Public konek As ADODB.Connection
Public RSadmin As ADODB.Recordset
Public RStransaksi As ADODB.Recordset
Public RSpelanggan As ADODB.Recordset
Public RSobat As ADODB.Recordset
Public RSdokter As ADODB.Recordset
Public RStransaksi_resep As ADODB.Recordset
Public RSumum As ADODB.Recordset

Public lokasidb As String

Public Sub konekdb()
Set konek = New ADODB.Connection
Set RSadmin = New ADODB.Recordset
Set RStransaksi = New ADODB.Recordset
Set RSpelanggan = New ADODB.Recordset
Set RSobat = New ADODB.Recordset
Set RSdokter = New ADODB.Recordset
Set RStransaksi_resep = New ADODB.Recordset
Set RSumum = New ADODB.Recordset

lokasidb = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbapotik.mdb"
konek.Open lokasidb
End Sub
