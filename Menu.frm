VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Menu 
   BackColor       =   &H0000FF00&
   Caption         =   "Menu"
   ClientHeight    =   10320
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17640
   DrawStyle       =   1  'Dash
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   10320
   ScaleWidth      =   17640
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9945
      Width           =   17640
      _ExtentX        =   31115
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   10500
      Left            =   -360
      Picture         =   "Menu.frx":2C4B1
      Top             =   -120
      Width           =   21000
   End
   Begin VB.Menu menu_master 
      Caption         =   "Menu"
      Begin VB.Menu menu_dataobat 
         Caption         =   "Data Obat"
      End
      Begin VB.Menu menu_dokter 
         Caption         =   "Data Dokter"
      End
      Begin VB.Menu Menu_HapusDS 
         Caption         =   "Data Pelanggan"
      End
   End
   Begin VB.Menu menu_transaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu menu_transaksi_umum 
         Caption         =   "Transaksi Umum"
      End
      Begin VB.Menu menu_transaksi_resep 
         Caption         =   "Transaksi Resep"
      End
   End
   Begin VB.Menu menu_report 
      Caption         =   "Laporan"
      Begin VB.Menu menu_laporan_umum 
         Caption         =   "Laporan Penjualan Umum"
      End
      Begin VB.Menu menu_laporan_resep 
         Caption         =   "Laporan Penjualan Resep"
      End
   End
   Begin VB.Menu aksi 
      Caption         =   "Aksi"
      Begin VB.Menu menu_datapembeli 
         Caption         =   "Daftar Data Pembeli"
      End
      Begin VB.Menu ganti_password 
         Caption         =   "Ganti Kata Sandi"
      End
      Begin VB.Menu menu_logout 
         Caption         =   "Keluar"
         WindowList      =   -1  'True
      End
      Begin VB.Menu tutup_program 
         Caption         =   "Tutup Program"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ganti_password_Click()
ubah_password.Show
End Sub

Private Sub menu_dataobat_Click()
Data_Obat.Show
End Sub

Private Sub menu_datapembeli_Click()
RegisData_Pelanggan.Show
End Sub

Private Sub menu_dokter_Click()
Data_Dokter.Show
End Sub

Private Sub Menu_HapusDS_Click()
DataPelanggan.Show
End Sub

Private Sub menu_laporan_resep_Click()
Report_Resep.Show
End Sub

Private Sub menu_laporan_umum_Click()
Report_Umum.Show
End Sub

Private Sub menu_logout_Click()
    warning = MsgBox("Apakah Anda Yakin Ingin Keluar!!", vbYesNo + vbInformation)
If warning = vbYes Then
    Login.Show
    Menu.Hide
End If
End Sub

Private Sub menu_transaksi_resep_Click()
Data_Transaksi_Resep.Show
End Sub

Private Sub menu_transaksi_umum_Click()
Data_Transaksi.Show
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(4) = Time$
End Sub

Private Sub tutup_program_Click()
    warning = MsgBox("Apakah Anda Yakin Ingin Menutup Program!!", vbYesNo + vbInformation)
If warning = vbYes Then
    End
End If
End Sub
