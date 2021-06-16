VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Report_Umum 
   BackColor       =   &H80000007&
   Caption         =   "Report_Umum"
   ClientHeight    =   1950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Laporan Penjualan Obat Apotik Kelompok 6"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Text            =   "Pilih"
         Top             =   480
         Width           =   1815
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   120
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CETAK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   2520
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pilih Tanggal   : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Report_Umum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CrystalReport1.SelectionFormula = "totext({data_transaksi.tanggal_transaksi})='" & CDate(Combo1) & "'"
CrystalReport1.ReportFileName = App.Path & "\report1.rpt"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 1
End Sub

Private Sub Form_Load()
Call konekdb
RStransaksi.Open "select distinct tanggal_transaksi from data_transaksi", konek
Combo1.Clear
Do While Not RStransaksi.EOF
    Combo1.AddItem RStransaksi!Tanggal_Transaksi
    RStransaksi.MoveNext
Loop
End Sub
