VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RegisData_Pelanggan 
   BackColor       =   &H000000FF&
   Caption         =   "Daftar Data Pelanggan"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10740
   LinkTopic       =   "Form2"
   ScaleHeight     =   6720
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   6600
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\apotik\dbapotik.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\apotik\dbapotik.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Data_Pelanggan"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2280
      TabIndex        =   15
      Top             =   1440
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Data_Supplier.frx":0000
      Height          =   3375
      Left            =   360
      TabIndex        =   14
      Top             =   3240
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Enabled         =   0   'False
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Kode_Pelanggan"
         Caption         =   "Kode_Pelanggan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nama_Pelanggan"
         Caption         =   "Nama_Pelanggan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Telp"
         Caption         =   "Telp"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Alamat"
         Caption         =   "Alamat"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BATAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   13
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cari"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7320
      TabIndex        =   6
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7320
      TabIndex        =   5
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pencarian"
      Height          =   855
      Left            =   360
      TabIndex        =   9
      Top             =   2400
      Width           =   9975
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label7 
         Caption         =   "Cari Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5520
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Telp/Hp"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5520
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Pelanggan"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode Pelanggan"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   11445
      Left            =   -720
      Picture         =   "Data_Supplier.frx":0015
      Top             =   -240
      Width           =   12000
   End
End
Attribute VB_Name = "RegisData_Pelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub
Sub cekdata()
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" Then
    MsgBox ("Data Pelanggan Belum Lengkap")
End If
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
    MsgBox ("Data Pelanggan Belum Lengkap"), vbCritical
    Call bersih
Else
    Adodc1.Recordset.AddNew
    Adodc1.Recordset!kode_pelanggan = Text1.Text
    Adodc1.Recordset!Nama_Pelanggan = Text2.Text
    Adodc1.Recordset!Telp = Text3.Text
    Adodc1.Recordset!Alamat = Text4.Text
        MsgBox "Data Pelanggan Disimpan", vbInformation, "INFO"
    DataGrid1.Refresh
    Adodc1.Recordset.Update
    Call bersih
End If
End Sub

Private Sub Form()
Call bersih
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Call bersih
End Sub

Private Sub Command3_Click()
warning = MsgBox("Apakah Anda Ingin Keluar!!", vbYesNo + vbInformation)
    If warning = vbYes Then
        RegisData_Pelanggan.Hide
    Else
    End If
End Sub

Private Sub Command5_Click()
Call konekdb
If Text5.Text = "" Then
        MsgBox "Silahkan Isi Data", vbCritical
Else
    RSpelanggan.CursorLocation = adUseClient
    RSpelanggan.Open "SELECT * FROM Data_Pelanggan WHERE Kode_Pelanggan like '%" & Text5 & "%'", konek
        If Not RSpelanggan.EOF Then
            With RSpelanggan
            With DataGrid1
                Set .DataSource = RSpelanggan
                .Refresh
            End With
            End With
        Else
            MsgBox "Data Tidak Ditemukan", vbExclamation
            Text5.Text = ""
        End If
End If
End Sub

Private Sub Form_Load()
Set MyControl = DataGrid1
WheelHook DataGrid1

Text1.Enabled = False
DataGrid1.Columns(0).Width = 1700
DataGrid1.Columns(1).Width = 3200
DataGrid1.Columns(2).Width = 2360
DataGrid1.Columns(3).Width = 2380
DataGrid1.HeadFont.Bold = True
End Sub

Sub kode_pelanggan()
Call konekdb
RSpelanggan.Open ("Select * from Data_Pelanggan Where Kode_Pelanggan In(Select Max(Kode_Pelanggan)From Data_Pelanggan)Order By Kode_Pelanggan Desc"), konek
RSpelanggan.Requery
    Dim urutan As String * 5
    Dim hitung As Long
    With RSpelanggan
        If .EOF Then
            urutan = "KP" + "001"
            Text1.Text = urutan
        Else
            hitung = Right(!kode_pelanggan, 3) + 1
            urutan = "KP" + Right("000" & hitung, 3)
        End If
            Text1.Text = urutan
    End With
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(Text2) = True Then
    MsgBox "Block Number", vbCritical
    Text2 = ""
Exit Sub
End If
    Call kode_pelanggan
    Text3.SetFocus
End If
End Sub

Private Sub Text3_Change()
If Len(Text3) > 13 Then
    MsgBox "Data Yang Anda Masukan Terlalu Banyak", vbCritical
        Text3.Text = ""
        Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(Text3) = False Then
    MsgBox "Block Huruf", vbCritical
    Text3 = ""
Exit Sub
End If
    Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(Text5) = True Then
    MsgBox "Block Number", vbCritical
    Text5 = ""
Exit Sub
End If
    Command5.SetFocus
End If
End Sub
