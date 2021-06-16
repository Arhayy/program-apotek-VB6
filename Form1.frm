VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Data_Obat 
   BackColor       =   &H0000FFFF&
   Caption         =   "Data Obat"
   ClientHeight    =   7065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   6960
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
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
      RecordSource    =   "tabel_obat"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2895
      Left            =   0
      TabIndex        =   13
      Top             =   4080
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Enabled         =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Kode_Obat"
         Caption         =   "Kode_Obat"
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
         DataField       =   "Nama_Obat"
         Caption         =   "Nama_Obat"
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
         DataField       =   "Satuan"
         Caption         =   "Satuan"
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
         DataField       =   "Harga"
         Caption         =   "Harga"
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
      BeginProperty Column04 
         DataField       =   "Stok"
         Caption         =   "Stok"
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
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command6 
      Caption         =   "UBAH DATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
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
      Height          =   400
      Left            =   6240
      TabIndex        =   11
      Top             =   1920
      Width           =   2340
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
      Left            =   6240
      TabIndex        =   10
      Top             =   960
      Width           =   2340
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cari"
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
      Left            =   6600
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Caption         =   "Pencarian"
      Height          =   855
      Left            =   -120
      TabIndex        =   6
      Top             =   3240
      Width           =   9135
      Begin VB.TextBox Text6 
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
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   5100
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cari Obat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command4 
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
      Left            =   7320
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HAPUS"
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
      Left            =   5760
      TabIndex        =   4
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
      Left            =   2280
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
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
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   2340
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
      TabIndex        =   1
      Top             =   1440
      Width           =   6300
   End
   Begin VB.TextBox text1 
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
      TabIndex        =   0
      Top             =   960
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   0
      Picture         =   "Form1.frx":0015
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "Data_Obat"
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
Text6.Text = ""
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox ("Data Obat Belum Lengkap"), vbCritical
Call bersih
Else
    Adodc1.Recordset.AddNew
    Adodc1.Recordset!Kode_Obat = Text1.Text
    Adodc1.Recordset!Nama_Obat = Text2.Text
    Adodc1.Recordset!satuan = Text3.Text
    Adodc1.Recordset!harga = Text4.Text
    Adodc1.Recordset!Stok = Text5.Text
        MsgBox "Data Berhasil Disimpan", vbInformation, "INFO"
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
    warning = MsgBox("Apakah Ingin Menghapus!!", vbYesNo + vbInformation)
If warning = vbYes Then
    Adodc1.Recordset.Delete
Else
End If
Call bersih
End Sub

Private Sub Command4_Click()
    warning = MsgBox("Apakah Anda Ingin Keluar!!", vbYesNo + vbInformation)
If warning = vbYes Then
    Data_Obat.Hide
Call bersih
Else
End If
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Find "Nama_Obat='" & Text6.Text & "'", , adSearchForward, 1
    If Not Adodc1.Recordset.EOF Then
            Text1.Text = Adodc1.Recordset!Kode_Obat
            Text2.Text = Adodc1.Recordset!Nama_Obat
            Text3.Text = Adodc1.Recordset!satuan
            Text4.Text = Adodc1.Recordset!harga
            Text5.Text = Adodc1.Recordset!Stok
    Else
        MsgBox ("Data Obat Tidak Ada"), vbExclamation
            Text6.Text = ""
    End If
End Sub

Private Sub Command6_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox ("Data Yang Ingin Diedit Belum Siap"), vbCritical
Call bersih
Else
    Adodc1.Recordset.Update
    Adodc1.Recordset!Kode_Obat = Text1.Text
    Adodc1.Recordset!Nama_Obat = Text2.Text
    Adodc1.Recordset!satuan = Text3.Text
    Adodc1.Recordset!harga = Text4.Text
    Adodc1.Recordset!Stok = Text5.Text
        MsgBox "Data Berhasil Diubah", vbInformation, "INFO"
Call bersih
End If
End Sub

Private Sub Form_Load()
Set MyControl = DataGrid1
WheelHook DataGrid1

DataGrid1.HeadFont.Bold = True
DataGrid1.Columns(0).Width = 1500
DataGrid1.Columns(1).Width = 2500
DataGrid1.Columns(2).Width = 1300
End Sub

Private Sub Text1_Change()
If Len(Text1) > 5 Then
    MsgBox "Kode Obat Yang Anda Masukan Terlalu Banyak", vbCritical
        Text1.Text = ""
        Text1.SetFocus
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(Text2) = True Then
    MsgBox "Block Number", vbCritical
    Text2 = ""
Exit Sub
End If
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(Text3) = True Then
    MsgBox "Block Number", vbCritical
    Text3 = ""
Exit Sub
End If
    Text4.SetFocus
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(Text4) = False Then
    MsgBox "Block Huruf", vbCritical
    Text4 = ""
Exit Sub
End If
    Text5.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(Text5) = False Then
    MsgBox "Block Huruf", vbCritical
    Text5 = ""
Exit Sub
End If
    Command1.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(Text6) = True Then
    MsgBox "Block Number", vbCritical
    Text6 = ""
Exit Sub
End If
    Command5.SetFocus
End If
End Sub
