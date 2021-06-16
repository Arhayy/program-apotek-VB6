VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Data_Dokter 
   Caption         =   "Data Dokter"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
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
      RecordSource    =   "Data_Dokter"
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
      Bindings        =   "Data_Dokter.frx":0000
      Height          =   1935
      Left            =   0
      TabIndex        =   11
      Top             =   3720
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3413
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "ID_Dokter"
         Caption         =   "ID_Dokter"
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
         DataField       =   "Nama_Dokter"
         Caption         =   "Nama_Dokter"
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
      BeginProperty Column03 
         DataField       =   "No_Telp"
         Caption         =   "No_Telp"
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
            ColumnWidth     =   764,787
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1365,165
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "KELUAR"
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
      Left            =   4320
      TabIndex        =   10
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HAPUS"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SIMPAN"
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
      Left            =   2160
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
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
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3120
      Width           =   855
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
      Left            =   1080
      TabIndex        =   5
      Top             =   3120
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Masukan Nama Dokter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   6135
      Begin VB.Label Label5 
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
         Height          =   405
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
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
      Height          =   390
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   3135
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
      Height          =   420
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
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
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
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
      Height          =   410
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "Data_Dokter.frx":0015
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "Data_Dokter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text5 = "" Then
    MsgBox "Data Tidak Di Temukan", vbExclamation
Else
Call konekdb
RSdokter.CursorLocation = adUseClient
RSdokter.Open "SELECT * FROM Data_Dokter WHERE ID_Dokter like '%" & Text5 & "%'", konek
    If Not RSdokter.EOF Then
        With RSdokter
        With DataGrid1
            Set .DataSource = RSdokter
                Text5 = ""
            .Refresh
        End With
        End With
    End If
End If
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
    MsgBox ("Data Dokter Belum Lengkap"), vbCritical
Call bersih
Else
    Adodc1.Recordset.AddNew
    Adodc1.Recordset!ID_Dokter = Text1.Text
    Adodc1.Recordset!Nama_Dokter = Text2.Text
    Adodc1.Recordset!Alamat = Text3.Text
    Adodc1.Recordset!No_Telp = Text4.Text
        MsgBox "Data Berhasil Disimpan", vbInformation, "INFO"
    Adodc1.Recordset.Update
Call bersih
End If
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
    Data_Dokter.Hide
Call bersih
Else
End If
End Sub

Private Sub Form_Load()
Set MyControl = DataGrid1
WheelHook DataGrid1

Text1.Enabled = False
DataGrid1.Columns(0).Width = 1200
DataGrid1.Columns(1).Width = 2000
DataGrid1.Columns(2).Width = 1200
DataGrid1.HeadFont.Bold = True
End Sub

Sub bersih()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
End Sub

Sub otomatis()
Call konekdb
RSdokter.Open ("Select * from Data_Dokter Where ID_Dokter In(Select Max(ID_Dokter)From Data_Dokter)Order By ID_Dokter Desc"), konek
RSdokter.Requery
    Dim urutan As String * 5
    Dim hitung As Long
    With RSdokter
        If .EOF Then
            urutan = "DR" + "001"
            Text1.Text = urutan
        Else
            hitung = Right(!ID_Dokter, 3) + 1
            urutan = "DR" + Right("000" & hitung, 3)
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
    Call otomatis
    Text3.SetFocus
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text4.SetFocus
End If
End Sub

Private Sub Text4_Change()
If Len(Text4) > 13 Then
    MsgBox "Data Yang Anda Masukan Terlalu Banyak", vbCritical
        Text4.Text = ""
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
    Command2.SetFocus
End If
End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(Text5) = True Then
    MsgBox "Block Number", vbCritical
    Text5 = ""
Exit Sub
End If
    Command1.SetFocus
End If
End Sub
