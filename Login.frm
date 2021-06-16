VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H80000007&
   Caption         =   "Login"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   27.75
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   3
      Top             =   6600
      WhatsThisHelpID =   8950
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Masuk"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   2
      Top             =   6600
      WhatsThisHelpID =   8950
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   885
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5520
      Width           =   8295
   End
   Begin VB.TextBox Text1 
      Height          =   885
      Left            =   3240
      TabIndex        =   0
      Top             =   4560
      Width           =   8295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Daftar?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   1800
      TabIndex        =   4
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "Login.frx":0000
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim query As String

Private Sub Command1_Click()
Call konekdb
If Text1.Text = "" Then
    MsgBox "Username Tidak Boleh Kosong!!", vbCritical
    Text1.SetFocus
ElseIf Text2.Text = "" Then
    MsgBox "Password Tidak Boleh Kosong!!", vbCritical
    Text2.SetFocus
Else
    query = "select * from dblogin where Username='" & Text1.Text & "' and Password='" & Text2.Text & "'"
    RSadmin.Open (query), konek
        If RSadmin.EOF Then
            MsgBox "Username atau Password Salah !!", vbExclamation
            Text1.Text = ""
            Text2.Text = ""
            Text1.SetFocus
        Else
            Unload Me
            Menu.Show
            Menu.StatusBar1.Panels(1) = RSadmin!Status
            Menu.StatusBar1.Panels(2) = RSadmin!Nama
            Menu.StatusBar1.Panels(3) = Date
                If Menu.StatusBar1.Panels(1) <> "Admin" Then
                    Menu.menu_master.Enabled = False
                    Menu.menu_report.Enabled = False
                Else
                    Menu.menu_master.Enabled = True
                    Menu.menu_report.Enabled = True
                End If
        End If
End If
End Sub

Private Sub Command2_Click()
    warning = MsgBox("Apakah Anda Ingin Keluar!!", vbYesNo + vbInformation)
If warning = vbYes Then
End
Else
End If
End Sub

Private Sub Form_Load()
Login.Hide
End Sub

Private Sub IMAGE1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Label4.ForeColor = vbRed
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Label4.ForeColor = vbBlue
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Register.Show
Login.Hide
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(Text1) = True Then
    MsgBox "Block Number", vbCritical
    Text1 = ""
Exit Sub
End If
    Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus
End Sub

Private Sub Timer1_Timer()
Label1.ForeColor = RGB(Rnd * 250, Rnd * 250, Rnd * 250)
End Sub
