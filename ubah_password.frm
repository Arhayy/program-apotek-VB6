VERSION 5.00
Begin VB.Form ubah_password 
   Caption         =   "Ganti Password"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3855
      Begin VB.TextBox Text3 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Password Baru"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Konfirmasi"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.TextBox Text2 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Password Lama"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Pengguna"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ubah_password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1 <> Menu.StatusBar1.Panels(2) Then
        MsgBox "Anda Tidak Berhak Mengganti Password", vbCritical
        Exit Sub
    Else
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call konekdb
        RSadmin.Open "Select * from dblogin where Nama='" & Text1 & "' and Password='" & Text2 & "'", konek
        If Not RSadmin.EOF Then
            Text3.SetFocus
        Else
            MsgBox "Password Salah", vbCritical
            Exit Sub
        End If
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text4 <> Text3 Then
        MsgBox "Password Konfirmasi Tidak Sesuai", vbCritical
        Exit Sub
    Else
        ubah = "update dblogin set [Password]='" & Text3 & "' where Nama='" & Menu.StatusBar1.Panels(2) & "'"
        konek.Execute ubah
            MsgBox "Password Berhasil Diubah", vbInformation, "INFO"
        Unload Me
    End If
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text3 = Text2 Then
        MsgBox "Ganti Dengan Password Yang Berbeda"
        Exit Sub
    Else
        Text4.SetFocus
    End If
End If
End Sub
