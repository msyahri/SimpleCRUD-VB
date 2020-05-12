VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Login"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1185
      Left            =   5160
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1125
      ScaleWidth      =   1125
      TabIndex        =   10
      Top             =   360
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1185
      Left            =   720
      Picture         =   "Form1.frx":6FFE
      ScaleHeight     =   1125
      ScaleWidth      =   1125
      TabIndex        =   9
      Top             =   360
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "© RPL ~ SMK Ma'arif NU 1 Ajibarang"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   210
      Left            =   2040
      TabIndex        =   12
      Top             =   5400
      Width           =   2610
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Login untuk mengakses program"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   2160
      TabIndex        =   11
      Top             =   1200
      Width           =   2745
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Info Login"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3000
      TabIndex        =   8
      Top             =   1800
      Width           =   960
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   720
      Top             =   2040
      Width           =   5655
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Anda memiliki 3 kali kesempatan untuk memasukkan informasi login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   720
      TabIndex        =   7
      Top             =   4800
      Width           =   5685
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   570
      Left            =   3000
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Username  :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password   :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   4
      Top             =   3000
      Width           =   1470
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Login As Integer


Private Sub Command1_Click()
User = Text1.Text
Password = Text2.Text

If User = "manusa" And Password = "manusaajibarang" Then
MsgBox "Akses Diizinkan, Selamat Datang di Program Entry data Manusa", vbInformation, "Info"
Form2.Show
Form1.Hide
Else
Login = Login + 1
MsgBox "Informasi Login Salah Sebanyak " & Login & " kali", vbExclamation, "Info"
If Login = 2 Then
MsgBox "Kesempatan login tinggal 1 kali", vbExclamation, "Info"
End If
If Login = 3 Then
MsgBox "Anda sudah 3 kali salah memasukkan informasi login, program akan tertutup", vbCritical, "Info"
End
End If
End If

End Sub

Private Sub Command2_Click()
pesan = MsgBox("Anda Yakin Ingin Keluar ?", vbQuestion + vbYesNo, "Konfirmasi")
If pesan = vbYes Then
End
Else
Form1.SetFocus
End If
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

