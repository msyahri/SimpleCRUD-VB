VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7230
   LinkTopic       =   "Form3"
   ScaleHeight     =   6375
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton tabout 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   11
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Rekayasa Perangkat Lunak @ SMK Ma'arif NU 1 AJIBARANG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1080
      TabIndex        =   10
      Top             =   4680
      Width           =   5100
   End
   Begin VB.Label aboutsa 
      Caption         =   "keisamega@gmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label aboutw 
      Caption         =   "facebook.com/wakhyu.shaputra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label aboutse 
      Caption         =   "facebook.com/setiawan.ownskin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label aboutms 
      Caption         =   "about.me/m.syahri"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Wahyu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1560
      TabIndex        =   5
      Top             =   3960
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Setiawan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1560
      TabIndex        =   4
      Top             =   3480
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Sartun"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1560
      TabIndex        =   3
      Top             =   3000
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "M. Syahri"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1560
      TabIndex        =   2
      Top             =   2520
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"Form3.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TENTANG"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
    As Long) As Long

Private Sub aboutms_Click()
    ShellExecute hwnd, "open", "https://about.me/msyahri/", _
    vbNullString, vbNullString, 1
End Sub

Private Sub aboutsa_Click()
    ShellExecute hwnd, "open", "mailto:keisamega@gmail.com", _
    vbNullString, vbNullString, 1
End Sub


Private Sub aboutse_Click()
    ShellExecute hwnd, "open", "https://facebook.com/setiawan.ownskin/", _
    vbNullString, vbNullString, 1
End Sub

Private Sub aboutw_Click()
    ShellExecute hwnd, "open", "https://facebook.com/wakhyu.shaputra/", _
    vbNullString, vbNullString, 1
End Sub

Private Sub tabout_Click()
Form3.Hide
Form2.Show
End Sub
