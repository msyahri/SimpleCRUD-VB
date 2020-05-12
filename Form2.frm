VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Utama"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10485
   LinkTopic       =   "Form2"
   ScaleHeight     =   8835
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Tentang"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      TabIndex        =   29
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox xtgl_lahir 
      Height          =   375
      Left            =   3240
      TabIndex        =   26
      ToolTipText     =   "Masukkan Tanggal Lahir"
      Top             =   4680
      Width           =   4335
   End
   Begin VB.CommandButton Trefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   23
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Tcari 
      Caption         =   "Cari Nama"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   22
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox xcari 
      Height          =   375
      Left            =   7200
      TabIndex        =   21
      ToolTipText     =   "Masukkan Nomor Telepon"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox xtelepon 
      Height          =   375
      Left            =   3240
      TabIndex        =   20
      ToolTipText     =   "Masukkan Nomor Telepon"
      Top             =   4080
      Width           =   4335
   End
   Begin VB.TextBox xalamat 
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      ToolTipText     =   "Masukkan Alamat"
      Top             =   3480
      Width           =   4335
   End
   Begin VB.TextBox xnama 
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      ToolTipText     =   "Masukkan Nama"
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox xnis 
      Height          =   375
      Left            =   3240
      TabIndex        =   17
      ToolTipText     =   "Masukkan NIS"
      Top             =   1080
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   8520
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
      Bindings        =   "Form2.frx":0000
      Height          =   2295
      Left            =   480
      TabIndex        =   16
      Top             =   6120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1740
      Left            =   8280
      Picture         =   "Form2.frx":0015
      ScaleHeight     =   1680
      ScaleWidth      =   1680
      TabIndex        =   15
      Top             =   720
      Width           =   1740
   End
   Begin VB.ComboBox cmbjurusan 
      Height          =   315
      ItemData        =   "Form2.frx":855F
      Left            =   3240
      List            =   "Form2.frx":8575
      TabIndex        =   5
      Text            =   "Pilih Jurusan"
      Top             =   2280
      Width           =   4335
   End
   Begin VB.ComboBox cmbjk 
      Height          =   315
      ItemData        =   "Form2.frx":8597
      Left            =   3240
      List            =   "Form2.frx":85A1
      TabIndex        =   4
      Text            =   "Jenis Kelamin"
      Top             =   2880
      Width           =   4335
   End
   Begin VB.CommandButton Ttambah 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Tedit 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Tkeluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Thapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Contoh : 01/01/2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   3240
      TabIndex        =   28
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Tanggal Lahir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   27
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "© 2018 RPL ~ MANUSA"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   330
      Left            =   7920
      TabIndex        =   25
      Top             =   8520
      Width           =   2130
   End
   Begin VB.Label Note 
      AutoSize        =   -1  'True
      Caption         =   "*) Jika tombol Tambah tidak aktif, klik refresh untuk mengaktifkan kembali"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1680
      TabIndex        =   24
      Top             =   5760
      Width           =   5160
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Aksi"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   8880
      TabIndex        =   14
      Top             =   2520
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Input Data"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3960
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor Induk Siswa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   12
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Lengkap Siswa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Jurusan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Jenis Kelamin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Nomor Telepon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   2775
      Left            =   8280
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Judul 
      AutoSize        =   -1  'True
      Caption         =   "Entry Data Siswa SMK MA'ARIF NU 1 AJIBARANG"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   5445
   End
   Begin VB.Shape Shape2 
      Height          =   4695
      Left            =   480
      Top             =   840
      Width           =   7455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Form3.Show
End Sub

Private Sub DataGrid1_Click()
xnis.Text = DataGrid1.Columns(0).Text
xnama.Text = DataGrid1.Columns(1).Text
cmbjurusan.Text = DataGrid1.Columns(2).Text
cmbjk.Text = DataGrid1.Columns(3).Text
xalamat.Text = DataGrid1.Columns(4).Text
xtelepon.Text = DataGrid1.Columns(5).Text
xtgl_lahir.Text = DataGrid1.Columns(6).Text
Ttambah.Enabled = False
End Sub

 Private Sub Form_Load()
koneksi
Adodc1.ConnectionString = Conn.ConnectionString
Adodc1.RecordSource = "select * from informasisiswa"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

 Private Sub tcari_Click()
Adodc1.Recordset.Filter = "Nama = '" & xcari.Text & "'"
If xcari = "" Then
MsgBox "Masukkan nama terlebih dahulu!", vbExclamation, "Hasil"
Else
If Adodc1.Recordset.RecordCount > 0 Then
xnis = Adodc1.Recordset.Fields(0)
xnama = Adodc1.Recordset.Fields(1)
cmbjurusan = Adodc1.Recordset.Fields(2)
cmbjk = Adodc1.Recordset.Fields(3)
xalamat = Adodc1.Recordset.Fields(4)
xtelepon = Adodc1.Recordset.Fields(5)
xtgl_lahir = Adodc1.Recordset.Fields(6)
MsgBox "Data ditemukan.", vbInformation, "Hasil"
Else
MsgBox "Data tidak ditemukan", vbExclamation, "Hasil"
Adodc1.Refresh
End If
End If
End Sub


Private Sub Tedit_Click()
If xnis.Text = "" Or xnama.Text = "" Or cmbjurusan.Text = "" Or cmbjk.Text = "" Or xalamat.Text = "" Or xtelepon.Text = "" Then
MsgBox " Pilih Data Terlebih Dahulu ", vbInformation, "Info"
Else

Adodc1.Recordset!nis = xnis.Text
Adodc1.Recordset!nama = xnama.Text
Adodc1.Recordset!jurusan = cmbjurusan.Text
Adodc1.Recordset!jk = cmbjk.Text
Adodc1.Recordset!alamat = xalamat.Text
Adodc1.Recordset!telepon = xtelepon.Text
Adodc1.Recordset!tgl_lahir = xtgl_lahir.Text
MsgBox "Data Berhasil Diubah", vbInformation, "Info"
DataGrid1.Refresh

End If
End Sub

Private Sub Thapus_Click()
If xnis.Text = "" Or xnama.Text = "" Or cmbjurusan.Text = "" Or cmbjk.Text = "" Or xalamat.Text = "" Or xtelepon.Text = "" Or xtgl_lahir.Text = "" Then
MsgBox " Pilih Data Terlebih Dahulu ", vbInformation, "Info"
Else
pesan = MsgBox("Hapus Data ?", vbQuestion + vbYesNo, "Konfirmasi")
End If
If pesan = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Data Telah Dihapus", vbInformation, "Info"
Else
Form2.SetFocus
End If
End Sub

Private Sub Tkeluar_Click()
pesan = MsgBox("Anda Yakin Ingin Keluar dari program ?", vbQuestion + vbYesNo, "Question")
If pesan = vbYes Then
End
Else
Form2.SetFocus
End If
End Sub


Private Sub Trefresh_Click()
Adodc1.Refresh
Ttambah.Enabled = True
xnis.Text = ""
xnama.Text = ""
cmbjurusan.Text = "Pilih Jurusan"
cmbjk.Text = "Jenis Kelamin"
xalamat.Text = ""
xtelepon.Text = ""
xtgl_lahir.Text = ""
End Sub

Private Sub Ttambah_Click()
If xnis.Text = "" Or xnama.Text = "" Or cmbjurusan.Text = "Pilih Jurusan" Or cmbjk.Text = "Jenis Kelamin" Or xalamat.Text = "" Or xtelepon.Text = "" Or xtgl_lahir.Text = "" Then
MsgBox " Data Belum Lengkap ", vbInformation, "Info"
Else

Adodc1.Recordset.AddNew
Adodc1.Recordset!nis = xnis.Text
Adodc1.Recordset!nama = xnama.Text
Adodc1.Recordset!jurusan = cmbjurusan.Text
Adodc1.Recordset!jk = cmbjk.Text
Adodc1.Recordset!alamat = xalamat.Text
Adodc1.Recordset!telepon = xtelepon.Text
Adodc1.Recordset!tgl_lahir = xtgl_lahir.Text
MsgBox "Data Berhasil Ditambahkan", vbInformation, "Info"
DataGrid1.Refresh

End If
End Sub


