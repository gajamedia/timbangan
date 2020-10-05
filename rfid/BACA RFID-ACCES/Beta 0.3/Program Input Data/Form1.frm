VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   ".: INPUT DATA :."
   ClientHeight    =   9045
   ClientLeft      =   2460
   ClientTop       =   1575
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   11175
   Begin VB.TextBox TOTAL_QUOTA 
      Height          =   375
      Left            =   6600
      TabIndex        =   40
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox TAMBAH_QUOTA 
      Height          =   375
      Left            =   6600
      TabIndex        =   39
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox SHIFT 
      Height          =   375
      Left            =   6600
      TabIndex        =   38
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox BAHAN_BAKAR 
      Height          =   375
      Left            =   6600
      TabIndex        =   37
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox DEPARTEMEN 
      Height          =   375
      Left            =   6600
      TabIndex        =   36
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox NO_PINTU 
      Height          =   375
      Left            =   2280
      TabIndex        =   35
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox JENIS 
      Height          =   375
      Left            =   2280
      TabIndex        =   34
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox TIPE_GARDAN 
      Height          =   375
      Left            =   2280
      TabIndex        =   33
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox TIPE_UNIT 
      Height          =   375
      Left            =   2280
      TabIndex        =   32
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox SUPLIER 
      Height          =   375
      Left            =   2280
      TabIndex        =   31
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox RFID 
      Height          =   375
      Left            =   2280
      TabIndex        =   15
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox NO_POLISI 
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox KETERANGAN 
      Height          =   375
      Left            =   6600
      TabIndex        =   13
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox QUOTA 
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Cmdtambah 
      Caption         =   "&TAMBAH"
      Height          =   375
      Left            =   9480
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Cmdsimpan 
      Caption         =   "&SIMPAN"
      Height          =   375
      Left            =   9480
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Cmdedit 
      Caption         =   "&UBAH"
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Cmdhapus 
      Caption         =   "&HAPUS"
      Height          =   375
      Left            =   9480
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Cmdbatal 
      Caption         =   "&BATAL"
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Cmdpertama 
      Caption         =   "PERTAMA"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Cmdsebelum 
      Caption         =   "SEBELUM"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Cmdberikut 
      Caption         =   "BERIKUT"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Cmdterakhir 
      Caption         =   "TERAKHIR"
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Cmdkeluar 
      BackColor       =   &H80000012&
      Caption         =   "AKTIVITAS"
      Height          =   375
      Left            =   9600
      MaskColor       =   &H80000010&
      TabIndex        =   2
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Cmdcari 
      Caption         =   "CARI"
      Height          =   375
      Left            =   9480
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Cmdtampil 
      Caption         =   "REFRESH"
      Height          =   375
      Left            =   9480
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8400
      Top             =   8520
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   7440
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAMA\Desktop\DATABASE.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAMA\Desktop\DATABASE.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"Form1.frx":0000
      Caption         =   "DATABASE RFID"
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
      Bindings        =   "Form1.frx":00AE
      Height          =   3255
      Left            =   480
      TabIndex        =   16
      Top             =   3960
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5741
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
      Caption         =   "TABEL_RFID"
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
            LCID            =   1057
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
            LCID            =   1057
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
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9000
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label14 
      Caption         =   "TOTAL_AKTUAL"
      Height          =   375
      Left            =   4920
      TabIndex        =   30
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "TAMBAH_QUOTA"
      Height          =   375
      Left            =   4920
      TabIndex        =   29
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label12 
      Caption         =   "BAHAN_BAKAR"
      Height          =   375
      Left            =   4920
      TabIndex        =   28
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "SHIFT"
      Height          =   375
      Left            =   4920
      TabIndex        =   27
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "DEPARTEMEN"
      Height          =   375
      Left            =   4920
      TabIndex        =   26
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "NO_PINTU"
      Height          =   375
      Left            =   600
      TabIndex        =   25
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "JENIS"
      Height          =   375
      Left            =   600
      TabIndex        =   24
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "TIPE_GARDAN"
      Height          =   375
      Left            =   600
      TabIndex        =   23
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "TIPE_UNIT"
      Height          =   375
      Left            =   600
      TabIndex        =   22
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "SUPLIER"
      Height          =   375
      Left            =   600
      TabIndex        =   21
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "RFID"
      Height          =   375
      Left            =   600
      TabIndex        =   20
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "NO_POLISI"
      Height          =   375
      Left            =   600
      TabIndex        =   19
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "KETERANGAN"
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "QUOTA"
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Baru As Boolean
Dim TERIMA_RFID As String


Private Sub cmdbatal_Click()
    Tombol True, True, False, False, True
    Kotak False, False, False, False, False, False, False, False, False, False, False, False, False, False
    Abu
    Adodc1.Recordset.Cancel
    Kosong
End Sub


Private Sub cmdberikut_Click()
    'Menuju ke record berikutnya
    Adodc1.Recordset.MoveNext
    'Jika berada di record terakhir menuju ke record terakhir
    If Adodc1.Recordset.EOF Then
        Adodc1.Recordset.MoveLast
    End If
End Sub


Private Sub cmdcari_Click()
Kriteria = InputBox("Masukkan NO_POLISI : ", "Mencari Data")
Adodc1.RecordSource = "SELECT RFID,SUPLIER,TIPE_UNIT,TIPE_GARDAN,JENIS,NO_POLISI,NO_PINTU,DEPARTEMEN,BAHAN_BAKAR,SHIFT,QUOTA,TAMBAH_QUOTA,TOTAL_QUOTA,KETERANGAN FROM TABEL_RFID Where NO_POLISI Like'" & "%" & Kriteria & "%" & "'"
Adodc1.Refresh

If Adodc1.Recordset.EOF Then
    MsgBox "Data Tidak Ditemukan!", vbCritical, "Data Tidak Ada"
End If
End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cmdedit_Click()
    Tombol False, False, True, True, False
    Kotak True, True, True, True, True, True, True, True, True, True, True, True, True, True
    
    Putih
    Timer1.Enabled = True
       
    With Adodc1.Recordset
        
        RFID.Text = .Fields("RFID")
        SUPLIER.Text = .Fields("SUPLIER")
        TIPE_UNIT.Text = .Fields("TIPE_UNIT")
        TIPE_GARDAN.Text = .Fields("TIPE_GARDAN")
        JENIS.Text = .Fields("JENIS")
        NO_POLISI.Text = .Fields("NO_POLISI")
        NO_PINTU.Text = .Fields("NO_PINTU")
        DEPARTEMEN.Text = .Fields("DEPARTEMEN")
        BAHAN_BAKAR.Text = .Fields("BAHAN_BAKAR")
        SHIFT.Text = .Fields("SHIFT")
        QUOTA.Text = .Fields("QUOTA")
        TAMBAH_QUOTA.Text = .Fields("TAMBAH_QUOTA")
        TOTAL_QUOTA.Text = .Fields("TOTAL_QUOTA")
        KETERANGAN.Text = .Fields("KETERANGAN")

    End With
    RFID.SetFocus
    Baru = False
End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cmdhapus_Click()
    Dim hapus
    Kotak False, False, False, False, False, False, False, False, False, False, False, False, False, False
        
    Abu
    hapus = MsgBox("Anda yakin data ini akan dihapus?", vbQuestion + vbYesNo, "Hapus Data")
    If hapus = vbYes Then
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveLast
    Else
        MsgBox "Data tidak jadi dihapus!", vbOKOnly + vbInformation, "Batal Menghapus"
    End If
End Sub


Private Sub cmdkeluar_Click()
    
    'Form2.Visible = True
    'Form2.Enabled = True
    Form1.Visible = False
    'Form1.Enabled = False
    'Form2.Timer2.Enabled = True
    
End Sub


Private Sub cmdpertama_Click()
    'Menuju ke record pertama
    Adodc1.Recordset.MoveFirst
End Sub


Private Sub cmdsebelum_Click()
    'Menuju ke record sebelumnya
    Adodc1.Recordset.MovePrevious
    'Jika berada di record pertama menuju ke record pertama
    If Adodc1.Recordset.BOF Then
        Adodc1.Recordset.MoveFirst
    End If
End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cmdsimpan_Click()
    Tombol True, True, False, False, True
    Kotak False, False, False, False, False, False, False, False, False, False, False, False, False, False
    
    Abu
    
    With Adodc1.Recordset
        If Baru Then
            .AddNew
            
            .Fields("RFID") = RFID.Text
            .Fields("SUPLIER") = SUPLIER.Text
            .Fields("TIPE_UNIT") = TIPE_UNIT.Text
            .Fields("TIPE_GARDAN") = TIPE_GARDAN.Text
            .Fields("JENIS") = JENIS.Text
            .Fields("NO_POLISI") = NO_POLISI.Text
            .Fields("NO_PINTU") = NO_PINTU.Text
            .Fields("DEPARTEMEN") = DEPARTEMEN.Text
            .Fields("BAHAN_BAKAR") = BAHAN_BAKAR.Text
            .Fields("SHIFT") = SHIFT.Text
            .Fields("QUOTA") = QUOTA.Text
            .Fields("TAMBAH_QUOTA") = TAMBAH_QUOTA.Text
            .Fields("TOTAL_QUOTA") = TOTAL_QUOTA.Text
            .Fields("KETERANGAN") = KETERANGAN.Text
            
            .Update
            .Sort = "SUPLIER"
        Else
            
            .Fields("RFID") = RFID.Text
            .Fields("SUPLIER") = SUPLIER.Text
            .Fields("TIPE_UNIT") = TIPE_UNIT.Text
            .Fields("TIPE_GARDAN") = TIPE_GARDAN.Text
            .Fields("JENIS") = JENIS.Text
            .Fields("NO_POLISI") = NO_POLISI.Text
            .Fields("NO_PINTU") = NO_PINTU.Text
            .Fields("DEPARTEMEN") = DEPARTEMEN.Text
            .Fields("BAHAN_BAKAR") = BAHAN_BAKAR.Text
            .Fields("SHIFT") = SHIFT.Text
            .Fields("QUOTA") = QUOTA.Text
            .Fields("TAMBAH_QUOTA") = TAMBAH_QUOTA.Text
            .Fields("TOTAL_QUOTA") = TOTAL_QUOTA.Text
            .Fields("KETERANGAN") = KETERANGAN.Text
            
            .Update
            .Sort = "SUPLIER"
            
        End If
    End With
    Baru = False
    
    Timer1.Enabled = False
    
    'Adodc1.Refresh
    'cmdtampil_Click
    Kosong
End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cmdtambah_Click()
    Kotak True, True, True, True, True, True, True, True, True, True, True, True, True, True
    Tombol False, False, True, True, False
    
    Putih
    Form1.Timer1.Enabled = True
    
    Baru = True
    Kosong
    RFID.Text = "INPUTKAN RFID"
    RFID.SetFocus
End Sub


Private Sub cmdtampil_Click()
    'Adodc1.RecordSource = "SELECT RFID,NO_POLISI,QUOTA_BBM FROM TABEL_RFID Where NO_POLISI Like'" & "%" & Kriteria & "%" & "' ORDER BY RFID"
    Adodc1.Refresh
    Adodc1.Refresh
End Sub


Private Sub cmdterakhir_Click()
    Adodc1.Recordset.MoveLast
End Sub


Public Sub Kotak(X1, X2, X3, X4, X5, X6, X7, X8, X9, X10, X11, X12, X13, X14 As Boolean)
    
    RFID.Enabled = X1
    SUPLIER.Enabled = X2
    TIPE_UNIT.Enabled = X3
    TIPE_GARDAN.Enabled = X4
    JENIS.Enabled = X5
    NO_POLISI.Enabled = X6
    NO_PINTU.Enabled = X7
    DEPARTEMEN.Enabled = X8
    BAHAN_BAKAR.Enabled = X9
    SHIFT.Enabled = X10
    QUOTA.Enabled = X11
    TAMBAH_QUOTA.Enabled = X12
    TOTAL_QUOTA.Enabled = X13
    KETERANGAN.Enabled = X14
    
    
End Sub


Public Sub Tombol(tambah, EDIT, simpan, batal, hapus As Boolean)
    Cmdtambah.Enabled = tambah
    Cmdedit.Enabled = EDIT
    Cmdsimpan.Enabled = simpan
    Cmdbatal.Enabled = batal
    Cmdhapus.Enabled = hapus
End Sub


Private Sub Form_Load()
        
    Timer1.Enabled = False
    MSComm1.PortOpen = True
    
    'Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\MYANOM\Desktop\DATABASE.mdb;Persist Security Info=False"
    'Adodc1.CommandType = adCmdText
    'Adodc1.RecordSource = "SELECT RFID,NO_POLISI,QUOTA_BBM FROM TABEL_RFID ORDER BY RFID"
    Adodc1.Refresh
    
    Baru = False
    Tombol True, True, False, False, True
    Kotak False, False, False, False, False, False, False, False, False, False, False, False, False, False
    Abu
    
End Sub


Public Sub Kosong()
    RFID.Text = "0"
    SUPLIER.Text = "0"
    TIPE_UNIT.Text = "0"
    TIPE_GARDAN.Text = "0"
    JENIS.Text = "0"
    NO_POLISI.Text = "0"
    NO_PINTU.Text = "0"
    DEPARTEMEN.Text = "0"
    BAHAN_BAKAR.Text = "0"
    SHIFT.Text = "0"
    QUOTA.Text = "0"
    TAMBAH_QUOTA.Text = "0"
    TOTAL_QUOTA.Text = "0"
    KETERANGAN.Text = "0"
End Sub


Public Sub Putih()
    RFID.BackColor = &H80000005
    SUPLIER.BackColor = &H80000005
    TIPE_UNIT.BackColor = &H80000005
    TIPE_GARDAN.BackColor = &H80000005
    JENIS.BackColor = &H80000005
    NO_POLISI.BackColor = &H80000005
    NO_PINTU.BackColor = &H80000005
    DEPARTEMEN.BackColor = &H80000005
    BAHAN_BAKAR.BackColor = &H80000005
    SHIFT.BackColor = &H80000005
    QUOTA.BackColor = &H80000005
    TAMBAH_QUOTA.BackColor = &H80000005
    TOTAL_QUOTA.BackColor = &H80000005
    KETERANGAN.BackColor = &H80000005
End Sub


Public Sub Abu()
    RFID.BackColor = &H80000000
    SUPLIER.BackColor = &H80000000
    TIPE_UNIT.BackColor = &H80000000
    TIPE_GARDAN.BackColor = &H80000000
    JENIS.BackColor = &H80000000
    NO_POLISI.BackColor = &H80000000
    NO_PINTU.BackColor = &H80000000
    DEPARTEMEN.BackColor = &H80000000
    BAHAN_BAKAR.BackColor = &H80000000
    SHIFT.BackColor = &H80000000
    QUOTA.BackColor = &H80000000
    TAMBAH_QUOTA.BackColor = &H80000000
    TOTAL_QUOTA.BackColor = &H80000000
    KETERANGAN.BackColor = &H80000000
End Sub


Private Sub Timer1_Timer()
        
    TERIMA_RFID = MSComm1.Input
    TERIMA_RFID = Left(TERIMA_RFID, 9)
    
    If TERIMA_RFID = "" Then
        'Text1.Text = ""
    Else
        RFID.Text = TERIMA_RFID
    End If
    
End Sub

