VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmmain 
   BackColor       =   &H00FF8080&
   Caption         =   "SISTEM INFORMASI TIMBANGAN - "
   ClientHeight    =   8160
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13230
   Icon            =   "frmsit30.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   3495
      Left            =   0
      ScaleHeight     =   3435
      ScaleWidth      =   13170
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   13230
      Begin VB.Image Image1 
         Height          =   12000
         Left            =   0
         Picture         =   "frmsit30.frx":57E2
         Top             =   0
         Width           =   21600
      End
   End
   Begin VB.Timer Timer3 
      Left            =   1080
      Top             =   4560
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   600
      Top             =   4560
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   4560
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7785
      Width           =   13230
      _ExtentX        =   23336
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2990
            MinWidth        =   2999
            Picture         =   "frmsit30.frx":3E036
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   16784
            MinWidth        =   2999
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2990
            MinWidth        =   2999
            Picture         =   "frmsit30.frx":3E13F
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mdiMenu 
      Caption         =   "&APLIKASI"
      Index           =   0
      Begin VB.Menu mnApp 
         Caption         =   "&Login"
         Index           =   0
      End
      Begin VB.Menu mnApp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnApp 
         Caption         =   "&Setting Timbangan"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnApp 
         Caption         =   "&Grup User"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnApp 
         Caption         =   "&User"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnApp 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnApp 
         Caption         =   "&Keluar"
         Index           =   6
      End
   End
   Begin VB.Menu mdiMenu 
      Caption         =   "&MASTER"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnMaster 
         Caption         =   "Data &Barang"
         Index           =   0
      End
      Begin VB.Menu mnMaster 
         Caption         =   "Data &Perusahaan"
         Index           =   1
      End
      Begin VB.Menu mnMaster 
         Caption         =   "Data &Truk"
         Index           =   2
      End
   End
   Begin VB.Menu mdiMenu 
      Caption         =   "&TRANSAKSI"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnTrans 
         Caption         =   "&Jadwal Timbang"
         Index           =   0
      End
      Begin VB.Menu mnTrans 
         Caption         =   "&Penimbangan"
         Index           =   1
      End
      Begin VB.Menu mnTrans 
         Caption         =   "&Maintenance Timbang"
         Index           =   2
      End
   End
   Begin VB.Menu mdiMenu 
      Caption         =   "&LAPORAN"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnLap 
         Caption         =   "Rekap Timbang (per &Nomer)"
         Index           =   0
      End
      Begin VB.Menu mnLap 
         Caption         =   "Rekap Timbang (per &Group)"
         Index           =   1
      End
      Begin VB.Menu mnLap 
         Caption         =   "Rekap Timbang (per &Tgl)"
         Index           =   2
      End
      Begin VB.Menu mnLap 
         Caption         =   "Rekap Timbang (per &Jam)"
         Index           =   3
      End
   End
   Begin VB.Menu mdiMenu 
      Caption         =   "&UTILITY"
      Index           =   4
      Visible         =   0   'False
      Begin VB.Menu mnUtil 
         Caption         =   "&Import Jadwal Timbang (PELINDO)"
         Index           =   0
      End
      Begin VB.Menu mnUtil 
         Caption         =   "&Monitoring Timbang"
         Index           =   1
      End
      Begin VB.Menu mnUtil 
         Caption         =   "&Sinkronisasi Data"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private m_BGnd As CMdiBackground
Dim fEncrypt As New Encrypt
Dim licpath As String

Private Sub MDIForm_Activate()
  
 If Not appon Then
 
    With frmlogin
        If Logon Then Logon = True
  'Menu_Visible False, False
        frmlogin.Visible = True
    End With
        frmCekTgl.Show 1
 End If
 appon = True

 End Sub

Private Sub MDIForm_Load()
 
 ' Setup background
   Set m_BGnd = New CMdiBackground
   Set m_BGnd.Graphic = Image1
   With m_BGnd
      Set .Client = Me
       .GraphicPosition = mdiStretched
       .AutoRefresh = True
    End With
    
  
 'Loadunload me endif Data Port Timbangan ---------------
  sPathData = App.Path & "\data.dat"
  If FileExists(sPathData) Then
   ncom = FileRead(sPathData, True)(1)
   baudrate = FileRead(sPathData, True)(2)
   parity = FileRead(sPathData, True)(3)
   databit = FileRead(sPathData, True)(4)
   stopbit = FileRead(sPathData, True)(5)
  Else
   ncom = -1
  End If
 '----------------------------------------
 End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
 End
End Sub

Private Sub mnApp_Click(Index As Integer)
 Select Case Index
 Case 0 'Log Out
  If Logon Then Logon = False
  Menu_Visible False, False
  frmlogin.Show 1
 Case 2 'Set Data Port Timbangan
  frmsetport.Show 1
 Case 3 'Set Grup User dan Menu
  frmGrupUser.Show
 Case 4 'Set User
  frmUser.Show
 Case 6 'Keluar Aplikasi
  End
 End Select
End Sub

Private Sub mnLap_Click(Index As Integer)
 Select Case Index
 Case 0     'Rekap Timbang per Nomer
  frmLapRekap.Show 1
 Case 1     'Rekap Timbang per Group
  frmLapRekapGrp.Show 1
 Case 2     'Rekap Timbang per Tgl
  frmLapRekapTgl.Show 1
 Case 3     'Rekap Timbang per Jam
  frmLapRekapJam.Show 1
 End Select
End Sub

Private Sub mnMaster_Click(Index As Integer)
 Select Case Index
 Case 0 'Data Barang
  frmBarang.Show
 Case 1 'Data Perusahaan
  frmPerusahaan.Show
 Case 2 'Data Angkutan
  frmTruk.Show
 End Select
End Sub

Private Sub mnTrans_Click(Index As Integer)
 Select Case Index
 Case 0 'Jadwal Timbang
  frmjadwal.Show
 Case 1 'Penimbangan
  frmtimbang.Show
 Case 2 'Maintenance
  frmMtimbang.Show
 End Select
End Sub

Private Sub mnUtil_Click(Index As Integer)
 Select Case Index
 Case 0     'Import Jadwal Timbang
  frmkjadwal.Show 1
 Case 1     'Montoring Timbang
  frmmontim.Show
 Case 2     'Sinkronisasi Data
  frmsinkron.Show 1
 End Select
End Sub

Private Sub Timer2_Timer()
 Dim licdate As String, liccount As String
 
 'Tampilkan Tanggal dan Jam
 'Pada Status Bar Form Utama ------------
  sb.Panels(3).Text = Date & "  " & Time
 '---------------------------------------
 
  licpath = App.Path & "\lic.gmd"
  If FileExists(licpath) Then
   licdate = FileRead(licpath, True)(1)
   licdate = fEncrypt.ChgPass(licdate, 2)
   liccount = FileRead(licpath, True)(2)
   liccount = fEncrypt.ChgPass(liccount, 2)
   
   'liccount = Val(liccount) - 1
   If licdate <> Format(Date, "dd/mm/yyyy") And _
     Val(liccount) > 0 Then
    liccount = Val(liccount) - 1
    
    If liccount > 0 Then
     FileWriteBinary fEncrypt.ChgPass(Format(Date, "dd/mm/yyyy"), 1), App.Path & "\lic.gmd", False
     FileWriteBinary fEncrypt.ChgPass(liccount, 1), App.Path & "\lic.gmd", True
    Else
     MsgBox "Program Expired !!!!" & vbCrLf & _
            "Hubungi Administrator" & vbCrLf & _
            "Untuk Melakukan Aktivasi Kembali", vbOKOnly + vbInformation, "ADMINISTRATOR"
     End
    End If
   
   Else
    If Val(liccount) = 0 Then
     MsgBox "Program Expired !!!!" & vbCrLf & _
            "Hubungi Administrator" & vbCrLf & _
            "Untuk Melakukan Aktivasi Kembali", vbOKOnly + vbInformation, "ADMINISTRATOR"
     End
    End If
   End If
  Else
   'FileWriteBinary fEncrypt.ChgPass(Format(Date, "dd/mm/yyyy"), 1), App.Path & "\lic.gmd"
   'FileWriteBinary fEncrypt.ChgPass("32", 1), App.Path & "\lic.gmd", True
   MsgBox "Program Expired !!!!" & vbCrLf & _
          "Hubungi Administrator" & vbCrLf & _
          "Untuk Melakukan Aktivasi Kembali", vbOKOnly + vbInformation, "ADMINISTRATOR"
   End
  End If
End Sub
