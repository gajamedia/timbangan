VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0B63DF3F-CC00-4D55-A1C9-CAFE70BB1B49}#1.0#0"; "xpctrl.ocx"
Object = "{CC696B60-4159-11D0-BDCB-0020A90B183A}#8.0#0"; "PVDateEdit9.ocx"
Begin VB.Form frmsinkron 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sinkronisasi Data"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pgb 
      Height          =   135
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbData 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   2535
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6350
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin XPCtrl.XPButton cmdBtn 
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cek Data"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmsinkron.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   0
         ItemData        =   "frmsinkron.frx":001C
         Left            =   1560
         List            =   "frmsinkron.frx":006B
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   780
         Width           =   1215
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   1
         ItemData        =   "frmsinkron.frx":0169
         Left            =   3240
         List            =   "frmsinkron.frx":01B8
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   780
         Width           =   1215
      End
      Begin PVDATE2Lib.PVDate2 pvDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   1455
         _Version        =   524288
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarFormat  =   4
         DateFormat      =   13
         Value           =   41522.7182175926
      End
      Begin XPCtrl.XPButton cmdBtn 
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Proses"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmsinkron.frx":02B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "sd"
         Height          =   195
         Index           =   4
         Left            =   2925
         TabIndex        =   8
         Top             =   840
         Width           =   165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TGL. TIMBANG"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JAM TIMBANG"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmsinkron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cons As New ADODB.Connection
Dim cmdOra As New ADODB.Command
Dim rsOra As New ADODB.Recordset

Dim con As New ADODB.Connection
Dim rsFind As New ADODB.Recordset
Dim rsFind2 As New ADODB.Recordset

Option Explicit

Private Sub cmdBtn_Click(Index As Integer)
  Dim sqlStr As String, cArmada As String
  Dim i As Integer, j As Integer
  
  Select Case Index
  Case 0    'Cek Data Jasatama vs Pelindo
   If cmbData(1).Text >= cmbData(0).Text Then
    'cek server jasatama
    sqlStr = "select * from tbtrans " & _
             "where (date(wkeluar)='" & Format(pvDate.Value, "yyyy-mm-dd") & _
             "' and (time(wkeluar)>='" & Trim(cmbData(0).Text) & _
             "' and time(wkeluar)<='" & Trim(cmbData(1).Text) & "')) and " & _
             " (nodermaga<>'MX' or nodermaga<>'LB' or nodermaga<>'0') and" & _
             " nomor_1b<>''"
    Set rsFind = con.Execute(sqlStr)
    i = rsFind.RecordCount
    Set rsFind = Nothing
    
    'cek server pelindo
    sqlStr = "select * from rf_truk_transaksi " & _
             "where to_char(tgl_keluar,'DD/MM/YYYY') = '" & Format(pvDate.Value, "dd/mm/yyyy") & _
             "' and (to_char(tgl_keluar,'HH24:MI:SS')>='" & Trim(cmbData(0).Text) & _
             "' and to_char(tgl_keluar,'HH24:MI:SS')<='" & Trim(cmbData(1).Text) & "')"
    
    Set rsFind = cons.Execute(sqlStr)
    j = rsFind.RecordCount
    Set rsFind = Nothing
    
    stbData.Panels(1).Text = "JASATAMA: " & Str(i) & " || PELINDO: " & Str(j)
    
    'If i <> j Then cmdBtn(1).Visible = True Else cmdBtn(1).Visible = False
   Else
    MsgBox "Pastikan Pengisian Jam Benar" & vbCrLf & _
           "Jam Awal harus lebih kecil nilainya dari Jam Akhir", _
           vbOKOnly + vbInformation, "JASATAMA"
   End If

  Case 1    'Proses sinkronisasi
   'cmdBtn(1).Visible = False

   sqlStr = "select * from tbtrans " & _
            "where (date(wkeluar)='" & Format(pvDate.Value, "yyyy-mm-dd") & _
            "' and (time(wkeluar)>='" & Trim(cmbData(0).Text) & _
            "' and time(wkeluar)<='" & Trim(cmbData(1).Text) & "')) and " & _
            " (nodermaga<>'MX' or nodermaga<>'LB' or nodermaga<>'0') and" & _
            " nomor_1b<>''"
   Set rsFind = con.Execute(sqlStr)
   pgb.Min = 0
   pgb.Max = rsFind.RecordCount
   pgb.Visible = True
   pgb.Value = 0
   While Not rsFind.EOF
    stbData.Panels(2).Text = "Proses............."
    DoEvents
        
    sqlStr = "select * from rf_truk_transaksi " & _
             "where no_dermaga='" & rsFind("nodermaga").Value & _
             "' and no_lambung='" & rsFind("nolambung").Value & _
             "' and nomer_gjt='" & rsFind("nomer").Value & _
             "' and to_char(tgl_masuk,'YYYY-MM-DD HH24:MI:SS')='" & Format(rsFind("wmasuk"), "YYYY-MM-DD HH:MM:SS") & _
             "' and to_char(tgl_keluar,'YYYY-MM-DD HH24:MI:SS')='" & Format(rsFind("wkeluar"), "YYYY-MM-DD HH:MM:SS") & _
             "' and tara=" & rsFind("tara").Value & " and bruto=" & rsFind("bruto").Value & ""
    Set rsOra = cons.Execute(sqlStr)
    If rsOra.EOF Then
     'Data Tidak ditemukan
     
     Set rsFind2 = con.Execute("select nmpt from tbtruk " & _
                   "where nolambung='" & rsFind("nolambung") & "'")
     If Not rsFind2.EOF Then cArmada = rsFind2("nmpt")
     Set rsFind2 = Nothing
     
     'On Error Resume Next
     cons.Execute "INSERT INTO rf_truk_transaksi(rfid,nomor_1b,no_lambung,no_polisi," & _
        "nama_armada,tgl_masuk,tgl_muat,tgl_keluar,bruto," & _
        "tara,no_dermaga,nokey_gjt,nomer_gjt,sink) " & _
        "VALUES ('" & rsFind("rfid") & "','" & rsFind("nomor_1b") & "','" & rsFind("nolambung") & _
        "','" & rsFind("nopol") & "','" & cArmada & _
        "',to_date('" & Format(rsFind("wmasuk"), "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')" & _
        ",to_date('" & Format(rsFind("wmuat"), "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')" & _
        ",to_date('" & Format(rsFind("wkeluar"), "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')" & _
        "," & rsFind("bruto") & "," & rsFind("tara") & ",'" & rsFind("nodermaga") & _
        "','" & rsFind("nokey") & "','" & rsFind("nomer") & "','1')"
    Else
     'Data Ditemukan
     
     cons.Execute "update rf_truk_transaksi set sink='1'," & _
             "nokey_gjt='" & rsFind("nokey") & _
             "' where no_dermaga='" & rsFind("nodermaga").Value & _
             "' and no_lambung='" & rsFind("nolambung").Value & _
             "' and nomer_gjt='" & rsFind("nomer").Value & _
             "' and to_char(tgl_masuk,'YYYY-MM-DD HH24:MI:SS')='" & Format(rsFind("wmasuk"), "YYYY-MM-DD HH:MM:SS") & _
             "' and to_char(tgl_keluar,'YYYY-MM-DD HH24:MI:SS')='" & Format(rsFind("wkeluar"), "YYYY-MM-DD HH:MM:SS") & _
             "' and tara=" & rsFind("tara").Value & " and bruto=" & rsFind("bruto").Value & _
             " and sink<>'1'"
    End If
    Set rsOra = Nothing
    
    rsFind.MoveNext
    If pgb.Value < pgb.Max Then pgb.Value = pgb.Value + 1
   Wend
   Set rsFind = Nothing
   
   cons.Execute "delete from rf_truk_transaksi " & _
            "where (to_char(tgl_keluar,'YYYY-MM-DD')='" & Format(pvDate.Value, "yyyy-mm-dd") & _
            "' and (to_char(tgl_keluar,'HH24:MI:SS')>='" & Trim(cmbData(0).Text) & _
            "' and to_char(tgl_keluar,'HH24:MI:SS')<='" & Trim(cmbData(1).Text) & "'))" & _
            " and sink='0'"
   
   pgb.Visible = False
   stbData.Panels(2).Text = "Selesai"
  End Select
End Sub

Private Sub Form_Load()
 'Open Database --------------------
  'MySQl
  con.CursorLocation = adUseClient
  con.ConnectionTimeout = 0
  con.Open "DSN=dstimbang2"
  
  'Oracle
  On Error Resume Next
  cons.CursorLocation = adUseClient
  cons.ConnectionTimeout = 0
  cons.Open "dscuker", "cuker", "cuker"
 '----------------------------------
End Sub

Private Sub Form_Unload(Cancel As Integer)
 con.Close
 cons.Close
End Sub
