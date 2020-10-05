VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{C2000000-FFFF-1100-8100-000000000001}#8.0#0"; "PVCurr.ocx"
Object = "{CC696B60-4159-11D0-BDCB-0020A90B183A}#8.0#0"; "PVDateEdit.ocx"
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{6D0AB599-2BA8-4EB4-A7AB-130031613F61}#2.2#0"; "RMChart.ocx"
Begin VB.Form frmmontim 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitoring Penimbangan"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   17940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   17940
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   360
      Top             =   240
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filter Timbang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   13200
      TabIndex        =   22
      Top             =   2880
      Width           =   4575
      Begin PVCurrencyLib.PVCurrency pvinfo 
         Height          =   255
         Index           =   7
         Left            =   2040
         TabIndex        =   27
         Top             =   2360
         Width           =   2055
         _Version        =   524288
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   16384
         BackColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   0
         BackColor       =   8438015
         ForeColor       =   16384
         TextDisabled    =   49152
         Symbol          =   ""
         DecimalPlaces   =   "0"
         Value           =   0
      End
      Begin PVCurrencyLib.PVCurrency pvinfo 
         Height          =   255
         Index           =   8
         Left            =   2040
         TabIndex        =   29
         Top             =   2820
         Width           =   2055
         _Version        =   524288
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   12582912
         BackColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   0
         BackColor       =   8438015
         ForeColor       =   12582912
         TextDisabled    =   12582912
         Symbol          =   ""
         DecimalPlaces   =   "0"
         Value           =   0
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   1
         ItemData        =   "frmmontim.frx":0000
         Left            =   3120
         List            =   "frmmontim.frx":004F
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   780
         Width           =   1215
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   0
         ItemData        =   "frmmontim.frx":014D
         Left            =   1440
         List            =   "frmmontim.frx":019C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   780
         Width           =   1215
      End
      Begin ButtonPlusCtl.ButtonPlus ButtonPlus1 
         Height          =   615
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Proses"
      End
      Begin PVDATE2Lib.PVDate2 pvDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   0
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
         Value           =   41236.4538541667
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JUMLAH TRUK/RIT"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   28
         Top             =   2880
         Width           =   1485
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TONASE KELUAR"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   2400
         Width           =   1350
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   4560
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "sd"
         Height          =   195
         Index           =   4
         Left            =   2800
         TabIndex        =   25
         Top             =   840
         Width           =   165
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JAM TIMBANG"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TGL. TIMBANG"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Info Timbang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   13200
      TabIndex        =   5
      Top             =   120
      Width           =   4575
      Begin PVCurrencyLib.PVCurrency pvinfo 
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   12
         Top             =   1800
         Width           =   2055
         _Version        =   524288
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   8421504
         BackColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   0
         BackColor       =   8438015
         ForeColor       =   8421504
         TextDisabled    =   8421504
         Symbol          =   ""
         DecimalPlaces   =   "0"
         Value           =   0
      End
      Begin PVCurrencyLib.PVCurrency pvinfo 
         Height          =   255
         Index           =   5
         Left            =   2040
         TabIndex        =   11
         Top             =   1560
         Width           =   2055
         _Version        =   524288
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   12582912
         BackColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   0
         BackColor       =   8438015
         ForeColor       =   12582912
         TextDisabled    =   12582912
         Symbol          =   ""
         DecimalPlaces   =   "0"
         Value           =   0
      End
      Begin PVCurrencyLib.PVCurrency pvinfo 
         Height          =   255
         Index           =   4
         Left            =   2040
         TabIndex        =   10
         Top             =   1320
         Width           =   2055
         _Version        =   524288
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   16576
         BackColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   0
         BackColor       =   8438015
         ForeColor       =   16576
         TextDisabled    =   16576
         Symbol          =   ""
         DecimalPlaces   =   "0"
         Value           =   0
      End
      Begin PVCurrencyLib.PVCurrency pvinfo 
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
         _Version        =   524288
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   65535
         BackColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   0
         BackColor       =   8438015
         ForeColor       =   65535
         TextDisabled    =   65535
         Symbol          =   ""
         DecimalPlaces   =   "0"
         Value           =   0
      End
      Begin PVCurrencyLib.PVCurrency pvinfo 
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   8
         Top             =   600
         Width           =   2055
         _Version        =   524288
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   49152
         BackColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   0
         BackColor       =   8438015
         ForeColor       =   49152
         TextDisabled    =   49152
         Symbol          =   ""
         DecimalPlaces   =   "0"
         Value           =   0
      End
      Begin PVCurrencyLib.PVCurrency pvinfo 
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   7
         Top             =   840
         Width           =   2055
         _Version        =   524288
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   192
         BackColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   0
         BackColor       =   8438015
         ForeColor       =   192
         TextDisabled    =   192
         Symbol          =   ""
         DecimalPlaces   =   "0"
         Value           =   0
      End
      Begin PVCurrencyLib.PVCurrency pvinfo 
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   2055
         _Version        =   524288
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   8438015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   0
         BackColor       =   8438015
         TextDisabled    =   -2147483640
         Symbol          =   ""
         DecimalPlaces   =   "0"
         Value           =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B/L"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "YG KELUAR"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTIMASI DI KAPAL"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RATA-RATA (AVG)"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EST. SISA RIT"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EST. JUMLAH RIT"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRUK KELUAR"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1140
      End
   End
   Begin RMChart.RMChartX RMChartX1 
      Height          =   3735
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   120
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   6588
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UltraGrid.SSUltraGrid mGrid 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   4471
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   2
      LayoutFlags     =   72351744
      Override        =   "frmmontim.frx":029A
      CaptionAppearance=   "frmmontim.frx":0318
      Caption         =   "Jadwal Penimbangan Yang Berlangsung"
   End
   Begin MSComctlLib.StatusBar sbForm 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   6630
      Width           =   17940
      _ExtentX        =   31644
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   29025
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
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
End
Attribute VB_Name = "frmmontim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rsFind As New ADODB.Recordset
Dim rsFin2 As New ADODB.Recordset
Dim sqlStr As String

Option Explicit

Private Sub ButtonPlus1_Click()
 On Error Resume Next
 
 pvinfo(7).Value = 0
 pvinfo(8).Value = 0
 
 If cmbData(1).Text >= cmbData(0).Text Then
   'Proses Filter Timbang --------------------
    sqlStr = "select sum(netto) as neto from vwtrans " & _
             "where nomer='" & mGrid.ActiveRow.Cells(0).Value & _
             "' and (date(wkeluar)='" & _
                Format(pvDate.Value, "yyyy-mm-dd") & _
             "' and (time(wkeluar)>='" & Trim(cmbData(0).Text) & _
             "' and time(wkeluar)<='" & Trim(cmbData(1).Text) & "'))"
   Set rsFind = con.Execute(sqlStr)
   pvinfo(7).Value = rsFind("neto").Value
   Set rsFind = Nothing
 
 
   sqlStr = "select * from vwtrans " & _
            "where nomer='" & mGrid.ActiveRow.Cells(0).Value & _
            "' and (date(wkeluar)='" & _
               Format(pvDate.Value, "yyyy-mm-dd") & _
            "' and (time(wkeluar)>='" & Trim(cmbData(0).Text) & _
            "' and time(wkeluar)<='" & Trim(cmbData(1).Text) & "'))"
     
   Set rsFind = con.Execute(sqlStr)
   pvinfo(8).Value = rsFind.RecordCount
   Set rsFind = Nothing
  
 Else
  MsgBox "Pastikan Pengisian Jam Benar" & vbCrLf & _
         "Jam Awal harus lebih kecil nilainya dari Jam Akhir", _
         vbOKOnly + vbInformation, "JASATAMA"
 End If
End Sub

Private Sub Form_Load()
 'Open Database --------------------
  con.CursorLocation = adUseClient
  con.ConnectionTimeout = 0
  con.Open "DSN=dstimbang2"
 '----------------------------------
 
 sqlStr = "select nomer,nodermaga,tujuan,barge,tugboat,pemilik," & _
          "tglsandar,tglbongkar,bl from tbjadwal " & _
          "where fin='0' and st='1'"
 Set rs = con.Execute(sqlStr)
 InitGrid
 mGrid_Click
 
 DoTheChart
 
 cmbData(0).ListIndex = 0
 cmbData(1).ListIndex = 0
End Sub

Private Sub InitGrid()
 Dim i As Integer
 
 Set mGrid.DataSource = rs
 For i = 0 To rs.Fields.Count - 1
  mGrid.Bands(0).Columns(i).Activation = ssActivationActivateOnly
 Next
 mGrid.Appearance.BackColor = &HC0FFFF
 mGrid.Bands(0).Override.HeaderAppearance.BackColor = &H996633
 mGrid.Bands(0).Override.HeaderAppearance.ForeColor = vbWhite
 mGrid.Bands(0).Override.HeaderAppearance.Font.Bold = True
 'dGrid.Bands(0).Override.RowAppearance.BackColor = &HFFCC33
 
' mGrid.Bands(0).Columns(1).Hidden = True
' mGrid.Bands(0).Columns(2).Hidden = True
' mGrid.Bands(0).Columns(3).Hidden = True
 mGrid.Bands(0).Columns(8).Hidden = True
' mGrid.Bands(0).Columns(11).Hidden = True
' mGrid.Bands(0).Columns(12).Hidden = True
' mGrid.Bands(0).Columns(16).Hidden = True

' mGrid.Bands(0).Columns(0).Width = 1200
' mGrid.Bands(0).Columns(1).Width = 1950
' mGrid.Bands(0).Columns(2).Width = 500
 
' mGrid.Bands(0).Columns(0).Header.Caption = "ID KELAS"
' mGrid.Bands(0).Columns(1).Header.Caption = "NAMA KELAS"
' mGrid.Bands(0).Columns(2).Header.Caption = "TARIF"
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set rs = Nothing
 con.Close
End Sub

Private Sub mGrid_Click()
 Dim nTruk As Variant
 
 On Error Resume Next
 If Not rs.EOF Then
  mGrid.ActiveRow.Selected = True
  
  'Tampilkan B/L
  If mGrid.ActiveRow.Cells(8).Value = 0 Then pvinfo(0).Text = "-" _
    Else _
     pvinfo(0).Text = mGrid.ActiveRow.Cells(8).Value
     
  'Tampilkan Netto
  pvinfo(1).Text = "-"
  sqlStr = "select sum(netto) as netto from vwtrans " & _
           "where nomer='" & mGrid.ActiveRow.Cells(0).Value & "'"
  Set rsFind = con.Execute(sqlStr)
  If Not rsFind.EOF Then
   pvinfo(1).Value = rsFind(0).Value
  End If
  Set rsFind = Nothing
  
  'Tampilkan Estimasi di kapal
  If pvinfo(1).Value > 0 Then
   pvinfo(2).Value = pvinfo(0).Value - pvinfo(1).Value
  Else
   pvinfo(2).Text = "-"
  End If
   
  'Tampilkan Average Netto
  pvinfo(3).Text = "-"
  sqlStr = "select avg(netto) as neto from vwtrans " & _
           "where nomer='" & mGrid.ActiveRow.Cells(0).Value & "'"
  Set rsFind = con.Execute(sqlStr)
  If Not rsFind.EOF Then
   pvinfo(3).Value = Round(rsFind("neto").Value, 0)
  End If
  Set rsFind = Nothing
  
  'Tampilkan Estimasi Sisa Rit
  If pvinfo(3).Value > 0 Then
   pvinfo(4).Value = Round(pvinfo(2).Value / pvinfo(3).Value, 0)
  Else
   pvinfo(4).Text = "-"
  End If
  
  'Tampilkan Estimasi Jumlah Rit
  If pvinfo(3).Value > 0 Then
   pvinfo(5).Value = Round(pvinfo(0).Value / pvinfo(3).Value, 0)
  Else
   pvinfo(5).Text = "-"
  End If
  
  'Tampilkan Jumlah Truk Yang Keluar
  pvinfo(6).Text = "-"
  sqlStr = "select * from vwtrans " & _
           "where nomer='" & mGrid.ActiveRow.Cells(0).Value & "'"
  Set rsFind = con.Execute(sqlStr)
  If Not rsFind.EOF Then
   pvinfo(6).Text = rsFind.RecordCount
  End If
  Set rsFind = Nothing
  
 End If
End Sub

Private Sub DoTheChart()
 Dim nRetval As Long, pneto As Double
 Dim hChart As Integer, wChart As Integer
 Dim oldnomer As String
    
 On Error Resume Next
 
 hChart = 250
 wChart = 865
   
 With RMChartX1
  .Reset
  .RMCBackColor = AliceBlue
  .RMCStyle = RMC_CTRLSTYLEFLATSHADOW
  .RMCWidth = wChart
  .RMCHeight = hChart
  .RMCBgImage = ""
  .Font = "Tahoma"
  .RMCToolTipWidth = 0
  .RMCHelpingGridSize = 0
  .RMCHelpingGridColor = Default
  .RMCBitmapColor = Default
  '************** Add Region 2 *****************************
   .AddRegion
   With .Region(1)
    .Left = 20
    .Top = 5
    .Width = -2
    .Height = hChart
    .Footer = ""
    '************** Add grid to region 2 *****************************
     .AddGrid
     With .Grid
      .BackColor = Beige
      .AsGradient = False
      .BicolorMode = RMC_BICOLOR_NONE
      .Left = 0
      .Top = 0
      .Width = 0
      .Height = 0
     End With 'Grid
     '************** Add data axis to region 2 *****************************
      .AddDataAxis
      With .DataAxis(1)
       .Alignment = RMC_DATAAXISBOTTOM
       .MinValue = 0
       .MaxValue = 130
       .TickCount = 14
       .FontSize = 8
       .TextColor = Black
       .LineColor = Black
       .LineStyle = RMC_LINESTYLESOLID
       .DecimalDigits = 0
       .AxisUnit = "%"
       .AxisText = ""
      End With 'DataAxis(1)
      '************** Add label axis to region 2 *****************************
       .AddLabelAxis
       With .LabelAxis
        .AxisCount = 1
        .TickCount = 5
        .Alignment = RMC_LABELAXISLEFT
        .FontSize = 7
        .TextColor = Black
        .TextAlignment = RMC_TEXTCENTER
        .LineColor = Black
        .LineStyle = RMC_LINESTYLESOLID
        .AxisText = "Grafik Progress Penimbangan" & vbCrLf & "Yang Berlangsung"
        .LabelString = ""
                
        sqlStr = "select nomer,barge,nodermaga from tbjadwal " & _
                 "where fin='0' and st='1'"
        Set rsFind = con.Execute(sqlStr)
        While Not rsFind.EOF
         .LabelString = .LabelString & rsFind("nomer") & _
            vbCrLf & rsFind("barge") & " (" & _
            rsFind("nodermaga") & ")"
         rsFind.MoveNext
         If Not rsFind.EOF Then .LabelString = .LabelString & "*"
        Wend
        Set rsFind = Nothing
       End With 'LabelAxis
       '************** Add Series 1 to region 2 *******************************
        .AddBarSeries
        With .BarSeries(1)
         .SeriesType = RMC_BARSINGLE 'RMC_BARSTACKED
         .SeriesStyle = RMC_COLUMN_FLAT
         .Lucent = True
         .Color = OrangeRed
         .Horizontal = True
         .WhichDataAxis = 1
         .ValueLabelOn = RMC_VLABEL_NONE
         .PointsPerColumn = 1
         .HatchMode = RMC_HATCHBRUSH_ON
         .DataString = ""
         
         '****** Set data values ******
         sqlStr = "select nomer,bl from tbjadwal " & _
                  "where fin='0' and st='1'"
         Set rsFind = con.Execute(sqlStr)
         While Not rsFind.EOF
          pneto = 0
          sqlStr = "select sum(netto) as neto from vwtrans " & _
                   "where nomer='" & Trim(rsFind("nomer").Value) & "'"
          Set rsFin2 = con.Execute(sqlStr)
          If Not IsNull(rsFin2("neto")) And _
             rsFind("bl") > 0 Then
           pneto = rsFin2("neto") / rsFind("bl")
          End If
          .DataString = .DataString & Trim(Str(Round(pneto * 100, 2)))
          Set rsFin2 = Nothing
          rsFind.MoveNext
          If Not rsFind.EOF Then _
             .DataString = .DataString & "*"
         Wend
         Set rsFind = Nothing
         
        End With 'BarSeries(1)
            '************** Add Series 2 to region 2 *******************************
'            .AddBarSeries
'            With .BarSeries(2)
'                .SeriesType = RMC_BARSTACKED
'                .SeriesStyle = RMC_COLUMN_FLAT
'                .Lucent = False
'                .Color = Default
'                .Horizontal = True
'                .WhichDataAxis = 1
'                .ValueLabelOn = RMC_VLABEL_NONE
'                .PointsPerColumn = 1
'                .HatchMode = RMC_HATCHBRUSH_OFF
                '****** Set data values ******
'                .DataString = "25*30*10"
'            End With 'BarSeries(2)
            '************** Add Series 3 to region 2 *******************************
'            .AddBarSeries
'            With .BarSeries(3)
'                .SeriesType = RMC_BARSTACKED
'                .SeriesStyle = RMC_COLUMN_FLAT
'                .Lucent = False
'                .Color = Default
'                .Horizontal = True
'                .WhichDataAxis = 1
'                .ValueLabelOn = RMC_VLABEL_NONE
'                .PointsPerColumn = 1
'                .HatchMode = RMC_HATCHBRUSH_OFF
                '****** Set data values ******
'                .DataString = "10*20*40"
'            End With 'BarSeries(3)
            '************** Add Series 4 to region 2 *******************************
'            .AddBarSeries
'            With .BarSeries(4)
'                .SeriesType = RMC_BARSTACKED
'                .SeriesStyle = RMC_COLUMN_FLAT
'                .Lucent = False
'                .Color = Default
'                .Horizontal = True
'                .WhichDataAxis = 1
'                .ValueLabelOn = RMC_VLABEL_NONE
'                .PointsPerColumn = 1
'                .HatchMode = RMC_HATCHBRUSH_OFF
                '****** Set data values ******
'                .DataString = "40*30*20"
'            End With 'BarSeries(4)
        End With 'Region(2)
        nRetval = .Draw(True)
    End With 'RMChartX1
End Sub

Private Sub Timer1_Timer()
 DoTheChart
 mGrid_Click
End Sub
