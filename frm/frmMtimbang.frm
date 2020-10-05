VERSION 5.00
Object = "{0B63DF3F-CC00-4D55-A1C9-CAFE70BB1B49}#1.0#0"; "xpctrl.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{CC696B60-4159-11D0-BDCB-0020A90B183A}#8.0#0"; "PVDateEdit9.ocx"
Begin VB.Form frmMtimbang 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maintenance Transaksi Timbang"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11565
   FillColor       =   &H00C00000&
   Icon            =   "frmMtimbang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   11565
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   1575
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin XPCtrl.XPButton cmdProses 
         Height          =   480
         Left            =   5040
         TabIndex        =   14
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   847
         BTYPE           =   3
         TX              =   "&Proses"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmMtimbang.frx":57E2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox chkData 
         BackColor       =   &H000080FF&
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   13
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Height          =   320
         Index           =   2
         Left            =   4365
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.CheckBox chkData 
         BackColor       =   &H000080FF&
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkData 
         BackColor       =   &H000080FF&
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Height          =   320
         Index           =   1
         Left            =   1720
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Height          =   320
         Index           =   0
         Left            =   3045
         MaxLength       =   4
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         ItemData        =   "frmMtimbang.frx":57FE
         Left            =   1680
         List            =   "frmMtimbang.frx":5826
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin PVDATE2Lib.PVDate2 pvDate 
         Height          =   300
         Left            =   1720
         TabIndex        =   7
         Top             =   960
         Width           =   1335
         _Version        =   524288
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Value           =   41522.7215509259
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Lambung"
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
         Left            =   2880
         TabIndex        =   12
         Top             =   645
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl.Timbang"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor"
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
         Left            =   240
         TabIndex        =   6
         Top             =   645
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan && Tahun"
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
         Left            =   240
         TabIndex        =   4
         Top             =   285
         Width           =   1230
      End
   End
   Begin UltraGrid.SSUltraGrid mGrid 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8070
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   2
      LayoutFlags     =   72351744
      Override        =   "frmMtimbang.frx":588D
      CaptionAppearance=   "frmMtimbang.frx":590B
      Caption         =   "Transaksi Timbang"
   End
   Begin XPCtrl.XPButton cmdCetak 
      Height          =   480
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   6360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Cetak Struk &In"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmMtimbang.frx":5957
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XPCtrl.XPButton cmdCetak 
      Height          =   480
      Index           =   1
      Left            =   2400
      TabIndex        =   16
      Top             =   6360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "Cetak Struk &Out"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmMtimbang.frx":5973
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XPCtrl.XPButton cmdDel 
      Height          =   480
      Left            =   9960
      TabIndex        =   17
      Top             =   6360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   847
      BTYPE           =   3
      TX              =   "&Hapus"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmMtimbang.frx":598F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   120
      Picture         =   "frmMtimbang.frx":59AB
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4620
   End
End
Attribute VB_Name = "frmMtimbang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rsDtl As New ADODB.Recordset
Dim sqlStr As String, lastclick As String

Dim A As Boolean, b As Boolean, c As Boolean
Dim nbulan, cBulan As String

Dim cNomer As String
'Dim cPemilik As String, cnDmg As String
'Dim nBruto, nTara, nNeto, cBarang As String
'Dim cnoLambung As String, cnoPol As String
'Dim wkMasuk, wkKeluar, cNmr As String

Option Explicit

'Private Function Cetak_Struk(cFile As String) As String
' Dim cData As String, fBaris As Long
' Dim i As Integer, nAw As Integer, nAk As Integer
' Dim cTemp As String, cField As String
 
' cData = FileRead(cFile, False, fBaris)(1): cData = vbNullString
' For i = 1 To fBaris
'  cTemp = FileRead(cFile, False)(i)
  
  'Membaca Apakah Ada Sebuah Field ? ------------------
'   nAw = InStr(cTemp, "<<"): nAk = InStr(cTemp, ">>")
'   If nAw <> 0 Then
'    cField = Mid(cTemp, nAw + 2, nAk - (nAw + 2))
'    If cField = "bruto" Then
'     cTemp = Mid(cTemp, 1, nAw - 1) & nBruto & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
'    ElseIf cField = "tara" Then
'     cTemp = Mid(cTemp, 1, nAw - 1) & nTara & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
'    ElseIf cField = "netto" Then
'     cTemp = Mid(cTemp, 1, nAw - 1) & nNeto & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
'    ElseIf cField = "nolambung" Then
'     cTemp = Mid(cTemp, 1, nAw - 1) & cnoLambung
'    ElseIf cField = "nopol" Then
'     cTemp = Mid(cTemp, 1, nAw - 1) & cnoPol
'    ElseIf cField = "wmasuk" Then
'     cTemp = Mid(cTemp, 1, nAw - 1) & Format(wkMasuk, "dd-mm-yyyy / HH:MM:SS")
'    ElseIf cField = "wkeluar" Then
'     cTemp = Mid(cTemp, 1, nAw - 1) & Format(wkKeluar, "dd-mm-yyyy / HH:MM:SS")
'    ElseIf cField = "barang" Then
'     cTemp = Mid(cTemp, 1, nAw - 1) & cBarang
'    ElseIf cField = "pemilik" Then
'     cTemp = Mid(cTemp, 1, nAw - 1) & cPemilik
'    ElseIf cField = "nomer" Then
'     cTemp = Mid(cTemp, 1, nAw - 1) & cNmr
'    ElseIf cField = "nodermaga" Then
'     cTemp = Mid(cTemp, 1, nAw - 1) & cnDmg
'    ElseIf cField = "nmoperator" Then
'     cTemp = Mid(cTemp, 1, nAw - 1) & UserID
'    End If
'   End If
  '----------------------------------------------------
  
'  cData = cData & cTemp
'  If i <> fBaris Then cData = cData & vbCrLf
' Next
' Cetak_Struk = cData
'End Function

Private Sub chkdata_Click(Index As Integer)
 A = False: b = False: c = False
 
 If chkData(0).Value = 1 Then A = True
 If chkData(1).Value = 1 Then b = True
 If chkData(2).Value = 1 Then c = True
End Sub

Private Sub cmbData_Click()
 nbulan = cmbData.ListIndex + 1
 If Len(Trim(Str(nbulan))) = 1 Then
  cBulan = "0" & Trim(Str(nbulan))
 Else
  cBulan = Trim(Str(nbulan))
 End If
 cNomer = Trim(txtData(0).Text) & "/" & Trim(cBulan) & "/" & Trim(txtData(1).Text)
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
 
 mGrid.Bands(0).Columns(1).Hidden = True
 mGrid.Bands(0).Columns(2).Hidden = True
 mGrid.Bands(0).Columns(3).Hidden = True
 mGrid.Bands(0).Columns(7).Hidden = True
 mGrid.Bands(0).Columns(11).Hidden = True
 mGrid.Bands(0).Columns(12).Hidden = True
 mGrid.Bands(0).Columns(16).Hidden = True

 mGrid.Bands(0).Columns(0).Width = 1200
' mGrid.Bands(0).Columns(1).Width = 1950
' mGrid.Bands(0).Columns(2).Width = 500
 
' mGrid.Bands(0).Columns(0).Header.Caption = "ID KELAS"
' mGrid.Bands(0).Columns(1).Header.Caption = "NAMA KELAS"
' mGrid.Bands(0).Columns(2).Header.Caption = "TARIF"
End Sub

Private Sub cmdCetak_Click(Index As Integer)
 Dim i As Byte
 
 If gNomer <> vbNullString Then
  Select Case Index
  Case 0     'Cetak Struk In
   'Cetak "Struk Masuk", Cetak_Struk(App.Path & "\rpt\strukin.txt")
   MsgBox Cetak_Struk2(App.Path & "\rpt\strukin.txt")
   Cetak_Struk App.Path & "\rpt\strukin.txt"
  Case 1     'Cetak Struk Out
   MsgBox Cetak_Struk2(App.Path & "\rpt\strukout.txt")
   For i = 0 To 1
    'Cetak "Struk Keluar", Cetak_Struk(App.Path & "\rpt\strukout.txt")
    Cetak_Struk App.Path & "\rpt\strukout.txt"
   Next
   'MsgBox Cetak_Struk(App.Path & "\rpt\strukout.txt")
  End Select
 End If
End Sub

Private Sub cmdDel_Click()
 Dim cKey As String
 
 'cKey = Year(wkMasuk) & Month(wkMasuk) & Day(wkMasuk) & Hour(wkMasuk) & Minute(wkMasuk) & Second(wkMasuk)
 cKey = Year(gMasuk) & Month(gMasuk) & Day(gMasuk) & _
        Hour(gMasuk) & Minute(gMasuk) & Second(gMasuk)
 
 sqlStr = "delete from tbtrans " & _
          "where nokey='" & Trim(cKey) & "'"
 
 If MsgBox("Yakin Data Dihapus ?", vbYesNo + vbQuestion + vbDefaultButton2, "JASATAMA") = vbYes Then
  con.Execute sqlStr
  
  cmdProses_Click
 End If
End Sub

Private Sub cmdProses_Click()
 sqlStr = "select * from vwtrans " & _
          "where left(nomer,4)='" & txtData(0).Text & _
          "' and mid(nomer,6,2)='" & cBulan & "' "
 
 If A Then  'Jika Nomer Dicentang
  sqlStr = sqlStr & "and nomer='" & Trim(cNomer) & "' "
 End If
 
 If b Then  'Jika Tanggal Dicentang
  'sqlStr = sqlStr & "and (year(wkeluar)=" & Year(pvDate.Value) & _
           " and month(wkeluar)=" & Month(pvDate.Value) & _
           " and day(wkeluar)=" & Day(pvDate.Value) & ") "
  sqlStr = sqlStr & "and date(wkeluar)='" & _
           Format(pvDate.Value, "yyyy-mm-dd") & "'"
 End If
 
 If c Then  'Jika No.Lambung Dicentang
  sqlStr = sqlStr & "and nolambung='" & Trim(txtData(2).Text) & "' "
 End If
 
 Set rs = con.Execute(sqlStr)
 InitGrid
End Sub

Private Sub Form_Load()
 'Open Database --------------------
  con.CursorLocation = adUseClient
  con.ConnectionTimeout = 0
  con.Open "DSN=dstimbang2"
 '----------------------------------
 
 cmbData.ListIndex = Month(Date) - 1
 txtData(0).Text = Year(Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set rs = Nothing
 con.Close
End Sub

Private Sub mGrid_Click()
 If mGrid.HasRows Then
'  nBruto = mGrid.ActiveRow.Cells(13).Value
'  nTara = mGrid.ActiveRow.Cells(14).Value
'  nNeto = mGrid.ActiveRow.Cells(15).Value
'  cnoLambung = mGrid.ActiveRow.Cells(9).Value
'  cnoPol = mGrid.ActiveRow.Cells(10).Value
'  wkMasuk = mGrid.ActiveRow.Cells(11).Value
'  wkKeluar = mGrid.ActiveRow.Cells(12).Value
'  cBarang = mGrid.ActiveRow.Cells(1).Value
'  cPemilik = mGrid.ActiveRow.Cells(3).Value
'  cnDmg = mGrid.ActiveRow.Cells(7).Value
'  cNmr = mGrid.ActiveRow.Cells(0).Value

  gBruto = mGrid.ActiveRow.Cells(13).Value
  gTara = mGrid.ActiveRow.Cells(14).Value
  gNetto = mGrid.ActiveRow.Cells(15).Value
  gLambung = mGrid.ActiveRow.Cells(9).Value
  gNopol = mGrid.ActiveRow.Cells(10).Value
  gMasuk = mGrid.ActiveRow.Cells(11).Value
  gKeluar = mGrid.ActiveRow.Cells(12).Value
  gBarang = mGrid.ActiveRow.Cells(1).Value
  gPemilik = mGrid.ActiveRow.Cells(3).Value
  gDermaga = mGrid.ActiveRow.Cells(7).Value
  gNomer = mGrid.ActiveRow.Cells(0).Value
 Else
  'cNmr = vbNullString
  gNomer = vbNullString
 End If
End Sub

Private Sub mGrid_DblClick()
 If mGrid.HasRows Then
  frmUpTrans.Show 1
 
  Set rs = con.Execute(sqlStr)
  InitGrid
 End If
End Sub

Private Sub txtData_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  Select Case Index
  Case 1
   cmbData_Click
  End Select
End Sub
