VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0B63DF3F-CC00-4D55-A1C9-CAFE70BB1B49}#1.0#0"; "xpctrl.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTruk 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DATA TRUK / ANGKUTAN"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmTruk.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   8191.492
   ScaleMode       =   0  'User
   ScaleWidth      =   10152.02
   ShowInTaskbar   =   0   'False
   Begin sit30.navCtrl navCtrl1 
      Height          =   735
      Left            =   3960
      TabIndex        =   12
      Top             =   6480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1296
   End
   Begin VB.Frame frCari 
      BackColor       =   &H000080FF&
      Caption         =   "Cari Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3960
      TabIndex        =   9
      Top             =   5040
      Width           =   6015
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Height          =   540
         Index           =   6
         Left            =   1800
         MaxLength       =   16
         TabIndex        =   11
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Lambung/Nama Truk/No.Polisi :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   3060
      End
   End
   Begin MSAdodcLib.Adodc dAdo 
      Height          =   330
      Left            =   1320
      Top             =   3960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc mAdo 
      Height          =   330
      Left            =   960
      Top             =   1320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Height          =   1575
      Left            =   3960
      TabIndex        =   6
      Top             =   1080
      Width           =   6015
      Begin XPCtrl.XPButton dtlBtn 
         Height          =   495
         Index           =   0
         Left            =   4680
         TabIndex        =   16
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Tambah"
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
         MICON           =   "frmTruk.frx":57E2
         PICN            =   "frmTruk.frx":57FE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtData 
         Height          =   300
         Index           =   5
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtData 
         Height          =   300
         Index           =   4
         Left            =   2040
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin XPCtrl.XPButton dtlBtn 
         Height          =   495
         Index           =   1
         Left            =   4680
         TabIndex        =   17
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Hapus"
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
         MICON           =   "frmTruk.frx":6210
         PICN            =   "frmTruk.frx":622C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Polisi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   885
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Lambung"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   855
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   420
         Index           =   1
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   360
         MaxLength       =   5
         TabIndex        =   0
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Penyedia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1515
      End
   End
   Begin MSComctlLib.StatusBar sbForm 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   13
      Top             =   7320
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   15134
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
   Begin UltraGrid.SSUltraGrid mGrid 
      Height          =   2775
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4895
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   2
      LayoutFlags     =   72351748
      RowConnectorColor=   65535
      Override        =   "frmTruk.frx":6C3E
      CaptionAppearance=   "frmTruk.frx":6CBC
      Caption         =   "Penyedia Truk / Angkutan"
   End
   Begin UltraGrid.SSUltraGrid dGrid 
      Height          =   3855
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6800
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   2
      LayoutFlags     =   72351744
      Override        =   "frmTruk.frx":6D08
      CaptionAppearance=   "frmTruk.frx":6D86
      Caption         =   "Data Truk / Angkutan"
   End
   Begin VB.Image Image1 
      Height          =   2145
      Left            =   3960
      Picture         =   "frmTruk.frx":6DD2
      Top             =   2760
      Width           =   6390
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   7455
      Left            =   0
      Top             =   -120
      Width           =   10095
   End
End
Attribute VB_Name = "frmTruk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'a>. Penghapusan Detail Secara Blok Belum
'    Terakomodir -- Masih Dipikirkan

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset, rsFind As New ADODB.Recordset
Dim sqlStr As String
Dim iDxForm As Integer, recPointer As Integer

Dim bEditing As Boolean, bEditKey As Boolean
Dim bSave As Boolean, bHapusDtl As Boolean
Dim tmpTiket As String, bOut As Boolean

Option Explicit

Private Sub dGrid_Click()
 If Not rs.EOF Then
  dGrid.ActiveRow.Cells(0).Activation = ssActivationActivateOnly
  If Not bEditing Then dGrid.ActiveRow.Selected = True
 End If
End Sub

Private Sub dtlBtn_Click(Index As Integer)
 Dim i As Integer
 
 Select Case Index
 Case 0             'Insert Truk
  If Trim(txtData(4).Text) <> vbNullString Then
   sqlStr = "select * from tbtruk " & _
            "where nolambung='" & Trim(txtData(4).Text) & "'"
   Set rsFind = con.Execute(sqlStr)
   If rsFind.EOF Then
    dGrid.Bands(0).AddNew
    dGrid.ActiveRow.Cells(0).Value = Trim(txtData(4).Text)
    dGrid.ActiveRow.Cells(1).Value = txtData(5).Text
    dGrid.ActiveRow.Cells(2).Value = Trim(txtData(0).Text)
    dGrid.ActiveRow.Cells(3).Value = txtData(1).Text
   
    txtData(4).Text = vbNullString
    txtData(5).Text = vbNullString
   Else
    MsgBox "No.Lambung " & Trim(txtData(4).Text) & " Sudah Ada." & vbCrLf & _
           "Gunakan Kode Lambung Yang Berbeda.", vbOKOnly + vbInformation, "JASATAMA"
   End If
   Set rsFind = Nothing
   
   txtData(4).SetFocus
  End If
 
 Case 1             'Hapus Truk
  If dGrid.HasRows Then
   dGrid.Bands(0).Override.AllowDelete = ssAllowDeleteYes
    
   bHapusDtl = True
   tmpTiket = dGrid.ActiveRow.Cells(1).Value
   sqlStr = "select * from tbTruk " & _
           "where nolambung='" & tmpTiket & "'"
   Set rsFind = con.Execute(sqlStr)
   If Not rsFind.EOF Then bHapusDtl = False
   Set rsFind = Nothing
    

   If bHapusDtl Then
    dGrid.DeleteSelectedRows
    Set dGrid.ActiveRow = dGrid.GetRow(ssChildRowFirst)
   End If

   dGrid.Bands(0).Override.AllowDelete = ssAllowDeleteNo
  End If
 End Select
End Sub

Private Sub Form_Activate()
 Me.Height = 8445
 Me.Width = 10155
 
 iDxFrm = 4
 If bOut Then
  bOut = False
  Unload Me
 End If
End Sub

Private Sub Form_Load()
 iDxForm = 4
 bEditing = True
 bEditKey = True
 Set GForm(iDxForm) = frmTruk
 
 'Open Database --------------------
  con.CursorLocation = adUseClient
  con.ConnectionTimeout = 0
  con.Open "DSN=dstimbang2"
 '----------------------------------
 
 sqlStr = "Select * from vwTrukPt order by kode"
                      
 dAdo.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;" & _
          "Data Source=dstimbang2" '& cODBCName
 DetilGrid "select * from tbTruk order by kdpt,nolambung"
 
 Set rs = con.Execute(sqlStr)
 If Not rs.EOF Then
  bOut = False
  InitGrid
  bEditing = False
  bEditKey = False
  recPointer = rs.RecordCount - 1
  rs.MoveLast
  PScatter
 Else
  bOut = True
  MsgBox "Data Penyedia Angkutan Tidak Ada." & vbCrLf & _
         "Isikan Terlebih Dahulu.", vbOKOnly + vbInformation, "JASATAMA"
 End If
 If Not bOut Then pActive
End Sub

'Sub Navigasi Tombol
Public Sub PControl(Index As Integer)
 Select Case Index
 Case 0 'First
  rs.MoveFirst
  recPointer = 0
  PScatter
 Case 1 'Prev
  If recPointer > 0 Then
   rs.MovePrevious
   recPointer = recPointer - 1
   PScatter
  Else
   MsgBox "Awal Record", vbInformation + vbOKOnly, "JASATAMA"
  End If
 Case 2 'Next
  If recPointer <> (rs.RecordCount - 1) Then
   rs.MoveNext
   recPointer = recPointer + 1
   PScatter
  Else
   MsgBox "Akhir Record", vbInformation + vbOKOnly, "JASATAMA"
  End If
 Case 3 'Last
  rs.MoveLast
  recPointer = (rs.RecordCount - 1)
  PScatter
 Case 4 'Add
  'bEditing = False
  'bEditKey = False
  'PForm
  Index = 11
  MsgBox "Gunakan Edit Untuk Memasukkan Data Truk", vbOKOnly + vbInformation, "JASATAMA"
 Case 5 'Edit
  bEditing = True
  PForm
 Case 6 'Del
  PDelete
  PScatter
  PForm
 Case 7 'Find
  PFind
 Case 8 'Print
  PPrint
 Case 9 'Close
  Unload Me
 Case 10 'Save
  PSave
  If bSave Then
   bEditing = False
   bEditKey = False
   PForm
  End If
 Case 11 'Batal
  If bEditKey Then PDelete
  If rs.RecordCount > 0 Then
   If bEditKey Then rs.MoveLast
   bEditing = False
   bEditKey = False
   PScatter
   PForm
  Else
   Unload Me
  End If
 End Select
End Sub

Private Sub PScatter()
 Dim i As Integer
 
 On Error Resume Next 'Jika ada Field yang bernilai NULL
 If rs.RecordCount > 0 Then
  txtData(0).Text = rs.Fields(0).Value
  txtData(1).Text = rs.Fields(1).Value
  
  'Detil Tiket
  DetilGrid "select * from tbTruk " & _
            "where kdpt='" & Trim(mGrid.ActiveRow.Cells(0).Value) & _
            "' order by kdpt,nolambung"
            
  mGrid.Refresh ssRefetchAndFireInitializeRow
  mGrid.ActiveRow.Selected = True
 Else
  bEditing = True
  bEditKey = True
 End If
End Sub

Private Sub PBlank()
 Dim i As Integer
  
 'Detil Tiket
 DetilGrid "select * from tbTruk " & _
           "where kdpt='" & Trim(txtData(0).Text) & _
           "' order by kdpt,nolambung"
End Sub

Private Sub PForm()
 Dim i As Integer
 
 pActive
 If bEditKey Then
  PBlank
 Else
  If bEditing Then
   txtData(4).SetFocus
  End If
 End If
End Sub

Private Sub pActive()
 If bEditing Then
  pActivForm
 Else
  pDeactivForm
 End If
End Sub

Private Sub pActivForm()
 Dim i As Integer
 
 mGrid.Enabled = False
 frCari.Enabled = False
 For i = 4 To 5
  txtData(i).Enabled = True
 Next
 For i = 0 To 1
  dtlBtn(i).Enabled = True
 Next
 navCtrl1.EditPos
End Sub

Private Sub pDeactivForm()
 Dim i As Integer
 
 mGrid.Enabled = True
 frCari.Enabled = True
 For i = 4 To 5
  txtData(i).Enabled = False
 Next
 For i = 0 To 1
  dtlBtn(i).Enabled = False
 Next
 navCtrl1.ViewPos
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set rs = Nothing
 con.Close
End Sub

Private Sub mGrid_Click()
 If Not rs.EOF Then
  mGrid.ActiveRow.Selected = True
  PScatter
 End If
End Sub

Private Sub mGrid_KeyUp(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
 mGrid_Click
End Sub

Private Sub txtData_GotFocus(Index As Integer)
 txtData(Index).BackColor = &HC0FFC0
 SendKeys "{home}+{end}"
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
  If Index = 6 Then
   PFind
  Else
   SendKeys vbTab
  End If
 Else
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
 End If
End Sub

Private Sub PSave()
 Dim sMsg As Variant
 
 bSave = False
 sMsg = MsgBox("Apakah Data Sudah Benar ?", vbYesNo + vbQuestion + vbDefaultButton2, "JASATAMA")
 If sMsg = vbYes Then
   bSave = True
          
   dGrid.Update
   rs.Requery
   rs.MoveLast
   InitGrid    'Refresh Grid Master
 End If
End Sub

Private Sub PDelete()
 'Sengaja Untuk Dikosongkan
End Sub

Private Sub InitGrid()
 Dim i As Integer
 
 Set mGrid.DataSource = rs
 For i = 0 To rs.Fields.Count - 1
  mGrid.Bands(0).Columns(i).Activation = ssActivationActivateOnly
 Next
 mGrid.Bands(0).Columns(0).Width = 800
 mGrid.Bands(0).Columns(1).Width = 2450
End Sub

Private Sub DetilGrid(MySql As String)
 dAdo.RecordSource = MySql
 dAdo.Refresh
 Set dGrid.DataSource = dAdo
 
 dGrid.Appearance.BackColor = &HC0FFFF
 dGrid.Bands(0).Columns(2).Hidden = True
 dGrid.Bands(0).Columns(3).Hidden = True
 dGrid.Bands(0).Override.HeaderAppearance.BackColor = &H996633
 dGrid.Bands(0).Override.HeaderAppearance.ForeColor = vbWhite
 dGrid.Bands(0).Override.HeaderAppearance.Font.Bold = True
 'dGrid.Bands(0).Override.RowAppearance.BackColor = &HFFCC33
 
 dGrid.Bands(0).Columns(0).Activation = ssActivationActivateOnly
 dGrid.Bands(0).Columns(0).Width = 1250
 dGrid.Bands(0).Columns(0).Header.Caption = "No.Lambung"
 dGrid.Bands(0).Columns(1).Width = 1950
 dGrid.Bands(0).Columns(1).Header.Caption = "No.Polisi"
End Sub

Private Sub PPrint()
' noRPT = 1
' frmPreview.Show
End Sub

Private Sub PFind()
 Dim sBukti As String, sLambung As String
 
 If frCari.Enabled Then
  sBukti = vbNullString
  sqlStr = "select kdpt,nolambung from tbTruk " & _
           "where nolambung='" & txtData(6).Text & _
           "' or kdpt='" & txtData(6).Text & _
           "' or nmpt LIKE '%" & txtData(6).Text & _
           "%' or nopolisi LIKE '%" & txtData(6).Text & "%'"
  Set rsFind = con.Execute(sqlStr)
  If Not rsFind.EOF Then
   sBukti = rsFind.Fields(0).Value
   sLambung = rsFind.Fields(1).Value
  End If
  Set rsFind = Nothing
  If sBukti <> vbNullString Then
   rs.Find "kode LIKE '%" & sBukti & "%'", , adSearchForward, 1
   PScatter
   
   If Not dAdo.Recordset.EOF Then
    dAdo.Recordset.Find "noLambung LIKE '%" & sLambung & "%'", , adSearchForward, 1
    dGrid.ActiveRow.Selected = True
   End If
  End If
 End If
End Sub

Private Sub txtData_LostFocus(Index As Integer)
 txtData(Index).BackColor = &HFFFFFF
End Sub
