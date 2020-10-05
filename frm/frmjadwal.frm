VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0B63DF3F-CC00-4D55-A1C9-CAFE70BB1B49}#1.0#0"; "xpctrl.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{C2000000-FFFF-1100-8100-000000000001}#8.0#0"; "PVCurr.ocx"
Object = "{CC696B60-4159-11D0-BDCB-0020A90B183A}#8.0#0"; "PVDateEdit9.ocx"
Begin VB.Form frmjadwal 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jadwal Timbang"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   FillColor       =   &H00C00000&
   Icon            =   "frmjadwal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8430
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   29
      Top             =   2520
      Width           =   3135
      Begin VB.TextBox txtdata 
         Height          =   285
         Index           =   11
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   230
         Width           =   1455
      End
      Begin VB.Label lbltboat 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   225
         Left            =   1320
         TabIndex        =   34
         Top             =   1080
         Width           =   1620
      End
      Begin VB.Label lblbge 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   225
         Left            =   1320
         TabIndex        =   33
         Top             =   720
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tug Boat ==>"
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
         Index           =   11
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barge      ==>"
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
         Index           =   10
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Permohonan"
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
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1365
      End
   End
   Begin sit30.navCtrl navCtrl1 
      Height          =   735
      Left            =   1200
      TabIndex        =   10
      Top             =   6120
      Width           =   6015
      _extentx        =   10610
      _extenty        =   1296
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   3495
      Left            =   3360
      TabIndex        =   12
      Top             =   2520
      Width           =   4935
      Begin PVCurrencyLib.PVCurrency pvbl 
         Height          =   300
         Left            =   3120
         TabIndex        =   8
         Top             =   3000
         Width           =   1695
         _Version        =   524288
         _ExtentX        =   2990
         _ExtentY        =   529
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Alignment       =   2
         Symbol          =   ""
         DecimalPlaces   =   "0"
         Value           =   0
      End
      Begin VB.TextBox txtdata 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "0"
         Top             =   680
         Width           =   375
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2400
         TabIndex        =   27
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   2400
         TabIndex        =   25
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txtdata 
         Height          =   285
         Index           =   8
         Left            =   1440
         TabIndex        =   7
         Top             =   2520
         Width           =   855
      End
      Begin XPCtrl.XPButton cmdBtn 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Mulai"
         ENAB            =   0   'False
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
         MICON           =   "frmjadwal.frx":57E2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtdata 
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   6
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   2400
         TabIndex        =   22
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtdata 
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   5
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   2400
         TabIndex        =   20
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtdata 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2400
         TabIndex        =   11
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtdata 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin PVDATE2Lib.PVDate2 pvDate 
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   1695
         _Version        =   524288
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
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
         Border          =   1
         DisplayFormat   =   6
         CalendarFormat  =   4
         DateFormat      =   13
         TimeStore       =   -1  'True
         HighlightColor  =   12632256
         BackColor       =   16777215
         ForeColor       =   0
         Value           =   41517.4031712963
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No.Dermaga"
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
         Index           =   8
         Left            =   3240
         TabIndex        =   28
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pemilik Brg."
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
         Index           =   7
         Left            =   240
         TabIndex        =   26
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B / L (Kg)"
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
         Index           =   6
         Left            =   2160
         TabIndex        =   24
         Top             =   3045
         Width           =   750
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   7920
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tug Boat"
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
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Barge"
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
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Sandar"
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
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   7920
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tujuan"
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
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomer       ==>"
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
         Index           =   0
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Barang"
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
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblnomer 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3120
         TabIndex        =   13
         Top             =   240
         Width           =   1620
      End
   End
   Begin UltraGrid.SSUltraGrid mGrid 
      Height          =   2295
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4048
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   2
      LayoutFlags     =   67108884
      BorderStyle     =   5
      TabNavigation   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "LIST JADWAL PENIMBANGAN"
   End
   Begin MSComctlLib.StatusBar sbForm 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   6975
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   12250
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
   Begin VB.Image Image1 
      Height          =   1965
      Left            =   120
      Picture         =   "frmjadwal.frx":57FE
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   3165
   End
End
Attribute VB_Name = "frmjadwal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset, rsFind As New ADODB.Recordset
Dim sqlStr As String, bSave As Boolean
Dim iDxForm As Integer, recPointer As Integer

Dim bEditing As Boolean, bEditKey As Boolean
Dim cNkey As String

Option Explicit

Private Sub cmdBtn_Click()
 Dim Asked As Variant, noRec As Variant
 Dim cNomer As String, cBulan As Variant
 Dim i As Byte, nAktif As Integer
 
 sqlStr = "select * from tbjadwal " & _
          "where fin='0'and st='1'"
 Set rsFind = con.Execute(sqlStr)
 If Not rsFind.EOF Then nAktif = rsFind.RecordCount
 Set rsFind = Nothing
 
 If nAktif < 5 Then
 
  cBulan = Month(Now)
  If Len(cBulan) = 1 Then
   cBulan = "0" & cBulan
  End If
 
  cNomer = Year(Now) & "/" & cBulan & "/"
  sqlStr = "select * from tbJadwal " & _
           "where left(nomer,8)='" & cNomer & "'"
  Set rsFind = con.Execute(sqlStr)
  If rsFind.EOF Then
   cNomer = cNomer & NumberToRomawi(1)
  Else
   cNomer = cNomer & NumberToRomawi(rsFind.RecordCount + 1)
  End If
  Set rsFind = Nothing
 
  noRec = rs.AbsolutePosition
  Asked = MsgBox("Sudah Mulai Bongkar ?", vbQuestion + vbYesNo + vbDefaultButton2, "JASATAMA")
  If Asked = vbYes Then
   sqlStr = "update tbJadwal set " & _
            "st='1' " & _
            ",tglbongkar='" & Format(Now, "yyyy-mm-dd HH:MM:ss") & _
            "',nomer='" & cNomer & _
            "',nodermaga='" & Trim(txtData(10).Text) & _
            "',usergrp='" & UserGroup & _
            "' where nokey='" & cNkey & "'"
   con.BeginTrans
   con.Execute sqlStr
   con.CommitTrans
  
   cmdBtn.Enabled = False
   txtData(10).Enabled = False
   rs.Requery
   rs.Move (noRec - 1), 1
   InitGrid
   mGrid.SetFocus
   mGrid.ActiveRow.Selected = True
  End If
 
 Else
  MsgBox "Sudah 5 Jadwal Timbang Yang Aktif", _
         vbOKOnly + vbInformation, "JASATAMA"
 End If
End Sub

Private Sub Form_Activate()
 Me.Width = 8520
 Me.Height = 7800
 
 iDxFrm = 5
End Sub

Private Sub Form_Load()
 iDxForm = 5
 bEditing = True
 bEditKey = True
 Set GForm(iDxForm) = frmjadwal
 
 'Open Database --------------------
  con.CursorLocation = adUseClient
  con.ConnectionTimeout = 0
  con.Open "DSN=dstimbang2"
 '----------------------------------
                   
 sqlStr = "Select * from tbjadwal where fin='0' order by nomer"
                       
 Set rs = con.Execute(sqlStr)
 If Not rs.EOF Then
   InitGrid
   bEditing = False
   bEditKey = False
   recPointer = rs.RecordCount - 1
   rs.MoveLast
   PScatter
 Else
   cNkey = Year(Date) & Month(Date) & _
          Day(Date) & Hour(Time) & _
          Minute(Time) & Second(Time)
   PBlank
 End If
 pActive
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
   MsgBox "Record Pertama", vbInformation + vbOKOnly, "JASATAMA"
  End If
 Case 2 'Next
  If recPointer <> (rs.RecordCount - 1) Then
   rs.MoveNext
   recPointer = recPointer + 1
   PScatter
  Else
   MsgBox "Record Terakhir", vbInformation + vbOKOnly, "JASATAMA"
  End If
 Case 3 'Last
  rs.MoveLast
  recPointer = (rs.RecordCount - 1)
  PScatter
 Case 4 'Add
  bEditing = True
  bEditKey = True
  PForm
 Case 5 'Edit
  bEditing = True
  PForm
 Case 6 'Del
  PDelete
  PScatter
  PForm
 Case 7 'Find
 Case 8 'Print
  'PPrint
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
  If rs.RecordCount > 0 Then
   bEditing = False
   bEditKey = False
   PScatter
   PForm
  Else
   Unload Me
  End If
 End Select
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

Private Sub pvbl_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtData_Change(Index As Integer)
  Dim strSql As String
 
  Select Case Index
  Case 0    'Barang
   strSql = "select nama from tbBarang " & _
      "where kode='" & txtData(Index).Text & "'"
   Set rsFind = con.Execute(strSql)
   If Not rsFind.EOF Then
    txtData(Index + 1).Text = rsFind.Fields("nama")
   Else
    txtData(Index + 1).Text = vbNullString
   End If
   Set rsFind = Nothing

  Case 3    'Tujuan
   strSql = "select nama from tbpt " & _
      "where kode='" & txtData(Index).Text & _
      "' and stp='1'"
   Set rsFind = con.Execute(strSql)
   If Not rsFind.EOF Then
    txtData(Index - 1).Text = rsFind.Fields("nama")
   Else
    txtData(Index - 1).Text = vbNullString
   End If
   Set rsFind = Nothing
  
  Case 5    'Barge
   strSql = "select nama from tbpt " & _
      "where kode='" & txtData(Index).Text & _
      "' and bge='1'"
   Set rsFind = con.Execute(strSql)
   If Not rsFind.EOF Then
    txtData(Index - 1).Text = rsFind.Fields("nama")
   Else
    txtData(Index - 1).Text = vbNullString
   End If
   Set rsFind = Nothing
   
  Case 7    'Tug Boat
   strSql = "select nama from tbpt " & _
      "where kode='" & txtData(Index).Text & _
      "' and tgb='1'"
   Set rsFind = con.Execute(strSql)
   If Not rsFind.EOF Then
    txtData(Index - 1).Text = rsFind.Fields("nama")
   Else
    txtData(Index - 1).Text = vbNullString
   End If
   Set rsFind = Nothing
   
  Case 8    'Pemilik
   strSql = "select nama from tbpt " & _
      "where kode='" & txtData(Index).Text & _
      "' and brg='1'"
   Set rsFind = con.Execute(strSql)
   If Not rsFind.EOF Then
    txtData(Index + 1).Text = rsFind.Fields("nama")
   Else
    txtData(Index + 1).Text = vbNullString
   End If
   Set rsFind = Nothing
  End Select
End Sub

Private Sub txtData_GotFocus(Index As Integer)
 SendKeys "{home}+{end}"
End Sub

Private Sub PForm()
 pActive
 If bEditKey Then
  PBlank
  'pvDate.SetFocus
  txtData(11).SetFocus
 Else
  If bEditing Then
   'txtdata(0).Enabled = False
   'pvDate.SetFocus
   txtData(11).SetFocus
  End If
 End If
End Sub

Public Sub pActive()
 If bEditing Then
  pActivForm
 Else
  pDeactivForm
 End If
End Sub

Private Sub pActivForm()
 mGrid.Enabled = False
 cmdBtn.Enabled = False
 txtData(10).Enabled = False
 pvDate.Enabled = True
 txtData(0).Enabled = True
 txtData(3).Enabled = True
 txtData(5).Enabled = True
 txtData(7).Enabled = True
 txtData(8).Enabled = True
 txtData(11).Enabled = True
 pvbl.Enabled = True
 
 navCtrl1.EditPos
End Sub

Private Sub pDeactivForm()
 mGrid.Enabled = True
 pvDate.Enabled = False
 txtData(0).Enabled = False
 txtData(3).Enabled = False
 txtData(5).Enabled = False
 txtData(7).Enabled = False
 txtData(8).Enabled = False
 txtData(11).Enabled = False
 pvbl.Enabled = False
 
 navCtrl1.ViewPos
End Sub

Public Sub PScatter()
 Dim i As Integer
 
 On Error Resume Next 'Jika ada Field yang bernilai NULL
 If rs.RecordCount > 0 Then
  cNkey = rs.Fields("nokey").Value
  lblnomer = rs.Fields("nomer").Value
  pvDate.Value = rs.Fields("tglsandar").Value
  txtData(1).Text = rs.Fields("barang").Value
  txtData(2).Text = rs.Fields("tujuan").Value
  txtData(4).Text = rs.Fields("barge").Value
  txtData(6).Text = rs.Fields("tugboat").Value
  txtData(9).Text = rs.Fields("pemilik").Value
  txtData(10).Text = rs.Fields("nodermaga").Value
  txtData(11).Text = rs.Fields("nomor_1b").Value
  pvbl.Text = rs.Fields("bl").Value
  If rs.Fields("st").Value = "1" Then
     cmdBtn.Enabled = False
     txtData(10).Enabled = False
  Else
     cmdBtn.Enabled = True
     txtData(10).Enabled = True
  End If
    
  mGrid.Refresh ssRefetchAndFireInitializeRow
  mGrid.ActiveRow.Selected = True
 Else
  cNkey = Year(Date) & Month(Date) & _
          Day(Date) & Hour(Time) & _
          Minute(Time) & Second(Time)
  pvDate.Value = Now
  For i = 0 To 9
   txtData(i).Text = vbNullString
  Next
  cmdBtn.Enabled = False
  txtData(10).Enabled = False
  
  bEditing = True
  bEditKey = True
 End If
End Sub

Public Sub PSave()
 Dim Asked As String
 Dim cnk As New ADODB.Connection
 
 bSave = False
 Asked = MsgBox("Data Akan Disimpan dan Pastikan Telah Waktu Sandar ?", vbQuestion + vbYesNo + vbDefaultButton2 _
         , "JASATAMA")
  If Asked = vbYes Then
     bSave = True
     sqlStr = "select * from tbJadwal " & _
              "where nokey='" & cNkey & "'"
     Set rsFind = con.Execute(sqlStr)
     If rsFind.EOF Then
       sqlStr = "insert into tbJadwal(nokey,barang,tujuan,barge,tugboat,pemilik," & _
                "tglsandar,tglbongkar,bl,nomor_1b) values('" & cNkey & "','" & _
                txtData(1).Text & "','" & txtData(2).Text & "','" & _
                txtData(4).Text & "','" & txtData(6).Text & "','" & txtData(9).Text & _
                "','" & Format(pvDate.Value, "yyyy-mm-dd HH:MM:SS") & _
                "','" & Format(pvDate.Value, "yyyy-mm-dd HH:MM:SS") & "'," & _
                pvbl.Value & ",'" & Trim(txtData(11).Text) & "')"
     Else
       sqlStr = "update tbJadwal set " & _
                "barang = '" & txtData(1).Text & _
                "',tujuan = '" & txtData(2).Text & _
                "',barge = '" & txtData(4).Text & _
                "',tugboat = '" & txtData(6).Text & _
                "',pemilik = '" & txtData(9).Text & _
                "',tglsandar = '" & Format(pvDate.Value, "yyyy-mm-dd HH:MM:SS") & _
                "',bl = " & pvbl.Value & _
                ",nomor_1b= '" & Trim(txtData(11).Text) & _
                "' where nokey = '" & cNkey & "'"
     End If
     Set rsFind = Nothing
        
     con.BeginTrans
     con.Execute sqlStr
     con.CommitTrans

     
     'ORACLE PELINDO UPDATE JADWAL =============================
      On Error Resume Next
      cnk.CursorLocation = adUseClient
      cnk.Open "dscuker", "cuker", "cuker"
     
      cnk.BeginTrans
      cnk.Execute "update rf_jadwal_kapal set " & _
         "status='1' " & _
         "where nomor_1b='" & Trim(txtData(11).Text) & "'"
      cnk.CommitTrans
     
      cnk.Close
     '==========================================================
     
     rs.Requery
     If rs("st").Value = "1" Then
      cmdBtn.Enabled = False
      txtData(10).Enabled = False
     Else
      cmdBtn.Enabled = True
      txtData(10).Enabled = True
     End If
     txtData(0).Text = vbNullString
     txtData(3).Text = vbNullString
     txtData(5).Text = vbNullString
     txtData(7).Text = vbNullString
     txtData(8).Text = vbNullString
     InitGrid
  End If
End Sub

Public Sub PDelete()
 Dim sMsg As Variant
 
 sMsg = MsgBox("Anda Yakin Telah Selesai Bongkar ?", vbYesNo + vbQuestion + vbDefaultButton2, "JASATAMA")
 If sMsg = vbYes Then
  sqlStr = "update tbjadwal set " & _
           "fin='1' " & _
           "where nokey='" & cNkey & "'"
  con.BeginTrans
  con.Execute sqlStr
  con.CommitTrans
  
  rs.Requery
  If Not rs.EOF Then rs.MoveLast
 End If
End Sub

Private Sub PBlank()
 Dim i As Integer
 
 pvDate.Value = Now
 For i = 0 To 9
  txtData(i).Text = vbNullString
 Next
 pvbl.Text = 0
 txtData(11).Text = vbNullString
 
 cNkey = Year(Date) & Month(Date) & _
         Day(Date) & Hour(Time) & _
         Minute(Time) & Second(Time)
End Sub

Private Sub InitGrid()
 Dim i As Integer
 
 Set mGrid.DataSource = rs
 For i = 0 To (rs.Fields.Count - 1)
  mGrid.Bands(0).Columns(i).Activation = ssActivationActivateOnly
 Next
 mGrid.Bands(0).Columns(0).Hidden = True
 mGrid.Bands(0).Columns(10).Hidden = True
 mGrid.Bands(0).Columns(11).Hidden = True
End Sub

Private Sub txtData_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Dim cnk As New ADODB.Connection
  
 If KeyCode = vbKeyF2 Then
  Select Case Index
  Case 0 'Barang
   sqlStr = "select kode,nama from tbbarang order by kode"
   ShowFind "DSN=dstimbang2", sqlStr, "MASTER BARANG", 1, 1
   txtData(0).Text = Scatter_Code
   txtData(1).Text = Scatter_Code1
  Case 3 'Tujuan
   sqlStr = "select kode,nama from vwstppt order by kode"
   ShowFind "DSN=dstimbang2", sqlStr, "DATA TUJUAN", 1, 1
   txtData(3).Text = Scatter_Code
   txtData(2).Text = Scatter_Code1
  Case 5 'Barge/Tongkang
   sqlStr = "select kode,nama from vwbgept order by kode"
   ShowFind "DSN=dstimbang2", sqlStr, "DAFTAR BARGE / TONGKANG", 1, 1
   txtData(5).Text = Scatter_Code
   txtData(4).Text = Scatter_Code1
  Case 7 'Tug Boat
   sqlStr = "select kode,nama from vwtgbpt order by kode"
   ShowFind "DSN=dstimbang2", sqlStr, "DAFTAR TUG BOAT", 1, 1
   txtData(7).Text = Scatter_Code
   txtData(6).Text = Scatter_Code1
  Case 8 'Pemilik Barang
   sqlStr = "select kode,nama from vwbrgpt order by kode"
   ShowFind "DSN=dstimbang2", sqlStr, "DATA PEMILIK BARANG", 1, 1
   txtData(8).Text = Scatter_Code
   txtData(9).Text = Scatter_Code1
  Case 11   'Jadwal Dari Pelindo
   On Error Resume Next
   cnk.CursorLocation = adUseClient
   cnk.Open "dscuker", "cuker", "cuker"
   
   sqlStr = "select nomor_1b,nama_barge,nama_tugboat,est_tgl_sandar " & _
            "from rf_jadwal_kapal " & _
            "where status='0'"
   ShowFind cnk.ConnectionString, sqlStr, "Jadwal PELINDO", 1, 1, 2
   txtData(11).Text = Scatter_Code
   lblbge.Caption = Scatter_Code1
   lbltboat.Caption = Scatter_Code2
   
   cnk.Close
  End Select
  Scatter_Code = vbNullString
  Scatter_Code1 = vbNullString
  Scatter_Code2 = vbNullString
 End If
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
  SendKeys vbTab
 End If
End Sub
