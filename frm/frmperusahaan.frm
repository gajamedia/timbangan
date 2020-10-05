VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPerusahaan 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Perusahaan"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmperusahaan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin sit30.navCtrl navCtrl1 
      Height          =   735
      Left            =   525
      TabIndex        =   7
      Top             =   2520
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1296
   End
   Begin MSComctlLib.StatusBar sbForm 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3405
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   9922
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
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   2295
      Left            =   1800
      TabIndex        =   8
      Top             =   120
      Width           =   5175
      Begin VB.CheckBox chkdata 
         BackColor       =   &H000080FF&
         Caption         =   "Pemilik Barang"
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CheckBox chkdata 
         BackColor       =   &H000080FF&
         Caption         =   "Stock Pile/Tujuan"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox chkdata 
         BackColor       =   &H000080FF&
         Caption         =   "Tug Boat"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox chkdata 
         BackColor       =   &H000080FF&
         Caption         =   "Barge"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   3
         Top             =   1440
         Width           =   855
      End
      Begin VB.CheckBox chkdata 
         BackColor       =   &H000080FF&
         Caption         =   "Angkutan"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         X1              =   0
         X2              =   5160
         Y1              =   980
         Y2              =   980
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "Penyedia"
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
         TabIndex        =   12
         Top             =   1035
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perusahaan"
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
         TabIndex        =   10
         Top             =   630
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Perusahaan"
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
         TabIndex        =   9
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   120
      Picture         =   "frmperusahaan.frx":57E2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmPerusahaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset, rsFind As New ADODB.Recordset
Dim sqlStr As String
Dim iDxForm As Integer, recPointer As Integer

Dim bEditing As Boolean, bEditKey As Boolean, bSave As Boolean

Dim cAgk As String, cBge As String, cTgb As String
Dim cStp As String, cBrg As String

Option Explicit

Private Sub chkdata_Click(Index As Integer)
 Select Case Index
 Case 0
  If chkData(0).Value = 1 Then cAgk = "1" Else cAgk = "0"
 Case 1
  If chkData(1).Value = 1 Then cBge = "1" Else cBge = "0"
 Case 2
  If chkData(2).Value = 1 Then cTgb = "1" Else cTgb = "0"
 Case 3
  If chkData(3).Value = 1 Then cStp = "1" Else cStp = "0"
 Case 4
  If chkData(4).Value = 1 Then cBrg = "1" Else cBrg = "0"
 End Select
End Sub

Private Sub Form_Activate()
 Me.Height = 4230
 Me.Width = 7200
 
 iDxFrm = 3
End Sub

Private Sub Form_Load()
 iDxForm = 3
 bEditing = True
 bEditKey = True
 Set GForm(iDxForm) = frmPerusahaan
 
 cAgk = "0": cBge = "0": cTgb = "0": cStp = "0": cBrg = "0"
  
 'Open Database --------------------
  con.CursorLocation = adUseClient
  con.ConnectionTimeout = 0
  con.Open "DSN=dstimbang2"
 '----------------------------------
 sqlStr = "Select * from tbPt order by kode"
 
 Set rs = con.Execute(sqlStr)
 If Not rs.EOF Then
  bEditing = False
  bEditKey = False
  recPointer = rs.RecordCount - 1
  rs.MoveLast
  PScatter
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

Private Sub PScatter()
 Dim i As Integer
 
 On Error Resume Next 'Jika ada Field yang bernilai NULL
 If rs.RecordCount > 0 Then
  For i = 0 To 1
   txtData(i).Text = rs.Fields(i).Value
  Next
  cAgk = rs.Fields(2).Value
  cBge = rs.Fields(3).Value
  cTgb = rs.Fields(4).Value
  cStp = rs.Fields(5).Value
  cBrg = rs.Fields(6).Value
  
  If cAgk = "0" Then chkData(0).Value = 0 Else chkData(0).Value = 1
  If cBge = "0" Then chkData(1).Value = 0 Else chkData(1).Value = 1
  If cTgb = "0" Then chkData(2).Value = 0 Else chkData(2).Value = 1
  If cStp = "0" Then chkData(3).Value = 0 Else chkData(3).Value = 1
  If cBrg = "0" Then chkData(4).Value = 0 Else chkData(4).Value = 1
 Else
  For i = 0 To 1
   txtData(i).Text = vbNullString
  Next
 
  For i = 0 To 4
   chkData(i).Value = 0
  Next
  cAgk = "0": cBge = "0": cTgb = "0": cStp = "0": cBrg = "0"
  
  bEditing = True
  bEditKey = True
 End If
End Sub

Private Sub PBlank()
 Dim i As Integer
 
 For i = 0 To 1
  txtData(i).Text = vbNullString
 Next
 
 For i = 0 To 4
  chkData(i).Value = 0
 Next
 cAgk = "0": cBge = "0": cTgb = "0": cStp = "0": cBrg = "0"
End Sub

Private Sub PForm()
 pActive
 If bEditKey Then
  PBlank
  txtData(0).SetFocus
 Else
  If bEditing Then
   txtData(0).Enabled = False
   txtData(1).SetFocus
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
 
 For i = 0 To 1
  txtData(i).Enabled = True
 Next
 For i = 0 To 4
  chkData(i).Enabled = True
 Next
 navCtrl1.EditPos
End Sub

Private Sub pDeactivForm()
 Dim i As Integer
 
 For i = 0 To 1
  txtData(i).Enabled = False
 Next
 For i = 0 To 4
  chkData(i).Enabled = False
 Next
 navCtrl1.ViewPos
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set rs = Nothing
 con.Close
End Sub

Private Sub txtData_GotFocus(Index As Integer)
 txtData(Index).BackColor = &HC0FFC0
 SendKeys "{home}+{end}"
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
  SendKeys vbTab
 Else
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
 End If
End Sub

Private Sub txtData_LostFocus(Index As Integer)
 txtData(Index).BackColor = &HFFFFFF
End Sub

Private Sub PSave()
 Dim sMsg As Variant, bFound As Boolean
 
 bFound = False
 bSave = False
 If Trim(txtData(0).Text) <> vbNullString Then
  sMsg = MsgBox("Apakah Data Sudah Benar ?", vbYesNo + vbQuestion + vbDefaultButton2, "JASATAMA")
  If sMsg = vbYes Then
   bSave = True
   Set rsFind = con.Execute("select * from tbPt " & _
                   "where kode ='" & Trim(txtData(0).Text) & "'")
   If Not rsFind.EOF Then
    bFound = True
   End If
   Set rsFind = Nothing
  
   con.BeginTrans
   If bFound Then
    sqlStr = "update tbPt set " & _
             "nama='" & Trim(txtData(1).Text) & _
             "',agk='" & cAgk & _
             "',bge='" & cBge & _
             "',tgb='" & cTgb & _
             "',stp='" & cStp & _
             "',brg='" & cBrg & _
             "' where kode ='" & Trim(txtData(0).Text) & "'"
   Else
    sqlStr = "insert into tbPt values('" & _
             Trim(txtData(0).Text) & "','" & Trim(txtData(1).Text) & _
             "','" & cAgk & "','" & cBge & "','" & cTgb & _
             "','" & cStp & "','" & cBrg & "')"
   End If
   con.Execute sqlStr
   con.CommitTrans
  
   rs.Requery
  End If
 Else
  MsgBox "Kode Perusahaan Tidak Boleh Kosong", vbOKOnly + vbInformation, "JASATAMA"
  txtData(0).SetFocus
 End If
End Sub

Private Sub PDelete()
 Dim sMsg As Variant
 
 sqlStr = "select * from tbtruk " & _
        "where kdpt='" & Trim(txtData(0).Text) & "'"
 Set rsFind = con.Execute(sqlStr)
 If rsFind.EOF Then
  sMsg = MsgBox("Anda Yakin Data Dihapus ?", vbYesNo + vbQuestion + vbDefaultButton2, "JASATAMA")
  If sMsg = vbYes Then
   sqlStr = "delete from tbPt " & _
            "where kode='" & Trim(txtData(0).Text) & "'"
   con.BeginTrans
   con.Execute sqlStr
   con.CommitTrans
  
   rs.Requery
   If Not rs.EOF Then rs.MoveLast
  End If
 Else
  MsgBox "Data Tidak Bisa Dihapus." & vbCrLf & _
         "Hapus Data Truknya Terlebih Dahulu.", vbOKOnly + vbCritical, "JASATAMA"
 End If
 Set rsFind = Nothing
End Sub
