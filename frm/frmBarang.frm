VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBarang 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Barang"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmBarang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6270
   Visible         =   0   'False
   Begin sit30.navCtrl navCtrl1 
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1296
   End
   Begin MSComctlLib.StatusBar sbForm 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2565
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8440
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
      Height          =   1455
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang"
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
         TabIndex        =   4
         Top             =   750
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Barang"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1245
      End
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   120
      Picture         =   "frmBarang.frx":57E2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   0
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "frmBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset, rsFind As New ADODB.Recordset
Dim sqlStr As String
Dim iDxForm As Integer, recPointer As Integer

Dim bEditing As Boolean, bEditKey As Boolean, bSave As Boolean

Option Explicit

Private Sub Form_Activate()
 Me.Height = 3390
 Me.Width = 6360
 
 iDxFrm = 2
End Sub

Private Sub Form_Load()
 iDxForm = 2
 bEditing = True
 bEditKey = True
 Set GForm(iDxForm) = frmBarang
  
 'Open Database --------------------
  con.CursorLocation = adUseClient
  con.ConnectionTimeout = 0
  con.Open "DSN=dstimbang2"
 '----------------------------------
 sqlStr = "Select * from tbBarang order by kode"
 
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
 Else
  For i = 0 To 1
   txtData(i).Text = vbNullString
  Next
  bEditing = True
  bEditKey = True
 End If
End Sub

Private Sub PBlank()
 Dim i As Integer
 
 For i = 0 To 1
  txtData(i).Text = vbNullString
 Next
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
 navCtrl1.EditPos
End Sub

Private Sub pDeactivForm()
 Dim i As Integer
 
 For i = 0 To 1
  txtData(i).Enabled = False
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
   Set rsFind = con.Execute("select * from tbBarang " & _
                   "where kode ='" & Trim(txtData(0).Text) & "'")
   If Not rsFind.EOF Then
    bFound = True
   End If
   Set rsFind = Nothing
  
   con.BeginTrans
   If bFound Then
    sqlStr = "update tbBarang set " & _
             "nama='" & Trim(txtData(1).Text) & _
             "' where kode ='" & Trim(txtData(0).Text) & "'"
   Else
    sqlStr = "insert into tbBarang values('" & _
             Trim(txtData(0).Text) & "','" & Trim(txtData(1).Text) & "')"
   End If
   con.Execute sqlStr
   con.CommitTrans
  
   rs.Requery
  End If
 Else
  MsgBox "Kode Barang Tidak Boleh Kosong", vbOKOnly + vbInformation, "JASATAMA"
  txtData(0).SetFocus
 End If
End Sub

Private Sub PDelete()
 Dim sMsg As Variant
 
 sMsg = MsgBox("Anda Yakin Data Dihapus ?", vbYesNo + vbQuestion + vbDefaultButton2, "JASATAMA")
 If sMsg = vbYes Then
  sqlStr = "delete from tbBarang " & _
           "where kode='" & Trim(txtData(0).Text) & "'"
  con.BeginTrans
  con.Execute sqlStr
  con.CommitTrans
  
  rs.Requery
  If Not rs.EOF Then rs.MoveLast
 End If
End Sub
