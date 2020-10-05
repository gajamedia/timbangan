VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{CFFE0A60-8E3A-11D3-BCC0-00104B9E0792}#1.0#0"; "SSInput1.ocx"
Begin VB.Form frmUser 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting User"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6255
   Begin sit30.navCtrl navCtrl1 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1296
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   1455
      Left            =   2160
      TabIndex        =   3
      Top             =   2520
      Width           =   3975
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin ActiveInput.SSComboBoxEx cmbData 
         DataField       =   "cKode"
         DataSource      =   "adoGroup"
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   65536
         BorderColor     =   134217736
         BorderStyle     =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "cmbData"
         ListField       =   "cKode"
         Style           =   2
         TimeFormat      =   1
         Sorted          =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama User"
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
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Grup"
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
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
   End
   Begin UltraGrid.SSUltraGrid mGrid 
      Height          =   2295
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
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
      Caption         =   "LIST NAMA USER"
   End
   Begin MSComctlLib.StatusBar sbForm 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   4950
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8414
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
      Height          =   3900
      Left            =   120
      Picture         =   "frmUser.frx":57E2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1980
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset, rsFind As New ADODB.Recordset
Dim xField(3) As String, sqlStr As String
Dim iDxForm As Integer, recPointer As Integer
Dim fEncrypt As New Encrypt

Dim bEditing As Boolean, bEditKey As Boolean
Dim bSave As Boolean, bOut As Boolean

Option Explicit

Private Sub Isi_Grup()
 Dim i As Integer
 
 i = 0
 cmbData.Clear
 sqlStr = "select cKode from tbGrup"
 Set rsFind = con.Execute(sqlStr)
 If Not rsFind.EOF Then
  rsFind.MoveFirst
  Do
   cmbData.AddItem rsFind.Fields("cKode").Value, i
   i = i + 1
   rsFind.MoveNext
  Loop Until rsFind.EOF
 Else
  MsgBox "Data Grup Belum Ada Silahkan Isi Grup User Terlebih Dahulu", vbOKOnly + vbCritical, "JASATAMA"
  bOut = True
 End If
 Set rsFind = Nothing
End Sub

Private Sub Form_Activate()
 Me.Width = 6345
 Me.Height = 5775
 
 iDxFrm = 0
 
 If bOut Then
  bOut = False
  Unload Me
 End If
End Sub

Private Sub Form_Load()
 iDxForm = 0
 bEditing = True
 bEditKey = True
 Set GForm(iDxForm) = frmUser
 
 'Open Database --------------------
  con.CursorLocation = adUseClient
  con.ConnectionTimeout = 0
  con.Open "DSN=dstimbang2"
 '----------------------------------
                   
 Isi_Grup
 If Not bOut Then
  sqlStr = "Select * from tbUser order by user_Name"
                       
  Set rs = con.Execute(sqlStr)
  If Not rs.EOF Then
   InitGrid
   bEditing = False
   bEditKey = False
   recPointer = rs.RecordCount - 1
   rs.MoveLast
   PScatter
  Else
   PBlank
  End If
  pActive
 End If
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

Private Sub txtData_GotFocus(Index As Integer)
 SendKeys "{home}+{end}"
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

Public Sub pActive()
 If bEditing Then
  pActivForm
 Else
  pDeactivForm
 End If
End Sub

Private Sub pActivForm()
 mGrid.Enabled = False
 txtData(0).Enabled = True
 txtData(1).Enabled = True
 cmbData.Enabled = True
 
 navCtrl1.EditPos
End Sub

Private Sub pDeactivForm()
 mGrid.Enabled = True
 txtData(0).Enabled = False
 txtData(1).Enabled = False
 cmbData.Enabled = False
 
 navCtrl1.ViewPos
End Sub

Public Sub PScatter()
 Dim i As Integer
 
 On Error Resume Next 'Jika ada Field yang bernilai NULL
 If rs.RecordCount > 0 Then
  txtData(0).Text = rs.Fields(0).Value
  txtData(1).Text = rs.Fields(2).Value
  cmbData.OverrideText = rs.Fields(1).Value
  
  mGrid.Refresh ssRefetchAndFireInitializeRow
  mGrid.ActiveRow.Selected = True
 Else
  For i = 0 To 2
   txtData(i).Text = vbNullString
  Next
  cmbData.ListIndex = 0
  bEditing = True
  bEditKey = True
 End If
End Sub

Public Sub PSave()
 Dim usrPass As Variant
 Dim Asked As String
 
 bSave = False
 Asked = MsgBox("Anda Yakin Data Akan Disimpan ?", vbQuestion + vbYesNo, "JASATAMA")
 If txtData(0).Text = vbNullString Then
        MsgBox "Nama User Harus Diisi", vbOKOnly + vbInformation, "JASATAMA"
        txtData(0).SetFocus
 Else
  If Asked = vbYes Then
     bSave = True
     usrPass = fEncrypt.ChgPass(txtData(1).Text, 1)
     sqlStr = "select * from tbUser " & _
              "where User_Name='" & Trim(txtData(0).Text) & "'"
     Set rsFind = con.Execute(sqlStr)
     If rsFind.EOF Then
         sqlStr = "insert into tbUser values('" & _
                 Trim(txtData(0).Text) & "','" & _
                 Trim(cmbData.Text) & "','" & _
                 usrPass & "')"
     Else
         sqlStr = "update tbUser set " & _
             "User_Pass = '" & usrPass & _
             "',User_Grup = '" & Trim(cmbData.OverrideText) & _
             "' where User_Name = '" & Trim(txtData(0).Text) & "'"
     End If
     Set rsFind = Nothing
        
     con.BeginTrans
     con.Execute sqlStr
     con.CommitTrans
        
     rs.Requery
     InitGrid
  End If
 End If
End Sub

Public Sub PDelete()
 Dim sMsg As Variant
 
 sMsg = MsgBox("Anda Yakin Data Dihapus ?", vbYesNo + vbQuestion + vbDefaultButton2, "PT.AGUNG SEDAYU")
 If sMsg = vbYes Then
  sqlStr = "delete from tbUser " & _
           "where User_Name='" & Trim(txtData(0).Text) & "'"
  con.BeginTrans
  con.Execute sqlStr
  con.CommitTrans
  
  rs.Requery
  If Not rs.EOF Then rs.MoveLast
 End If
End Sub

Private Sub PBlank()
 Dim i As Integer
 
 For i = 0 To 1
  txtData(i).Text = vbNullString
 Next
 cmbData.ListIndex = 0
End Sub

Private Sub InitGrid()
 Dim i As Integer
 
 Set mGrid.DataSource = rs
 For i = 0 To 2
  mGrid.Bands(0).Columns(i).Activation = ssActivationActivateOnly
 Next
 mGrid.Bands(0).Columns(2).Hidden = True
 mGrid.Bands(0).Columns(0).Width = 2000
 mGrid.Bands(0).Columns(0).Header.Caption = "NAMA USER"
 mGrid.Bands(0).Columns(1).Header.Caption = "GRUP USER"
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
  SendKeys vbTab
 End If
End Sub
