VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGrupUser 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Group User"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "frmGrupUser.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   9750
   Begin sit30.navCtrl navCtrl1 
      Height          =   735
      Left            =   1920
      TabIndex        =   6
      Top             =   7440
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1296
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   9495
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin MSComctlLib.TreeView MenuTree 
         Height          =   3750
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   9090
         _ExtentX        =   16034
         _ExtentY        =   6615
         _Version        =   393217
         LineStyle       =   1
         Style           =   6
         Checkboxes      =   -1  'True
         SingleSel       =   -1  'True
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSAdodcLib.Adodc adoMain 
         Height          =   375
         Left            =   3840
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=dsRetail"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "dsRetail"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Group"
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
         Top             =   240
         Width           =   1050
      End
   End
   Begin MSComctlLib.StatusBar sbForm 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   8310
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   14579
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
   Begin UltraGrid.SSUltraGrid sGrid 
      Height          =   2535
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4471
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   2
      LayoutFlags     =   67108868
      BorderStyle     =   5
      TabNavigation   =   1
      Caption         =   "LIST GROUP USER"
   End
   Begin VB.Image Image1 
      Height          =   2540
      Left            =   120
      Picture         =   "frmGrupUser.frx":57E2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3060
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   8415
      Left            =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmGrupUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset, rsFind As New ADODB.Recordset
Dim sqlStr As String
Dim iDxForm As Integer, recPointer As Integer
Dim Menu1 As String, Menu2 As String
Dim Menu3 As String, Menu4 As String, Menu5 As String
Dim bEditing As Boolean, bEditKey As Boolean, bSave As Boolean

Dim oNode As Node

Option Explicit

Private Sub Form_Activate()
 Me.Width = 9840
 Me.Height = 9135
 
 MenuTree.Nodes(1).Expanded = True
 iDxFrm = 1
End Sub

Private Sub Form_Load()
 iDxForm = 1
 bEditing = True
 bEditKey = True
 Set GForm(iDxForm) = frmGrupUser
 
 'Open Database --------------------
  con.CursorLocation = adUseClient
  con.ConnectionTimeout = 0
  con.Open "DSN=dstimbang2"
 '----------------------------------
                   
 sqlStr = "Select * from tbGrup order by cKode"
 
 Set rs = con.Execute(sqlStr)
 If Not rs.EOF Then
  InitGrid
  pLoadMenu
  bEditing = False
  bEditKey = False
  recPointer = rs.RecordCount - 1
  rs.MoveLast
  PScatter
 Else
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
   MsgBox "Record Awal", vbInformation + vbOKOnly, "JASATAMA"
  End If
 Case 2 'Next
  If recPointer <> (rs.RecordCount - 1) Then
   rs.MoveNext
   recPointer = recPointer + 1
   PScatter
  Else
   MsgBox "Record Akhir", vbInformation + vbOKOnly, "JASATAMA"
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

Public Sub pActive()
 If bEditing Then
  pActivForm
 Else
  pDeactivForm
 End If
End Sub

Private Sub PForm()
 pActive
 If bEditKey Then
  PBlank
  Text1.SetFocus
 Else
  If bEditing Then
   Text1.Enabled = False
   MenuTree.SetFocus
  End If
 End If
End Sub

Private Sub pActivForm()
 Text1.Enabled = True
 MenuTree.Enabled = True
 sGrid.Enabled = False
 navCtrl1.EditPos
End Sub

Private Sub pDeactivForm()
 Text1.Enabled = False
 MenuTree.Enabled = False
 sGrid.Enabled = True
 
 navCtrl1.ViewPos
End Sub

Public Sub PSave()
 Dim Asked As String
 
 bSave = False
 Asked = MsgBox("Anda Yakin Data Akan Disimpan ?", vbQuestion + vbYesNo, "JASATAMA")
 If Text1.Text = vbNullString Then
   MsgBox "Nama Grup Harus Diisi", vbOKOnly + vbInformation, "JASATAMA"
   Text1.SetFocus
 Else
  If Asked = vbYes Then
    bSave = True
    Call pAfter_Save
    sqlStr = "select * from tbgrup where ckode='" & Trim(Text1.Text) & "'"
    Set rsFind = con.Execute(sqlStr)
    If rsFind.EOF Then
      sqlStr = "insert into tbgrup values ('" & Trim(Text1.Text) & _
               "','" & Menu1 & "','" & Menu2 & "','" & _
               Menu3 & "','" & Menu4 & "','" & Menu5 & "')"
    Else
      sqlStr = "update tbgrup set cMenu1='" & Menu1 & _
               "',cMenu2='" & Menu2 & "',cMenu3='" & Menu3 & _
               "',cMenu4='" & Menu4 & "',cMenu5='" & Menu5 & _
               "' where cKode='" & Trim(Text1.Text) & "'"
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

Public Sub pBefore_Save()
End Sub

Public Sub pAfter_Save()
    Dim i As Integer, cLvl As String

    Menu1 = ""
    Menu2 = ""
    Menu3 = ""
    Menu4 = ""
    Menu5 = ""
    For i = 1 To MenuTree.Nodes.Count
        cLvl = Left(MenuTree.Nodes(i).Key, 2)
        Select Case cLvl
            Case "m1"
                If MenuTree.Nodes(i).Checked Then
                    Menu1 = Menu1 & "1"
                Else
                    Menu1 = Menu1 & "0"
                End If
            Case "m2"
                If MenuTree.Nodes(i).Checked Then
                    Menu2 = Menu2 & "1"
                Else
                    Menu2 = Menu2 & "0"
                End If
            Case "m3"
                If MenuTree.Nodes(i).Checked Then
                    Menu3 = Menu3 & "1"
                Else
                    Menu3 = Menu3 & "0"
                End If
            Case "m4"
                If MenuTree.Nodes(i).Checked Then
                    Menu4 = Menu4 & "1"
                Else
                    Menu4 = Menu4 & "0"
                End If
            Case "m5"
                If MenuTree.Nodes(i).Checked Then
                    Menu5 = Menu5 & "1"
                Else
                    Menu5 = Menu5 & "0"
                End If
        End Select
    Next
End Sub

Public Sub PScatter()
 Dim i As Integer
 
 On Error Resume Next 'Jika ada Field yang bernilai NULL
 If rs.RecordCount > 0 Then
  Text1.Text = rs.Fields("cKode").Value
  Scatter_Memvar
   
  sGrid.Refresh ssRefetchAndFireInitializeRow
  sGrid.ActiveRow.Selected = True
 Else
  Text1.Text = vbNullString
  bEditing = True
  bEditKey = True
 End If

End Sub

Public Sub Scatter_Memvar()
 'Text1.Text = adoMain.Recordset.Fields("cKode")
 pViewMenu
End Sub

'Subroutine untuk Check Uncheck Child Menu
Sub Child_Recursive(vIndex, vStatus As Integer)
Dim Child_Index As Integer
    MenuTree.Nodes(vIndex).Checked = vStatus
    If MenuTree.Nodes(vIndex).Children Then
        MenuTree.Nodes(vIndex).Checked = vStatus
        Child_Index = MenuTree.Nodes(vIndex).Child.Index
        Call Child_Recursive(Child_Index, vStatus)
    End If
    While vIndex <> MenuTree.Nodes(vIndex).LastSibling.Index
        vIndex = MenuTree.Nodes(vIndex).Next.Index
        If MenuTree.Nodes(vIndex).Children Then
            MenuTree.Nodes(vIndex).Checked = vStatus
            Child_Index = MenuTree.Nodes(vIndex).Child.Index
            Call Child_Recursive(Child_Index, vStatus)
        End If
        MenuTree.Nodes(vIndex).Checked = vStatus
    Wend
End Sub

'subroutine untuk Check Uncheck Parent Menu
Sub Parent_Recursive(vIndex, vStatus As Integer)
Dim i, Parent_Index As Integer
    Select Case vStatus
    Case 0
        If vIndex <> MenuTree.Nodes(vIndex).Root.Index Then
            i = MenuTree.Nodes(MenuTree.Nodes(vIndex).Parent.Index).Child.Index
            Do While i <> MenuTree.Nodes(vIndex).LastSibling.Index
                If MenuTree.Nodes(i).Checked Or MenuTree.Nodes(vIndex).LastSibling.Checked Then
                    MenuTree.Nodes(i).Parent.Checked = True
                    Exit Do
                Else
                    MenuTree.Nodes(i).Parent.Checked = False
                    i = MenuTree.Nodes(i).Next.Index
                End If
            Loop
            Parent_Index = MenuTree.Nodes(i).Parent.Index
            Call Parent_Recursive(Parent_Index, 0)
        End If
    Case 1
        MenuTree.Nodes(vIndex).Checked = True
        If vIndex <> MenuTree.Nodes(vIndex).Root.Index Then
            Parent_Index = MenuTree.Nodes(vIndex).Parent.Index
            Call Parent_Recursive(Parent_Index, 1)
        End If
    End Select
End Sub

Private Sub pLoadMenu()
 Set rsFind = con.Execute("select * from tbmenu order by menu_id")
 Set oNode = MenuTree.Nodes.Add(, , "m0", "Menu")
 MenuTree.Nodes(1).Checked = False
 While Not rsFind.EOF
   If Trim(rsFind.Fields("Menu_Level").Value) = "m0" Then
     Set oNode = MenuTree.Nodes.Add("m0", tvwChild, Trim(rsFind.Fields("Menu_ID").Value), Trim(rsFind.Fields("Menu_Text").Value))
   Else
     Set oNode = MenuTree.Nodes.Add(Trim(rsFind.Fields("Menu_Level").Value), tvwChild, Trim(rsFind.Fields("Menu_ID").Value), Trim(rsFind.Fields("Menu_Text").Value))
   End If
   rsFind.MoveNext
 Wend
 Set rsFind = Nothing
End Sub

Private Sub pViewMenu()
    Dim i As Integer, cLvl As String, vMenu As Boolean
    Dim nMenu1 As Integer, nMenu2 As Integer, nMenu3 As Integer
    Dim nMenu4 As Integer, nMenu5 As Integer
    Dim Menu1 As String, Menu2 As String, Menu3 As String
    Dim Menu4 As String, Menu5 As String
    nMenu1 = 1
    nMenu2 = 1
    nMenu3 = 1
    nMenu4 = 1
    nMenu5 = 1
    vMenu = False
    
        Menu1 = Trim(rs.Fields("cMenu1").Value)
        Menu2 = Trim(rs.Fields("cMenu2").Value)
        Menu3 = Trim(rs.Fields("cMenu3").Value)
        Menu4 = Trim(rs.Fields("cMenu4").Value)
        Menu5 = Trim(rs.Fields("cMenu5").Value)
        For i = 1 To MenuTree.Nodes.Count
            cLvl = Mid(MenuTree.Nodes(i).Key, 1, 2)
            Select Case cLvl
            Case "m1"
                If Mid(Trim(Menu1), nMenu1, 1) = "1" Then
                    MenuTree.Nodes(i).Checked = True
                    vMenu = True
                Else
                    MenuTree.Nodes(i).Checked = False
                End If
                nMenu1 = nMenu1 + 1
            Case "m2"
                If Mid(Trim(Menu2), nMenu2, 1) = "1" Then
                    vMenu = True
                    MenuTree.Nodes(i).Checked = True
                Else
                    MenuTree.Nodes(i).Checked = False
                End If
                nMenu2 = nMenu2 + 1
            Case "m3"
                If Mid(Trim(Menu3), nMenu3, 1) = "1" Then
                    vMenu = True
                    MenuTree.Nodes(i).Checked = True
                Else
                    MenuTree.Nodes(i).Checked = False
                End If
                nMenu3 = nMenu3 + 1
            Case "m4"
                If Mid(Trim(Menu4), nMenu4, 1) = "1" Then
                    vMenu = True
                    MenuTree.Nodes(i).Checked = True
                Else
                    MenuTree.Nodes(i).Checked = False
                End If
                nMenu4 = nMenu4 + 1
            Case "m5"
                If Mid(Trim(Menu5), nMenu5, 1) = "1" Then
                    vMenu = True
                    MenuTree.Nodes(i).Checked = True
                Else
                    MenuTree.Nodes(i).Checked = False
                End If
                nMenu5 = nMenu5 + 1
            End Select
        Next
        MenuTree.Nodes(1).Checked = True
        MenuTree.Nodes(1).Expanded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
  con.Close
End Sub

Private Sub MenuTree_GotFocus()
  SendKeys "{home}"
End Sub

Private Sub MenuTree_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim idx, i As Integer
    
    'kalo checkbox nyala
    If Node.Checked = True Then
        'kalo salah satu child nyala maka parent juga nyala
        If Node.Index <> Node.Root.Index Then
            Node.Parent.Checked = True
            idx = Node.Parent.Index
            Call Parent_Recursive(idx, 1)
        End If
        'kalo checkbox parent nyala maka semua child juga nyala
        If Node.Children Then
            idx = Node.Child.Index
            Node.Child.LastSibling.Checked = True
            Call Child_Recursive(idx, 1)
        End If
    'kalo checkbox mati
    Else
        'kalo semua checkbox child mati maka parent juga mati
        If Node.Index <> Node.Root.Index Then
            idx = Node.Index
            Call Parent_Recursive(idx, 0)
        End If
        'kalo checkbox parent mati maka semua child juga mati
        If Node.Children Then
            idx = Node.Child.Index
            Node.Child.LastSibling.Checked = False
            Call Child_Recursive(idx, 0)
        End If
    End If
End Sub

Private Sub PBlank()
 Text1.Text = vbNullString
 MenuTree.Nodes.Clear
 pLoadMenu
End Sub

Public Sub PDelete()
 Dim sMsg As Variant
 
 sMsg = MsgBox("Anda Yakin Data Dihapus ?", vbYesNo + vbQuestion + vbDefaultButton2, "JASATAMA")
 If sMsg = vbYes Then
  sqlStr = "delete from tbGrup where cKode='" & _
           Trim(Text1.Text) & "'"
  con.BeginTrans
  con.Execute sqlStr
  con.CommitTrans
  
  rs.Requery
  If Not rs.EOF Then rs.MoveLast
  sGrid.SetFocus
 End If
End Sub

Private Sub InitGrid()
 Dim i As Integer
 
 Set sGrid.DataSource = rs
 sGrid.Bands(0).Columns(0).Activation = ssActivationActivateOnly
 
 For i = 1 To 5
  sGrid.Bands(0).Columns(i).Hidden = True
 Next
End Sub

Private Sub sGrid_Click()
 If Not rs.EOF Then
  sGrid.ActiveRow.Selected = True
  PScatter
 End If
End Sub

Private Sub sGrid_KeyUp(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
 sGrid_Click
End Sub

Private Sub Text1_GotFocus()
 SendKeys "{home}+{end}"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  SendKeys vbTab
 Else
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
 End If
End Sub

Private Sub Text1_LostFocus()
 If bEditing Then
   rsFind.Open "select * from tbgrup", con.ConnectionString, adOpenKeyset, adLockOptimistic
   rsFind.Find "cKode='" & Text1.Text & "'"
   If Not rsFind.EOF Then
    MsgBox "Data Sudah Ada, " & vbCrLf & _
           "Coba Ketik Lagi", vbInformation + vbOKOnly, "JASATAMA"
    Text1.SetFocus
   End If
   rsFind.Close
 End If
End Sub
