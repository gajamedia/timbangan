VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmkjadwal 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jadwal Penimbangan Dari PELINDO"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11550
   FillColor       =   &H00C00000&
   Icon            =   "frmkjadwal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton cmdexport 
         Caption         =   "Export To SIT"
         Height          =   375
         Left            =   9720
         TabIndex        =   2
         Top             =   3120
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2970
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Klik Dobel Untuk Edit || Del untuk Hapus Data"
         Top             =   120
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   5239
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nomor IB"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Kode Pelanggan/Pemilik"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Pelanggan / Pemilik Barang"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "kode barge"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nama Barge"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Kode TugBoat"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Nama TugBoat"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "No.Dmg"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Est.Tgl. Sandar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Act. Tgl. Sandar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "KD.Tujuan"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Tujuan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Status"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "BL"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmkjadwal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private myList1 As ListItem
Dim cnk As New ADODB.Connection
Dim cn As New ADODB.Connection
Dim cmdJadwal As New ADODB.Command
Dim rsJadwal As New ADODB.Recordset
Dim nItem As Integer, cNkey As String

Private Sub RefreshList()
 Dim i As Integer

 nChecked = 0: nItem = 0
 cmdJadwal.CommandType = adCmdText
 cmdJadwal.CommandText = _
    "SELECT * FROM rf_jadwal_kapal " & _
    "WHERE status='0' " & _
    "order by nomor_1b"
 cmdJadwal.ActiveConnection = cnk
 Set rsJadwal = cmdJadwal.Execute
 If rsJadwal.RecordCount > 0 Then
  nItem = rsJadwal.RecordCount
  rsJadwal.MoveFirst
  ListView1.ListItems.Clear
  For i = 1 To rsJadwal.RecordCount
   On Error Resume Next
   Set myList1 = ListView1.ListItems.Add(, , rsJadwal("nomor_1b").Value)
   myList1.SubItems(1) = rsJadwal("kd_pelanggan").Value
   myList1.SubItems(2) = rsJadwal.Fields("nm_pelanggan").Value
   myList1.SubItems(3) = rsJadwal("kd_barge").Value
   myList1.SubItems(4) = rsJadwal("nama_barge").Value
   myList1.SubItems(5) = rsJadwal("kd_tugboat").Value
   myList1.SubItems(6) = rsJadwal("nama_tugboat").Value
   myList1.SubItems(7) = rsJadwal("no_dermaga").Value
   myList1.SubItems(8) = rsJadwal("est_tgl_sandar").Value
   myList1.SubItems(9) = rsJadwal("act_tgl_sandar").Value
   myList1.SubItems(10) = rsJadwal("kd_tujuan").Value
   myList1.SubItems(11) = rsJadwal("nm_tujuan").Value
   myList1.SubItems(13) = rsJadwal("tonase_bl").Value
   rsJadwal.MoveNext
  Next i
 Else
  ListView1.ListItems.Clear
 End If
 Set rsJadwal = Nothing
End Sub

Private Sub cmdexport_Click()
 Dim sqlStr As String, itemIdx As Integer
 Dim nomor_1b As String, cKdagen As String
 Dim nmAgen As String, cKdpemilik As String
 Dim nmPemilik As String, cKdTb As String
 Dim nmTb As String, cKdBarge As String
 Dim nmBarge As String, dTglSandar As Date
 Dim nBL As Variant
  
 itemIdx = 1
 If MsgBox("Jadwal Akan Di Ekspor ke Database Jasatama ?", _
    vbYesNo + vbInformation + vbDefaultButton2, _
    "PT.Gresik Jasatama") = vbYes Then
  For itemIdx = 1 To nItem
    
   If ListView1.ListItems(itemIdx).Checked Then
    cNkey = Year(Date) & Month(Date) & _
            Day(Date) & Hour(Time) & _
            Minute(Time) & (Second(Time) + itemIdx)
  
    nomor_1b = ListView1.ListItems(itemIdx).Text
    cKdagen = ListView1.ListItems(itemIdx).ListSubItems(11).Text
    cKdpemilik = ListView1.ListItems(itemIdx).ListSubItems(2).Text
    cKdTb = ListView1.ListItems(itemIdx).ListSubItems(6).Text
    cKdBarge = ListView1.ListItems(itemIdx).ListSubItems(4).Text
    dTglSandar = ListView1.ListItems(itemIdx).ListSubItems(9).Text
    nBL = Val(ListView1.ListItems(itemIdx).ListSubItems(13).Text)
  
    sqlStr = "insert into tbJadwal(nokey,barang,tujuan,barge,tugboat,pemilik," & _
             "tglsandar,tglbongkar,nomor_1b,bl) values('" & cNkey & "','BATUBARA','" & _
             cKdagen & "','" & cKdBarge & "','" & cKdTb & "','" & _
             cKdpemilik & "','" & Format(dTglSandar, "yyyy-mm-dd hh:mm:ss") & _
             "','" & Format(dTglSandar, "yyyy-mm-dd hh:mm:ss") & _
             "','" & nomor_1b & "'," & nBL & ")"
             
    cn.Execute sqlStr
    
    cmdJadwal.CommandType = adCmdText
    cmdJadwal.CommandText = _
       "UPDATE rf_jadwal_kapal " & _
       "SET status='1' " & _
       "WHERE nomor_1b='" & nomor_1b & "'"
    cmdJadwal.ActiveConnection = cnk
    Set rsJadwal = cmdJadwal.Execute
    Set rsJadwal = Nothing
   End If
  Next
 End If
 RefreshList
End Sub

Private Sub Form_Load()
 On Error Resume Next
 cnk.CursorLocation = adUseClient
 cnk.Open "dscuker", "cuker", "cuker"

 cn.CursorLocation = adUseClient
 cn.Open "DSN=dstimbang2"

 RefreshList
End Sub

Private Sub Form_Unload(Cancel As Integer)
 cn.Close
 'cnk.Close
End Sub
