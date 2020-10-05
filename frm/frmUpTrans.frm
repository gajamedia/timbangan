VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0B63DF3F-CC00-4D55-A1C9-CAFE70BB1B49}#1.0#0"; "xpctrl.ocx"
Begin VB.Form frmUpTrans 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Transaksi Timbang"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8910
   HasDC           =   0   'False
   Icon            =   "frmUpTrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   3255
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtData 
         Height          =   300
         Index           =   5
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   15
         Top             =   2280
         Width           =   1335
      End
      Begin XPCtrl.XPButton cmdBtn 
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         Top             =   2780
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&UPDATE"
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
         MICON           =   "frmUpTrans.frx":57E2
         PICN            =   "frmUpTrans.frx":57FE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   2160
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   1
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtData 
         Height          =   300
         Index           =   0
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "No.Polisi"
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
         Left            =   4440
         TabIndex        =   16
         Top             =   2325
         Width           =   735
      End
      Begin VB.Label lblReg 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "Pt.Truk"
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
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblReg 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "Lambung"
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
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "Pemilik Barang"
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
         TabIndex        =   12
         Top             =   1725
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
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
         Height          =   225
         Left            =   240
         TabIndex        =   11
         Top             =   640
         Width           =   570
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         X1              =   0
         X2              =   7440
         Y1              =   2680
         Y2              =   2680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
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
         Left            =   240
         TabIndex        =   10
         Top             =   1365
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "Barge/Tongkang"
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
         TabIndex        =   9
         Top             =   1005
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00D8E9EC&
         BackStyle       =   0  'Transparent
         Caption         =   "Nomer Timbang"
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
         Top             =   270
         Width           =   1350
      End
   End
   Begin MSComctlLib.StatusBar sbForm 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   3510
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   13097
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
      Height          =   3255
      Left            =   120
      Picture         =   "frmUpTrans.frx":598B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmUpTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rsFind As New ADODB.Recordset
Dim sqlStr As String, cMasuk

Option Explicit

Private Sub cmdBtn_Click()
 If Trim(txtData(5).Text) <> vbNullString Then
  sqlStr = "update tbtrans set " & _
           "nomer='" & Trim(txtData(0).Text) & _
           "',nopol='" & Trim(txtData(5).Text) & _
           "' where wmasuk='" & Format(cMasuk, "yyyy-mm-dd HH:MM:SS") & "'"
  con.BeginTrans
  con.Execute sqlStr
  con.CommitTrans
  Unload Me
 End If
End Sub

Private Sub Form_Load()
 Dim i As Integer, cNol As String
 
 With frmMtimbang.mGrid.ActiveRow
  cMasuk = .Cells(11).Value
  txtData(0).Text = .Cells(0).Value
  txtData(1).Text = .Cells(4).Value
  txtData(2).Text = .Cells(5).Value
  txtData(3).Text = .Cells(6).Value
  txtData(4).Text = .Cells(3).Value
  
  lblReg(0).Caption = .Cells(9).Value
  lblReg(1).Caption = .Cells(8).Value
  txtData(5).Text = .Cells(10).Value
 End With
 
 'Open Database --------------------
  con.CursorLocation = adUseClient
  con.ConnectionTimeout = 0
  con.Open "DSN=dstimbang2"
 '----------------------------------
End Sub

Private Sub Form_Unload(Cancel As Integer)
 con.Close
End Sub

Private Sub txtData_GotFocus(Index As Integer)
 txtData(Index).BackColor = &HC0FFC0
 SendKeys "{home}+{end}"
End Sub

Private Sub txtData_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then
  Select Case Index
  Case 0    'Jadwal Timbang
   sqlStr = "select nomer,tujuan,barge,tugboat,pemilik from tbJadwal "
   ShowFind "DSN=dstimbang2", sqlStr, "DATA JADWAL", 0, 1, 2, 3, 4
   txtData(0).Text = Scatter_Code
   txtData(1).Text = Scatter_Code1
   txtData(2).Text = Scatter_Code2
   txtData(3).Text = Scatter_Code3
   txtData(4).Text = Scatter_Code4
  End Select
 End If
End Sub

Private Sub txtData_LostFocus(Index As Integer)
 txtData(Index).BackColor = &HFFFFFF
End Sub
