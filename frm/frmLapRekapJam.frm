VERSION 5.00
Object = "{0B63DF3F-CC00-4D55-A1C9-CAFE70BB1B49}#1.0#0"; "xpctrl.ocx"
Object = "{CC696B60-4159-11D0-BDCB-0020A90B183A}#8.0#0"; "PVDateEdit9.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLapRekapJam 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekap Timbang (Per Jam)"
   ClientHeight    =   3300
   ClientLeft      =   15
   ClientTop       =   90
   ClientWidth     =   6165
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmLapRekapJam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   3015
      Left            =   1800
      TabIndex        =   12
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   1
         ItemData        =   "frmLapRekapJam.frx":57E2
         Left            =   720
         List            =   "frmLapRekapJam.frx":57EF
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CheckBox chkData 
         BackColor       =   &H000080FF&
         Caption         =   "Check1"
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         Top             =   960
         Width           =   255
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   300
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         Top             =   1335
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "99:99:99"
         PromptChar      =   "_"
      End
      Begin XPCtrl.XPButton cmd 
         Height          =   495
         Left            =   2640
         TabIndex        =   11
         Top             =   2400
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Cetak"
         ENAB            =   -1  'True
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
         MICON           =   "frmLapRekapJam.frx":57FF
         PICN            =   "frmLapRekapJam.frx":581B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   0
         ItemData        =   "frmLapRekapJam.frx":622D
         Left            =   1680
         List            =   "frmLapRekapJam.frx":6255
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Height          =   320
         Index           =   0
         Left            =   3045
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Height          =   320
         Index           =   1
         Left            =   1680
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin Crystal.CrystalReport CR 
         Left            =   2160
         Top             =   2400
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   1
         WindowControlBox=   -1  'True
         WindowMaxButton =   0   'False
         WindowMinButton =   0   'False
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowNavigationCtls=   -1  'True
         WindowShowCancelBtn=   0   'False
         WindowShowPrintBtn=   -1  'True
         WindowShowExportBtn=   -1  'True
         WindowShowZoomCtl=   -1  'True
         WindowShowProgressCtls=   0   'False
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin PVDATE2Lib.PVDate2 pvDate 
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   1320
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
         CalendarFormat  =   4
         TimeStore       =   -1  'True
         Value           =   41544.0092708333
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   300
         Index           =   1
         Left            =   3000
         TabIndex        =   6
         Top             =   1335
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "99:99:99"
         PromptChar      =   "_"
      End
      Begin PVDATE2Lib.PVDate2 pvDate 
         Height          =   300
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1335
         _Version        =   524288
         _ExtentX        =   2355
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Enabled         =   0   'False
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
         Enabled         =   0   'False
         TimeStore       =   -1  'True
         Value           =   41544.0092708333
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   300
         Index           =   2
         Left            =   1680
         TabIndex        =   8
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "99:99:99"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   300
         Index           =   3
         Left            =   3000
         TabIndex        =   9
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "hh:mm:ss"
         Mask            =   "99:99:99"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
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
         TabIndex        =   19
         Top             =   2445
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sd"
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
         Left            =   2640
         TabIndex        =   18
         Top             =   1830
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sd"
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
         Left            =   2640
         TabIndex        =   17
         Top             =   1365
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Timbang"
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
         Left            =   1680
         TabIndex        =   16
         Top             =   1005
         Width           =   1155
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
         TabIndex        =   15
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   4200
         Y1              =   2280
         Y2              =   2280
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
         TabIndex        =   14
         Top             =   285
         Width           =   1230
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
         TabIndex        =   13
         Top             =   645
         Width           =   570
      End
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   120
      Picture         =   "frmLapRekapJam.frx":62BC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmLapRekapJam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cNomer As String, nbulan
Dim cBulan As String, sqlStr As String
Dim dTgl, dTgl2

Option Explicit

Private Sub chkdata_Click()
 If chkData.Value = Unchecked Then
  pvDate(1).Enabled = False
  mskData(2).Enabled = False
  mskData(3).Enabled = False
 Else
  pvDate(1).Enabled = True
  mskData(2).Enabled = True
  mskData(3).Enabled = True
 End If
End Sub

Private Sub cmbData_Click(Index As Integer)
 Select Case Index
 Case 0
  nbulan = cmbData(0).ListIndex + 1
 End Select
End Sub

Private Sub cmd_Click()
 
 On Error Resume Next
 If Len(Trim(Str(nbulan))) = 1 Then
  cBulan = "0" & Trim(Str(nbulan))
 Else
  cBulan = Trim(Str(nbulan))
 End If
 
 dTgl = pvDate(0).Value: dTgl2 = pvDate(1).Value
 
 cNomer = Trim(txtData(0).Text) & "/" & Trim(cBulan) & "/" & Trim(txtData(1).Text)
  
 CR.Formulas(0) = "cOperator='" & UserID & "'"
 CR.Formulas(1) = "cShift='" & cmbData(1).Text & "'"

 CR.ReportFileName = App.Path & "\rpt\rptrekapjam.rpt"
   
 If chkData.Value = Unchecked Then
  CR.SelectionFormula = "{vwrptrans.nomer} = '" & cNomer & _
                       "' and ({vwrptrans.wkeluar} >= datetime(" & Year(dTgl) & "," & _
                       Month(dTgl) & "," & Day(dTgl) & "," & Hour(mskData(0).Text) & "," & _
                       Minute(mskData(0).Text) & "," & Second(mskData(0).Text) & _
                       ") and {vwrptrans.wkeluar} <= datetime(" & Year(dTgl) & "," & _
                       Month(dTgl) & "," & Day(dTgl) & "," & Hour(mskData(1).Text) & "," & _
                       Minute(mskData(1).Text) & "," & Second(mskData(1).Text) & "))"
 Else
  CR.SelectionFormula = "{vwrptrans.nomer} = '" & cNomer & _
                       "' and (({vwrptrans.wkeluar} >= datetime(" & Year(dTgl) & "," & _
                       Month(dTgl) & "," & Day(dTgl) & "," & Hour(mskData(0).Text) & "," & _
                       Minute(mskData(0).Text) & "," & Second(mskData(0).Text) & _
                       ") and {vwrptrans.wkeluar} <= datetime(" & Year(dTgl) & "," & _
                       Month(dTgl) & "," & Day(dTgl) & "," & Hour(mskData(1).Text) & "," & _
                       Minute(mskData(1).Text) & "," & Second(mskData(1).Text) & "))" & _
                       " or ({vwrptrans.wkeluar} >= datetime(" & Year(dTgl2) & "," & _
                       Month(dTgl2) & "," & Day(dTgl2) & "," & Hour(mskData(2).Text) & "," & _
                       Minute(mskData(2).Text) & "," & Second(mskData(2).Text) & _
                       ") and {vwrptrans.wkeluar} <= datetime(" & Year(dTgl2) & "," & _
                       Month(dTgl2) & "," & Day(dTgl2) & "," & Hour(mskData(3).Text) & "," & _
                       Minute(mskData(3).Text) & "," & Second(mskData(3).Text) & ")))"
 End If
  
 CR.Action = 1
End Sub

Private Sub Form_Load()
 cmbData(0).ListIndex = Month(Date) - 1
 txtData(0).Text = Year(Date)
End Sub

Private Sub txtData_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then
  Select Case Index
  Case 2 'Grup Operator
   sqlStr = "select ckode from tbgrup order by ckode"
   ShowFind "DSN=dstimbang2", sqlStr, "MASTER GROUP", 0, 0
   txtData(2).Text = Scatter_Code
  End Select
 End If
End Sub
