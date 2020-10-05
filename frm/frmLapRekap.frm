VERSION 5.00
Object = "{0B63DF3F-CC00-4D55-A1C9-CAFE70BB1B49}#1.0#0"; "XPCTRL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmLapRekap 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekap Timbang (per Nomer)"
   ClientHeight    =   2070
   ClientLeft      =   15
   ClientTop       =   90
   ClientWidth     =   6165
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmLapRekap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   1815
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin XPCtrl.XPButton cmd 
         Height          =   495
         Left            =   2640
         TabIndex        =   6
         Top             =   1200
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
         MICON           =   "frmLapRekap.frx":57E2
         PICN            =   "frmLapRekap.frx":57FE
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
         ItemData        =   "frmLapRekap.frx":6210
         Left            =   1680
         List            =   "frmLapRekap.frx":6238
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Height          =   320
         Index           =   0
         Left            =   3045
         MaxLength       =   4
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Height          =   320
         Index           =   1
         Left            =   1680
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin Crystal.CrystalReport CR 
         Left            =   2160
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowBorderStyle=   1
         WindowControlBox=   -1  'True
         WindowMaxButton =   0   'False
         WindowMinButton =   0   'False
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCancelBtn=   0   'False
         WindowShowProgressCtls=   0   'False
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   4200
         Y1              =   1080
         Y2              =   1080
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   645
         Width           =   570
      End
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   120
      Picture         =   "frmLapRekap.frx":629F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmLapRekap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cNomer As String, nbulan
Dim cBulan As String

Option Explicit

Private Sub cmbData_Click()
 nbulan = cmbData.ListIndex + 1
End Sub

Private Sub cmd_Click()
 If Len(Trim(Str(nbulan))) = 1 Then
  cBulan = "0" & Trim(Str(nbulan))
 Else
  cBulan = Trim(Str(nbulan))
 End If
 
 cNomer = Trim(txtData(0).Text) & "/" & Trim(cBulan) & "/" & Trim(txtData(1).Text)

 CR.ReportFileName = App.Path & "\rpt\rptrekap.rpt"
 CR.SelectionFormula = "{vwtrans.nomer} = '" & cNomer & "'"
 CR.Action = 1
End Sub

Private Sub Form_Load()
 cmbData.ListIndex = Month(Date) - 1
 txtData(0).Text = Year(Date)
End Sub
