VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{0B63DF3F-CC00-4D55-A1C9-CAFE70BB1B49}#1.0#0"; "xpctrl.ocx"
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Object = "{C2000000-FFFF-1100-8100-000000000001}#8.0#0"; "PVCurr.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmtimbang 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi Penimbangan"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8205
   Icon            =   "frmtimbang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   8205
   Begin MSWinsockLib.Winsock sckHover 
      Index           =   6
      Left            =   7200
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3000
   End
   Begin MSWinsockLib.Winsock sckHover 
      Index           =   5
      Left            =   6720
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3000
   End
   Begin MSWinsockLib.Winsock sckHover 
      Index           =   4
      Left            =   6240
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3000
   End
   Begin MSWinsockLib.Winsock sckHover 
      Index           =   3
      Left            =   5760
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3000
   End
   Begin MSWinsockLib.Winsock sckHover 
      Index           =   2
      Left            =   5280
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3000
   End
   Begin MSWinsockLib.Winsock sckHover 
      Index           =   1
      Left            =   4800
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3000
   End
   Begin MSWinsockLib.Winsock sckHover 
      Index           =   0
      Left            =   4320
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3000
   End
   Begin MSWinsockLib.Winsock sckGateOut 
      Left            =   4800
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3000
   End
   Begin MSWinsockLib.Winsock sckGateIn 
      Left            =   4320
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3000
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   1320
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1320
      Top             =   5760
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      Height          =   1815
      Left            =   120
      TabIndex        =   24
      Top             =   6600
      Width           =   7935
      Begin PVCurrencyLib.PVCurrency pvcur 
         Height          =   330
         Index           =   2
         Left            =   5640
         TabIndex        =   12
         Top             =   600
         Width           =   2175
         _Version        =   524288
         _ExtentX        =   3836
         _ExtentY        =   582
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Enabled         =   0   'False
         Appearance      =   0
         BorderStyle     =   1
         Alignment       =   2
         ReadOnly        =   -1  'True
         FormatNegative  =   4
         Symbol          =   ""
         DecimalPlaces   =   "0"
         DecimalSeparator=   ","
         Value           =   0
      End
      Begin PVCurrencyLib.PVCurrency pvcur 
         Height          =   330
         Index           =   1
         Left            =   5640
         TabIndex        =   11
         Top             =   240
         Width           =   2175
         _Version        =   524288
         _ExtentX        =   3836
         _ExtentY        =   582
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Appearance      =   1
         Alignment       =   2
         FormatNegative  =   1
         Symbol          =   ""
         DecimalPlaces   =   "0"
         DecimalSeparator=   ","
         Value           =   0
      End
      Begin PVCurrencyLib.PVCurrency pvcur 
         Height          =   330
         Index           =   3
         Left            =   5640
         TabIndex        =   13
         Top             =   960
         Width           =   2175
         _Version        =   524288
         _ExtentX        =   3836
         _ExtentY        =   582
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Enabled         =   0   'False
         Appearance      =   0
         BorderStyle     =   1
         Alignment       =   2
         FormatNegative  =   1
         Symbol          =   ""
         DecimalPlaces   =   "0"
         DecimalSeparator=   ","
         Value           =   0
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   1080
         TabIndex        =   10
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   1080
         TabIndex        =   9
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1080
         TabIndex        =   8
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
      Begin XPCtrl.XPButton cmdBtn 
         Height          =   375
         Index           =   1
         Left            =   5400
         TabIndex        =   14
         Top             =   1365
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Simpan"
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
         MICON           =   "frmtimbang.frx":57E2
         PICN            =   "frmtimbang.frx":57FE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin XPCtrl.XPButton cmdBtn 
         Height          =   375
         Index           =   2
         Left            =   6600
         TabIndex        =   15
         Top             =   1365
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Batal"
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
         MICON           =   "frmtimbang.frx":590F
         PICN            =   "frmtimbang.frx":592B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         X1              =   4080
         X2              =   4080
         Y1              =   120
         Y2              =   1800
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Kotor"
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
         Left            =   4320
         TabIndex        =   31
         Top             =   315
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Kosong"
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
         Left            =   4320
         TabIndex        =   30
         Top             =   690
         Width           =   1155
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Berat Muatan"
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
         Left            =   4320
         TabIndex        =   29
         Top             =   1065
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TugBoat"
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
         TabIndex        =   28
         Top             =   1335
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barge"
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
         TabIndex        =   27
         Top             =   975
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tujuan"
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
         TabIndex        =   26
         Top             =   615
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barang"
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
         TabIndex        =   25
         Top             =   255
         Width           =   690
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   4245
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   150
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Armada Pengangkut"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   3975
      Begin PVCurrencyLib.PVCurrency pvcur 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   1815
         _Version        =   524288
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         Alignment       =   2
         FormatNegative  =   1
         Symbol          =   ""
         DecimalPlaces   =   "0"
         DecimalSeparator=   ","
         Value           =   0
      End
      Begin VB.TextBox txtData 
         Height          =   300
         Index           =   0
         Left            =   1560
         TabIndex        =   0
         Top             =   915
         Width           =   1335
      End
      Begin VB.TextBox txtrfid 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "T00000000"
         Top             =   360
         Width           =   3375
      End
      Begin XPCtrl.XPButton cmdBtn 
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   4
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Masuk"
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
         MICON           =   "frmtimbang.frx":5A2C
         PICN            =   "frmtimbang.frx":5A48
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
         Index           =   2
         Left            =   1560
         TabIndex        =   2
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1560
         TabIndex        =   1
         Top             =   1305
         Width           =   1335
      End
      Begin VB.TextBox txthrfid 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "T00000000"
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
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
         Left            =   2040
         TabIndex        =   23
         Top             =   2205
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         X1              =   0
         X2              =   3960
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perusahaan"
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
         TabIndex        =   21
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Polisi"
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
         TabIndex        =   20
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No.Lambung"
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
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1275
      Left            =   120
      MousePointer    =   2  'Cross
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "123456"
      ToolTipText     =   "Click Layar Ini untuk Isi Nilai Tara/Bruto"
      Top             =   120
      Width           =   3975
   End
   Begin MSComctlLib.StatusBar sbForm 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   8580
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   11853
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   4200
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Klik Dobel Untuk Edit || Del untuk Hapus Data"
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Lambung"
         Object.Width           =   1500
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No.Polisi"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Tara"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "nokey"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "wmasuk"
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "wmuat"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "rfid"
         Object.Width           =   0
      EndProperty
   End
   Begin UltraGrid.SSUltraGrid mGrid 
      Height          =   2295
      Left            =   2160
      TabIndex        =   6
      Top             =   4200
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4048
      _Version        =   131072
      GridFlags       =   17040388
      UpdateMode      =   2
      LayoutFlags     =   67108884
      BorderStyle     =   5
      TabNavigation   =   1
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
      Caption         =   "JADWAL TIMBANG"
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   360
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   4800
      ParitySetting   =   2
      DataBits        =   7
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   120
      Picture         =   "frmtimbang.frx":5AFC
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1935
   End
End
Attribute VB_Name = "frmtimbang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cons As New ADODB.Connection
Dim cmdOra As New ADODB.Command
Dim rsOra As New ADODB.Recordset

Dim con As New ADODB.Connection
Dim rsJadwal As New ADODB.Recordset, rsAngkut As New ADODB.Recordset
Dim rsFind As New ADODB.Recordset

Dim sqlStr As String, cNoKey As String
Dim nRecList As Variant, lastclick As String
Private myList1 As ListItem, cMode As String
Dim cNomer As String, cParity As String
Dim tmplambung As String, cPemilik As String
Dim dMasuk As Variant, dKeluar As Variant
Dim refreshrate As Long, cnDmg As String
Dim dMuat As Variant, cNo1B As String
Dim dsisaUper As Variant, refreshrate2 As Long

Dim cRFID As String                             'Gate In/Out
Dim chRFID(6) As String, noderHover As String   'Hover
Dim sCetakIn As Boolean, sCetakOut As Boolean
Dim fEncrypt As New Encrypt

Option Explicit

Private Sub BukaPort()
 On Error Resume Next
 
 If MSComm1.PortOpen Then MSComm1.PortOpen = False
 'Inisialisasi Port --------
  MSComm1.CommPort = ncom
  If parity = "0" Then
   cParity = "e"
  Else
   If parity = "1" Then
    cParity = "n"
   Else
    cParity = "o"
   End If
  End If
  MSComm1.Settings = Trim(Str(baudrate)) & "," & Trim(cParity) & "," & Trim(databit) & "," & Trim(stopbit)
  MSComm1.InputMode = comInputModeBinary
 '--------------------------
 
 MSComm1.PortOpen = True        'Port Timbangan
 
 MSComm2.PortOpen = True        'Port LED DISPLAY
 clear_LED
End Sub

Private Sub RefreshList()
 Dim i As Integer
 Dim tmp_listtview As ListItem
 
 sqlStr = "Select nolambung,nopol,tara,nokey,wmasuk,wmuat,rfid " & _
          "From tbtrans " & _
          "Where maskel='0' " & _
          "order by wmasuk desc"
 Set rsAngkut = con.Execute(sqlStr)
 If rsAngkut.RecordCount > 0 Then
  nRecList = rsAngkut.RecordCount
  rsAngkut.MoveFirst
  ListView1.ListItems.Clear
  
  For i = 1 To rsAngkut.RecordCount
   Set myList1 = ListView1.ListItems.Add(, , rsAngkut.Fields(0).Value)
   myList1.SubItems(1) = rsAngkut.Fields(1).Value
   myList1.SubItems(2) = rsAngkut.Fields(2).Value
   myList1.SubItems(3) = IIf(IsNull(rsAngkut.Fields(3).Value), "", rsAngkut.Fields(3).Value)
   myList1.SubItems(4) = rsAngkut.Fields(4).Value
   myList1.SubItems(5) = rsAngkut.Fields(5).Value
   myList1.SubItems(6) = rsAngkut.Fields(6).Value
   rsAngkut.MoveNext
  Next i
      
    Set tmp_listtview = ListView1.FindItem(lastclick, lvwSubItem)
    If Not tmp_listtview Is Nothing Then
        tmp_listtview.EnsureVisible
        tmp_listtview.Selected = True
    End If
 Else
  ListView1.ListItems.Clear
  nRecList = 0
 End If
 Set rsAngkut = Nothing
End Sub

Private Sub cmdBtn_Click(Index As Integer)
 Dim i As Byte, spdata As String
 Dim liccount As String
 
 gBruto = "0": gTara = "0": gNetto = "0"
 If cMode <> "1" Then
  gTara = pvcur(0).Text
 Else
  gBruto = pvcur(1).Text
  gTara = pvcur(2).Text
  
  pvcur(3).Text = pvcur(1).Value - pvcur(2).Value
  gNetto = pvcur(3).Text
 End If
 
 gLambung = txtData(0).Text:  gNopol = txtData(1).Text
 gBarang = txtData(3).Text: gPemilik = cPemilik
 gNomer = cNomer: gDermaga = cnDmg
 gRFID = txtrfid.Text
 
 'Baca Lisensi ----------------------------
  spdata = App.Path & "\lic.gmd"
  liccount = FileRead(spdata, True)(2)
  liccount = fEncrypt.ChgPass(liccount, 2)
 '-----------------------------------------
 
 Select Case Index
 Case 0 'insert data awal timbang (Gate In) dan cetak
  cNoKey = Year(Date) & Month(Date) & _
           Day(Date) & Hour(Time) & _
           Minute(Time) & Second(Time)
  dMasuk = Now: gMasuk = dMasuk
  
  'Cek Lisensi Program
  '---------------------------------------------------------
   If Val(liccount) = 1 Then
    pvcur(0).Text = 0
    gTara = "0"
   End If
  '---------------------------------------------------------
  
  If MsgBox(Cetak_Struk2(App.Path & "\rpt\strukin.txt"), vbYesNo + vbDefaultButton2, "JASATAMA - Cetak Struk In ?") = vbYes Then
   sqlStr = "insert into tbtrans(nokey,nolambung,nopol,tara,wmasuk,wmuat,rfid) values('" & _
            cNoKey & "','" & Trim(txtData(0).Text) & "','" & Trim(txtData(1).Text) & _
            "'," & pvcur(0).Value & ",'" & Format(dMasuk, "yyyy-mm-dd HH:MM:SS") & _
            "','" & Format(dMasuk, "yyyy-mm-dd HH:MM:SS") & "','" & Left(Trim(gRFID), 9) & "')"
   con.BeginTrans
   con.Execute sqlStr
   con.CommitTrans
  
   'Refresh ListView  ----
    RefreshList
    cmdBtn(0).Enabled = False
   '-------------------------------------
   
   sCetakIn = True
   Cetak_Struk App.Path & "\rpt\strukin.txt"
   sCetakOut = False
  End If
  
 Case 1, 2 '1:update data timbang (Gate Out) | 2:Batal Timbang Keluar
  
  'Cek Lisensi Program
  '---------------------------
   If Val(liccount) = 1 Then
    pvcur(1).Text = 0
    gBruto = "0"
   End If
  '----------------------------
  
  'Prosedur Update Data Timbang --------------------------------------
   If Index = 1 Then
    gMasuk = dMasuk
    dKeluar = Now: gKeluar = dKeluar
    If MsgBox(Cetak_Struk2(App.Path & "\rpt\strukout.txt"), vbYesNo + vbDefaultButton2, "JASATAMA - Cetak Struk Out ?") = vbYes Then
     sqlStr = "update tbtrans set " & _
              "nomer='" & cNomer & _
              "',wkeluar='" & Format(dKeluar, "yyyy-mm-dd HH:MM:SS") & _
              "',bruto=" & pvcur(1).Value & _
              ",maskel='1' " & _
              ",nodermaga='" & cnDmg & _
              "',usergrp='" & UserGroup & _
              "',nomor_1b='" & cNo1B & _
              "' where nokey='" & cNoKey & "'"
     con.BeginTrans
     con.Execute sqlStr
     con.CommitTrans
    
     
     On Error Resume Next
     'Prosedur entry insert database ke oracle
     '----------------------------------------
     If Trim(cnDmg) <> "MX" And _
        Trim(cnDmg) <> "LB" And _
        Trim(cnDmg) <> "0" Then
      cmdOra.CommandType = adCmdText
      cmdOra.CommandText = _
        "INSERT INTO rf_truk_transaksi(rfid,nomor_1b,no_lambung,no_polisi," & _
        "nama_armada,tgl_masuk,tgl_muat,tgl_keluar,bruto," & _
        "tara,no_dermaga,truck_id,nomer_gjt,sink,nokey_gjt) " & _
        "VALUES ('" & txtrfid.Text & "','" & cNo1B & "','" & txtData(0).Text & _
        "','" & txtData(1).Text & "','" & txtData(2).Text & _
        "',to_date('" & Format(dMasuk, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')" & _
        ",to_date('" & Format(dMuat, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')" & _
        ",to_date('" & Format(dKeluar, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')" & _
        "," & pvcur(1).Value & "," & pvcur(2).Value & ",'" & cnDmg & _
        "','" & cNoKey & "','" & cNomer & "','0','" & cNoKey & "')"
      cmdOra.ActiveConnection = cons
      Set rsOra = cmdOra.Execute
      Set rsOra = Nothing
     End If
     '----------------------------------------
              
'     sCetakOut = True
     For i = 0 To 1
      Cetak_Struk App.Path & "\rpt\strukout.txt"
     Next
        
     RefreshList
     
     
'     sCetakOut = False
    End If
   End If
  '-------------------------------------------------------------------
  
  'Enable Frame Angkutan ----
   txtData(0).Enabled = True
   pvcur(0).Enabled = True
  '---------------------------
  
  'Disable Frame Bawah -------
   cmdBtn(1).Enabled = False
   cmdBtn(2).Enabled = False
   mGrid.Enabled = False
   pvcur(1).Enabled = False
   
   On Error Resume Next
   mGrid.ActiveRow.Selected = False
  '--------------------------
 End Select
 
 cMode = "0"
 'Clear Field -------------------
  For i = 0 To 3
   pvcur(i).Text = "0"
  Next
  For i = 0 To 6
   txtData(i).Text = vbNullString
  Next
  sbForm.Panels(1).Text = vbNullString
  txtrfid.Text = "T00000000"
  cNo1B = ""
 '-------------------------------
 txtData(0).SetFocus
 txtData(0).BackColor = &HC0FFC0
End Sub

Private Sub Form_Activate()
 If ncom = -1 Then
  Unload Me
  MsgBox "Port Data Timbangan Belum di Setting", vbOKOnly + vbCritical, "JASATAMA"
 End If
End Sub

Private Sub Form_Load()
 Dim i As Integer
 
 'Open Database --------------------
  'MySQl
  con.CursorLocation = adUseClient
  con.ConnectionTimeout = 0
  con.Open "DSN=dstimbang2"
  
  'Oracle
  On Error Resume Next
  cons.CursorLocation = adUseClient
  cons.ConnectionTimeout = 0
  cons.Open "dscuker", "cuker", "cuker"
 '----------------------------------
 
 sqlStr = "select * from tbJadwal Where st='1' and fin='0' order by nomer"
 Set rsJadwal = con.Execute(sqlStr)
 InitGrid
 RefreshList
 If ncom <> -1 Then BukaPort
 refreshrate = GetTickCount
 refreshrate2 = GetTickCount
  
 'RFID Section -----------------------------------------------------------
  If baudrate = "9600" Then     'Gate Out
   If sckGateOut.State = 0 Then
    sckGateOut.Close
    sckGateOut.LocalPort = 0
   End If
  
   sckGateOut.Connect "192.168.1.204", 3000
  Else                            'Gate In
   If sckGateIn.State = 0 Then
    sckGateIn.Close
    sckGateIn.LocalPort = 0
   End If
    
   sckGateIn.Connect "192.168.1.205", 3000
  End If
  
  'Hover
  For i = 0 To 6
   If sckHover(i).State = 0 Then
    sckHover(i).Close
    sckHover(i).LocalPort = 0
   End If
   sckHover(i).Connect "192.168.1.22" & Trim(Str(i + 1)), 3000
  Next
 
 '------------------------------------------------------------------------
End Sub

Private Sub Form_Unload(Cancel As Integer)
 clear_LED
 Set rsJadwal = Nothing
 If MSComm1.PortOpen Then MSComm1.PortOpen = False
 If MSComm2.PortOpen Then MSComm2.PortOpen = False
 con.Close
 'cons.Close
End Sub

Private Sub InitGrid()
 Dim i As Integer
 
 Set mGrid.DataSource = rsJadwal
 For i = 0 To (rsJadwal.Fields.Count - 1)
  mGrid.Bands(0).Columns(i).Activation = ssActivationActivateOnly
 Next
 mGrid.Bands(0).Columns(0).Hidden = True
 mGrid.Bands(0).Columns(2).Width = 500
 mGrid.Bands(0).Columns(3).Hidden = True
 mGrid.Bands(0).Columns(10).Hidden = True
 mGrid.Bands(0).Columns(11).Hidden = True
End Sub

Private Sub ListView1_Click()
 If nRecList > 0 Then
  lastclick = ListView1.SelectedItem.ListSubItems(3).Text
  refreshrate = GetTickCount
 End If
End Sub

Private Sub ListView1_DblClick()
 Dim sRFID As String
 
 'pilih data di listview untuk timbang keluar
 If nRecList > 0 Then
  cMode = "1"   'mode timbang keluar
  'Disable Frame Angkutan ----
   pvcur(0).Text = "0"
   txtData(0).Enabled = False
   pvcur(0).Enabled = False
  '---------------------------
  cmdBtn(2).Enabled = True 'Enable Tombol Batal
  mGrid.Enabled = True  'Enable List Jadwal
  
  txtData(0).Text = ListView1.ListItems(ListView1.SelectedItem.Index).Text
  pvcur(2).Value = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(2).Text
  cNoKey = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(3).Text
  dMasuk = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(4).Text
  dMuat = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(5).Text
  
  sRFID = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(6).Text
  If sRFID = vbNullString Then
   txtrfid.Text = "T00000000"
  Else
   txtrfid.Text = sRFID
  End If
 End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
 Dim tmp_listtview As ListItem
  
 If (Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z") Or _
    (Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z") Or _
    (Chr(KeyAscii) >= "1" And Chr(KeyAscii) <= "9") Or _
    Chr(KeyAscii) = "0" Then
    tmplambung = tmplambung + Chr(KeyAscii)
    Set tmp_listtview = ListView1.FindItem(tmplambung, lvwSubItem, , lvwPartial)
    If Not tmp_listtview Is Nothing Then
        tmp_listtview.EnsureVisible
        tmp_listtview.Selected = True
    End If
 End If
 If KeyAscii = 8 Then
    If Text2.Visible Then tmplambung = Left(tmplambung, Len(tmplambung) - 1)
 End If
 If KeyAscii = 13 Then
    If Trim(Text2.Text) = Trim(ListView1.ListItems(ListView1.SelectedItem.Index).Text) Then _
       Call ListView1_DblClick
 End If
 If KeyAscii = 27 Then
    tmplambung = vbNullString
    Text2.Visible = False
 End If

 If Len(tmplambung) > 0 Then
    Text2.Text = UCase$(tmplambung)
    Text2.Visible = True
    Text2.Width = Len(tmplambung) * 140
 Else
    Text2.Visible = False
 End If
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
 Dim ask As Variant, no As String

 If KeyCode = vbKeyDelete And _
    nRecList > 0 And cMode <> "1" Then
  ask = MsgBox("Apakah Data Akan Dihapus ?", vbYesNo + vbQuestion + vbDefaultButton2, "JASATAMA")
  If ask = vbYes Then
   no = ListView1.ListItems(ListView1.SelectedItem.Index).ListSubItems(3).Text
   
   sqlStr = "delete from tbtrans " & _
            "where nokey='" & no & "'"
   con.BeginTrans
   con.Execute sqlStr
   con.CommitTrans
   
   RefreshList
  End If
 End If
End Sub

Private Sub mGrid_Click()
 If Not rsJadwal.EOF Then
  mGrid.ActiveRow.Selected = True
  
  cNomer = rsJadwal("nomer").Value
  txtData(3).Text = rsJadwal("barang").Value
  txtData(4).Text = rsJadwal("tujuan").Value
  txtData(5).Text = rsJadwal("barge").Value
  txtData(6).Text = rsJadwal("tugboat").Value
  cPemilik = rsJadwal("pemilik").Value
  cnDmg = rsJadwal("nodermaga").Value
  cNo1B = rsJadwal("nomor_1b").Value
  
  'Enable Frame Bawah -------
   If pvcur(1).Value > 0 Then cmdBtn(1).Enabled = True
   pvcur(1).Enabled = True
  '--------------------------
  
  refreshrate2 = GetTickCount
  
  'Prosedur Mencari Sisa Uper dan Tampilkan
  '------------------------------------------
'    cmdOra.CommandType = adCmdText
'    cmdOra.CommandText = _
       "select uper_setor,uper_terpakai " & _
       "from rf_jadwal_kapal " & _
       "where nomor_1b='" & cNo1B & "'"
'    cmdOra.ActiveConnection = cons
'    Set rsOra = cmdOra.Execute
'    If Not rsOra.EOF Then
'     dsisaUper = rsOra("uper_setor").Value - rsOra("uper_terpakai").Value
'     MSComm2.Output = "A4488" & "Rp. " & Trim(Str(dsisaUper)) & Chr(13)
'    End If
'    Set rsOra = Nothing
  '------------------------------------------
 End If
End Sub

Private Sub MSComm1_OnComm()
 Dim Buffer, sNominal As String
 Dim i As Byte
   
 On Error Resume Next
    
 Text1.Text = vbNullString
 Buffer = MSComm1.Input
 Buffer = Mid(StrConv(Buffer, vbUnicode), 1, 18)
 
 If Right(Buffer, 2) = vbCrLf Then
   Text1.Text = Format(Mid$(Buffer, 9, 6), "00000#")
   If Left(Buffer, 2) = "ST" Then
     Text1.Enabled = True
   Else
     Text1.Enabled = False
   End If
 End If
End Sub

Private Sub pvcur_Change(Index As Integer)
 On Error Resume Next
 
 Select Case Index
 Case 0     'TARA
  If pvcur(0).Value > 0 Then
   MSComm2.Output = "A3488" & Trim(pvcur(0).Text) & " KG" & vbCr
   cmdBtn(0).Enabled = True
  Else
   cmdBtn(0).Enabled = False
  End If
 Case 1     'BRUTO
  If pvcur(1).Value > 0 Then
   If pvcur(1).Value >= pvcur(2).Value Then
    pvcur(3).Value = pvcur(1).Value - pvcur(2).Value
    cmdBtn(1).Enabled = True
   Else
    pvcur(3).Text = "0"
    cmdBtn(1).Enabled = False
   End If
  Else
   cmdBtn(1).Enabled = False
  End If
 Case 3     'NETTO / MUATAN
  If pvcur(3).Value > 0 Then _
    MSComm2.Output = "A3488" & Trim(pvcur(3).Text) & " KG" & vbCr
 End Select
End Sub

Private Sub pvcur_GotFocus(Index As Integer)
 pvcur(Index).BackColor = &HC0FFC0
End Sub

Private Sub pvcur_lostFocus(Index As Integer)
 pvcur(Index).BackColor = &HFFFFFF
End Sub

'Prosedur Baca RFID Saat Gate Out ---------------------------------
Private Sub RFID_IN()
  
  On Error Resume Next
  
  Dim itmlv As ListItem, cNoJd As String
   
  cNoJd = vbNullString
  sqlStr = "select nolambung,nodermaga from tbtrans " & _
           "where left(rfid,9)='" & txtrfid.Text & _
           "' and maskel='0'"
  Set rsFind = con.Execute(sqlStr)
  If Not rsFind.EOF Then
   tmplambung = rsFind("nolambung").Value
   cNoJd = rsFind("nodermaga").Value
   
   With ListView1
    Set itmlv = .FindItem(tmplambung, lvwText, , lvwPartial)
    If Not itmlv Is Nothing Then
     .ListItems(itmlv.Index).Selected = True
     .SetFocus
    End If
   End With
   Set itmlv = Nothing
   
   Call ListView1_DblClick
      
  End If
  Set rsFind = Nothing
  
  If cNoJd <> vbNullString Or cNoJd <> "0" Then
   rsJadwal.MoveFirst
   rsJadwal.Find "nodermaga = '" & cNoJd & "'"
   If Not rsJadwal.EOF Then
    mGrid.ActiveRow.Selected = True
    mGrid_Click
   End If
  End If
  
End Sub   '-----------------------------------------------------------

'Baca RFID Gate In
Private Sub sckGateIn_DataArrival(ByVal bytesTotal As Long)
 Dim gjt As String
 
' If sCetakIn = False Then
  sckGateIn.GetData gjt$, vbString
  cRFID = Trim(Right(cRFID & gjt$, 10))
  txtrfid.Text = Left(cRFID, 9)
' Else
'  MsgBox "Sistem Sedang Mencetak.......", vbInformation + vbOKOnly, "JASATAMA"
' End If
End Sub

'Baca RFID Gate Out
Private Sub sckGateOut_DataArrival(ByVal bytesTotal As Long)
 Dim gjt As String
 
' If sCetakOut = False Then
  sckGateOut.GetData gjt$, vbString
  cRFID = Trim(Right(cRFID & gjt$, 10))
  txtrfid.Text = Left(cRFID, 9)
 
  If Len(Trim(txtrfid.Text)) = 9 Then RFID_IN
' Else
'  MsgBox "Sistem Sedang Mencetak.......", vbInformation + vbOKOnly, "JASATAMA"
' End If
End Sub

'Prosedur untuk proses pencatatan waktu muat dan dermaga
'Melalui RFID di Hover
'---------------------------------------------------------------------
Private Sub RFID_Hover(cIP As String, ckRFID As String)
 Dim nKey As String
 
 On Error Resume Next
 
 sqlStr = "select noder from tbhover " & _
          "where noip='" & cIP & "'"
 Set rsFind = con.Execute(sqlStr)
 If Not rsFind.EOF Then
  noderHover = rsFind("noder").Value
 Else
  noderHover = "0"
 End If
 Set rsFind = Nothing
 
 sqlStr = "select nokey from tbtrans " & _
          "where rfid='" & ckRFID & _
          "' and maskel='0'"
 Set rsFind = con.Execute(sqlStr)
 If Not rsFind.EOF Then
  nKey = rsFind("nokey").Value
 Else
  nKey = vbNullString
 End If
 Set rsFind = Nothing
 
 If nKey <> vbNullString Then
  sqlStr = "update tbtrans set " & _
           "nodermaga='" & noderHover & _
           "',wmuat='" & Format(Now, "yyyy-mm-dd HH:MM:SS") & _
           "' where nokey='" & nKey & "'"
  con.Execute sqlStr
 End If
End Sub '----------------------------------------------------------------------

'Baca RFID Hover
Private Sub sckHover_DataArrival(Index As Integer, ByVal bytesTotal As Long)
 Dim gjt As String
 
 sckHover(Index).GetData gjt$, vbString
 chRFID(Index) = Trim(Right(chRFID(Index) & gjt$, 10))
 txthrfid.Text = Left(chRFID(Index), 9)
 
 If Len(Trim(txthrfid.Text)) = 9 Then RFID_Hover sckHover(Index).RemoteHostIP, Trim(txthrfid.Text)
End Sub

Private Sub Text1_Click()
 If cMode = "1" Then
  pvcur(1).Text = Text1.Text    'Bruto (Out)
 Else
  pvcur(0).Text = Text1.Text    'Tara (In)
 End If
End Sub

Private Sub Timer1_Timer()
 MSComm1_OnComm
 
 If (GetTickCount - refreshrate > 25000) Then
  RefreshList
  refreshrate = GetTickCount
  tmplambung = vbNullString
  Text2.Visible = False
 End If
 
 If (GetTickCount - refreshrate2 > 15000) Then
  rsJadwal.Requery
  InitGrid
  refreshrate2 = GetTickCount
 End If
 
 On Error Resume Next
 'If cNo1B <> "" Then
  'Prosedur Mencari Sisa Uper dan Tampilkan
  '------------------------------------------
    cmdOra.CommandType = adCmdText
    cmdOra.CommandText = _
       "select uper_setor,uper_terpakai " & _
       "from rf_jadwal_kapal " & _
       "where nomor_1b='" & cNo1B & "'"
    cmdOra.ActiveConnection = cons
    Set rsOra = cmdOra.Execute
    If Not rsOra.EOF Then
     dsisaUper = rsOra("uper_setor").Value - rsOra("uper_terpakai").Value
     MSComm2.Output = "A4488" & "Rp. " & Trim(Str(dsisaUper)) & Chr(13)
    End If
    Set rsOra = Nothing
  '------------------------------------------
 'End If
End Sub

Private Sub txtData_Change(Index As Integer)
 On Error Resume Next
 
 Select Case Index
 Case 0     'Data Truk
   sqlStr = "select * from tbtruk " & _
            "where nolambung='" & Trim(txtData(0).Text) & "'"
   Set rsAngkut = con.Execute(sqlStr)
   If Not rsAngkut.EOF Then
    'If cMode <> "1" Then cmdBtn(0).Enabled = True
    txtData(1).Text = rsAngkut("nopolisi").Value
    txtData(2).Text = rsAngkut("nmpt").Value
    
    MSComm2.Output = "A1488" & Trim(txtData(0).Text) & vbCr
    MSComm2.Output = "A2488" & Trim(txtData(1).Text) & vbCr
   Else
    'cmdBtn(0).Enabled = False
    txtData(1).Text = vbNullString
    txtData(2).Text = vbNullString
   End If
   Set rsAngkut = Nothing
            
 Case 6        'Nama Kapal
   MSComm2.Output = "A5488" & Trim(txtData(6).Text) & vbCr
 End Select
End Sub

Private Sub txtData_GotFocus(Index As Integer)
 SendKeys "{home}+{end}"
 txtData(Index).BackColor = &HC0FFC0
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

Private Sub clear_LED()
 On Error Resume Next
 
 MSComm2.Output = "A1488" & "TERMINAL" & Chr(13)
 MSComm2.Output = "A2488" & "CURAH" & Chr(13)
 MSComm2.Output = "A3488" & "KERING" & Chr(13)
 MSComm2.Output = "A4488" & "PT. GRESIK" & Chr(13)
 MSComm2.Output = "A5488" & "JASATAMA" & Chr(13)
End Sub
