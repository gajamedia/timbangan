VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCekTgl 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Tanggal dan Waktu"
   ClientHeight    =   5385
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCekTgl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOK 
      BackColor       =   &H80000001&
      Caption         =   "O K"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1560
      TabIndex        =   5
      Top             =   4440
      Width           =   1290
   End
   Begin VB.Timer tmrCurrDate 
      Interval        =   1000
      Left            =   2640
      Top             =   1920
   End
   Begin VB.Timer tmrCurrTime 
      Interval        =   1000
      Left            =   2160
      Top             =   1920
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H000080FF&
      Caption         =   "Tidak, Lakukan Perubahan Manual"
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   1125
      Width           =   2940
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H000080FF&
      Caption         =   "Tidak, Lakukan Perubahan Manual"
      Height          =   315
      Left            =   225
      TabIndex        =   2
      Top             =   2685
      Width           =   2820
   End
   Begin VB.Timer tmrBlink 
      Interval        =   500
      Left            =   1680
      Top             =   2160
   End
   Begin VB.Frame Frames 
      BackColor       =   &H000080FF&
      Enabled         =   0   'False
      Height          =   1275
      Index           =   1
      Left            =   75
      TabIndex        =   3
      Top             =   3000
      Width           =   4260
      Begin MSComCtl2.DTPicker dpTime 
         Height          =   630
         Left            =   150
         TabIndex        =   4
         Top             =   375
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1111
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16711935
         Format          =   88276994
         CurrentDate     =   38557
      End
   End
   Begin VB.Frame Frames 
      BackColor       =   &H000080FF&
      Enabled         =   0   'False
      Height          =   1155
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   1125
      Width           =   4260
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   630
         Left            =   150
         TabIndex        =   9
         Top             =   375
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1111
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16711680
         CalendarForeColor=   0
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   12632256
         CustomFormat    =   "dd MMMM  yyyy"
         Format          =   88276995
         CurrentDate     =   41160
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   75
      Picture         =   "frmCekTgl.frx":000C
      Top             =   150
      Width           =   480
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pastikan Tanggal dan Waktu Benar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Index           =   0
      Left            =   705
      TabIndex        =   8
      Top             =   120
      Width           =   3510
      WordWrap        =   -1  'True
   End
   Begin VB.Label Labels 
      BackColor       =   &H000080FF&
      Caption         =   "Sudah Benar ?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   6
      Top             =   900
      Width           =   1665
   End
   Begin VB.Label Labels 
      BackColor       =   &H000080FF&
      Caption         =   "Sudah Benar ?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   75
      TabIndex        =   7
      Top             =   2460
      Width           =   1665
   End
End
Attribute VB_Name = "frmCekTgl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnOK_Click()
    If Check1.Value = 1 Then Date = dtpDate.Value
    If Check2.Value = 1 Then Time = dpTime.Value
    Unload Me
End Sub

Private Sub Check1_Click()
    DisplayCap
    If Check1.Value = 1 Then
        Frames(0).Enabled = True
        tmrCurrDate.Enabled = False
    Else
        Frames(0).Enabled = False
        dtpDate.Value = Date
        tmrCurrDate.Enabled = True
    End If
End Sub

Private Sub Check2_Click()
    DisplayCap
    If Check2.Value = 1 Then
        Frames(1).Enabled = True
        tmrCurrTime.Enabled = False
    Else
        Frames(1).Enabled = False
        tmrCurrTime.Enabled = True
        dpTime.Value = Time
    End If
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
    dpTime.Value = Time
End Sub


Private Sub tmrBlink_Timer()
    Labels(0).Visible = Not Labels(0).Visible
End Sub

Private Sub tmrCurrDate_Timer()
    If dtpDate.Value <> Date Then dtpDate.Value = Date
End Sub

Private Sub tmrCurrTime_Timer()
    dpTime.Value = Time
End Sub

Private Sub DisplayCap()
    If Check1.Value = 1 Or Check2.Value = 1 Then
        btnOK.Caption = "Setting"
    Else
        btnOK.Caption = "Tutup"
    End If
End Sub
