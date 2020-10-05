VERSION 5.00
Begin VB.Form frmsetport 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Port Timbangan"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "frmsetport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   0
      ItemData        =   "frmsetport.frx":57E2
      Left            =   4080
      List            =   "frmsetport.frx":57F2
      TabIndex        =   13
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "Set Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2655
      Begin VB.Image Image1 
         Height          =   2400
         Left            =   120
         Picture         =   "frmsetport.frx":5806
         Stretch         =   -1  'True
         Top             =   840
         Width           =   2400
      End
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   4
      ItemData        =   "frmsetport.frx":60B1
      Left            =   3135
      List            =   "frmsetport.frx":60BB
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   3
      ItemData        =   "frmsetport.frx":60C5
      Left            =   3135
      List            =   "frmsetport.frx":60D8
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   2
      ItemData        =   "frmsetport.frx":60F0
      Left            =   3135
      List            =   "frmsetport.frx":60FD
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      Index           =   1
      ItemData        =   "frmsetport.frx":6110
      Left            =   3135
      List            =   "frmsetport.frx":6132
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   2640
      Top             =   3375
      Width           =   3855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "COM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   525
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      X1              =   2640
      X2              =   6480
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stop Bit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4695
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Data Bit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4695
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Parity"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4695
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Baud Rate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4695
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Setting Data Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pilih Koneksi Port"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2655
      TabIndex        =   5
      Top             =   0
      Width           =   3840
   End
End
Attribute VB_Name = "frmsetport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdsimpan_Click()
 If FileExists(sPathData) Then
  FileWriteBinary Trim(cmbData(0).Text), App.Path & "\data.dat", False
 Else
  FileWriteBinary Trim(cmbData(0).Text), App.Path & "\data.dat"
 End If
 FileWriteBinary cmbData(1).Text, App.Path & "\data.dat", True
 FileWriteBinary cmbData(2).ListIndex, App.Path & "\data.dat", True
 FileWriteBinary Trim(cmbData(3).Text), App.Path & "\data.dat", True
 FileWriteBinary cmbData(4).Text, App.Path & "\data.dat", True
 
 MsgBox "Data Port Telah Disimpan", vbOKOnly + vbInformation, "SIT30 - JASATAMA"
End Sub

Private Sub Form_Load()
 Dim i As Byte
 
 If FileExists(sPathData) Then
  cmbData(0).Text = FileRead(sPathData, True)(1)
  cmbData(1).Text = FileRead(sPathData, True)(2)
  cmbData(2).ListIndex = FileRead(sPathData, True)(3)
  cmbData(3).Text = FileRead(sPathData, True)(4)
  cmbData(4).Text = FileRead(sPathData, True)(5)
 Else
  For i = 0 To 4
   cmbData(i).ListIndex = 0
  Next
 End If
End Sub
