VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   ".: RFID MONITORING SYSTEM DISPLAY "
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form3"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Begin VB.TextBox JUDUL 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1260
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "RFID Monitoring System"
      Top             =   480
      Width           =   13575
   End
   Begin VB.TextBox LED1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1860
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "NO_MOBIL   QUOTA"
      Top             =   2280
      Width           =   14175
   End
   Begin VB.TextBox LED2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1890
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   " **          0L"
      Top             =   4200
      Width           =   14175
   End
   Begin VB.TextBox LED3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1890
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "PENGISIAN  TOTAL"
      Top             =   6360
      Width           =   14175
   End
   Begin VB.TextBox LED4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1830
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   " 0L          0L"
      Top             =   8280
      Width           =   14175
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   2280
      X2              =   17040
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   17040
      X2              =   17040
      Y1              =   2160
      Y2              =   10200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   2280
      X2              =   2280
      Y1              =   2160
      Y2              =   10200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   2280
      X2              =   17040
      Y1              =   10200
      Y2              =   10200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   2280
      X2              =   17040
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

