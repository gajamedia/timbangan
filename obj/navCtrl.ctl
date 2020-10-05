VERSION 5.00
Object = "{B8CF4BAA-3FDB-4253-9313-3861F6D8D086}#1.0#0"; "senxpctl.ocx"
Begin VB.UserControl navCtrl 
   BackStyle       =   0  'Transparent
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   ScaleHeight     =   765
   ScaleWidth      =   6045
   Begin OsenXPCntrl.OsenXPButton btnCtrl 
      Height          =   735
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Add"
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
      MICON           =   "navCtrl.ctx":0000
      PICN            =   "navCtrl.ctx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton btnCtrl 
      Height          =   735
      Index           =   5
      Left            =   3000
      TabIndex        =   5
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Edit"
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
      MICON           =   "navCtrl.ctx":0A2E
      PICN            =   "navCtrl.ctx":0A4A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton btnCtrl 
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "First"
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
      MICON           =   "navCtrl.ctx":145C
      PICN            =   "navCtrl.ctx":1478
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton btnCtrl 
      Height          =   735
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Prev"
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
      MICON           =   "navCtrl.ctx":1E8A
      PICN            =   "navCtrl.ctx":1EA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton btnCtrl 
      Height          =   735
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Next"
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
      MICON           =   "navCtrl.ctx":28B8
      PICN            =   "navCtrl.ctx":28D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton btnCtrl 
      Height          =   735
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Last"
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
      MICON           =   "navCtrl.ctx":32E6
      PICN            =   "navCtrl.ctx":3302
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton btnCtrl 
      Height          =   735
      Index           =   6
      Left            =   3600
      TabIndex        =   6
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Del"
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
      MICON           =   "navCtrl.ctx":3D14
      PICN            =   "navCtrl.ctx":3D30
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton btnCtrl 
      Height          =   735
      Index           =   7
      Left            =   4200
      TabIndex        =   7
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Brow"
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
      MICON           =   "navCtrl.ctx":4742
      PICN            =   "navCtrl.ctx":475E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton btnCtrl 
      Height          =   735
      Index           =   8
      Left            =   4800
      TabIndex        =   8
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Print"
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
      MICON           =   "navCtrl.ctx":5170
      PICN            =   "navCtrl.ctx":518C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton btnCtrl 
      Height          =   735
      Index           =   9
      Left            =   5400
      TabIndex        =   9
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Close"
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
      MICON           =   "navCtrl.ctx":5B9E
      PICN            =   "navCtrl.ctx":5BBA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton btnCtrl 
      Height          =   735
      Index           =   10
      Left            =   2400
      TabIndex        =   10
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Save"
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
      MICON           =   "navCtrl.ctx":65CC
      PICN            =   "navCtrl.ctx":65E8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton btnCtrl 
      Height          =   735
      Index           =   11
      Left            =   3000
      TabIndex        =   11
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Batal"
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
      MICON           =   "navCtrl.ctx":6FFA
      PICN            =   "navCtrl.ctx":7016
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "navCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub btnCtrl_Click(Index As Integer)
 Call GForm(iDxFrm).PControl(Index)
 
 Select Case Index
 Case 4
  EditPos
 Case 11
  ViewPos
 End Select
End Sub

Public Sub EditPos()
 Dim i As Integer
 
 For i = 0 To 9
  btnCtrl(i).Visible = False
 Next
 For i = 10 To 11
  btnCtrl(i).Visible = True
 Next
End Sub

Public Sub ViewPos()
 Dim i As Integer
 
 For i = 0 To 9
  btnCtrl(i).Visible = True
 Next
 For i = 10 To 11
  btnCtrl(i).Visible = False
 Next
End Sub

Private Sub btnCtrl_GotFocus(Index As Integer)
 btnCtrl(Index).ColorScheme = [Force Standard]
End Sub

Private Sub btnCtrl_LostFocus(Index As Integer)
 btnCtrl(Index).ColorScheme = [Use Windows]
End Sub
