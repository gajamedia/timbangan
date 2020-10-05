VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmScatter 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lookup Data"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10170
   Icon            =   "FrmScatter.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   10170
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox TxtFind 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Top             =   3120
      Width           =   7095
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   10170
      _ExtentX        =   17939
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
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Search ID"
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
   Begin MSDataGridLib.DataGrid DataGridView 
      Align           =   1  'Align Top
      Bindings        =   "FrmScatter.frx":57E2
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Data"
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
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
End
Attribute VB_Name = "FrmScatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim nSort As Integer, Byt_ColSort As Integer

Private Sub DataGridView_DblClick()
    If Not Adodc2.Recordset.EOF Then
       On Error Resume Next
            Save_Code Adodc2.Recordset.Fields(0), _
                 Adodc2.Recordset.Fields(Field_No), _
                 Adodc2.Recordset.Fields(Field_No1), _
                 Adodc2.Recordset.Fields(Field_No2), _
                 Adodc2.Recordset.Fields(Field_No3), _
                 Adodc2.Recordset.Fields(Field_No4)
        Else
            Save_Code "", "", "", "", "", ""
    End If
    Unload Me
End Sub

Private Sub DataGridView_HeadClick(ByVal ColIndex As Integer)
    Dim str_SortOrder As String
 
    If Byt_ColSort = ColIndex + 1 Then
        str_SortOrder = " desc"
        Byt_ColSort = 0
    Else
        str_SortOrder = " asc"
        Byt_ColSort = ColIndex + 1
    End If
    Adodc2.Recordset.Sort = DataGridView.Columns(ColIndex).DataField & str_SortOrder
End Sub

Private Sub DataGridView_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 35
        Adodc2.Recordset.MoveLast
    Case 36
        Adodc2.Recordset.MoveFirst
    End Select
End Sub

Private Sub DataGridView_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13
        If Not Adodc2.Recordset.EOF Then
            On Error Resume Next
            Save_Code Adodc2.Recordset.Fields(0), _
            Adodc2.Recordset.Fields(Field_No), _
            Adodc2.Recordset.Fields(Field_No1), _
            Adodc2.Recordset.Fields(Field_No2), _
            Adodc2.Recordset.Fields(Field_No3), _
            Adodc2.Recordset.Fields(Field_No4)
        Else
            Save_Code "", "", "", "", "", ""
        End If
        Unload Me
    Case 27
        Unload Me
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 35 Then Adodc2.Recordset.MoveLast
    'DatagridColumnAutoResize DataGridView, FrmScatter
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Adodc2.ConnectionString = CnString
    Adodc2.RecordSource = RsString
    Adodc2.Refresh
    DataGridView.Refresh
    DataGridView.MarqueeStyle = 3
    Combo1.Clear
    For i = 0 To Adodc2.Recordset.Fields.Count - 1
        Combo1.AddItem Adodc2.Recordset.Fields(i).Name
        'If UCase(Adodc2.Recordset.Fields(i).Name) = "KODE" Then nScat = i
    Next
    Combo1.ListIndex = nScat
    nSort = 1
 
    DatagridColumnAutoResize DataGridView, FrmScatter
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    DataGridView.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 700
    TxtFind.Top = DataGridView.Height + 100
    Combo1.Top = DataGridView.Height + 100
    Label1.Top = DataGridView.Height + 100
    DatagridColumnAutoResize DataGridView, FrmScatter
End Sub

Private Sub TxtFind_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 And _
    TxtFind <> vbNullString Then
  If Not Adodc2.Recordset.EOF Then
   Adodc2.Recordset.Find Combo1.Text & " LIKE '%" & TxtFind.Text & "%'"
   If Adodc2.Recordset.EOF Then
    Adodc2.Refresh
   Else
    DataGridView.SetFocus
   End If
  End If
 End If
End Sub
