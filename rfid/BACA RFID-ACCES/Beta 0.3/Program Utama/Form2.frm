VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ".: DATABASE AKTIVITAS PENGISIAN BBM :."
   ClientHeight    =   10575
   ClientLeft      =   2910
   ClientTop       =   585
   ClientWidth     =   15225
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   15225
   Begin VB.CommandButton PRINT 
      Caption         =   "PRINT"
      Height          =   375
      Left            =   10560
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Timer TimerX 
      Interval        =   500
      Left            =   13200
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Left            =   12000
      Top             =   240
   End
   Begin VB.CommandButton COM3 
      Caption         =   "DISPLAY"
      Height          =   375
      Left            =   10560
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   11520
      Top             =   240
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11040
      Top             =   240
   End
   Begin VB.TextBox LED4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   12000
      TabIndex        =   13
      Text            =   " 0L          0L"
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox LED3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   450
      Left            =   12000
      TabIndex        =   12
      Text            =   "PENGISIAN  TOTAL"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox LED2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   12000
      TabIndex        =   11
      Text            =   " **          0L"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox LED1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   12000
      TabIndex        =   10
      Text            =   "NO_MOBIL   QUOTA"
      Top             =   840
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form2.frx":0000
      Height          =   2415
      Left            =   360
      TabIndex        =   8
      Top             =   480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4260
      _Version        =   393216
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
      Caption         =   "QUERY RFID"
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
            LCID            =   1057
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
            LCID            =   1057
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   360
      Top             =   3120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAMA\Desktop\DATABASE.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAMA\Desktop\DATABASE.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT RFID,NO_POLISI,QUOTA FROM TABEL_RFID ORDER BY RFID"
      Caption         =   "QUERY - CEK DATA RFID"
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
   Begin VB.CommandButton COM2 
      Caption         =   "FLOW"
      Height          =   375
      Left            =   10560
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   1440
      Width           =   3495
   End
   Begin VB.CommandButton COM1 
      Caption         =   "RFID"
      Height          =   375
      Left            =   10560
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   840
      Width           =   3495
   End
   Begin VB.Timer Timer0 
      Interval        =   1000
      Left            =   10560
      Top             =   240
   End
   Begin VB.CommandButton EDIT 
      Caption         =   "INPUT DATA"
      Height          =   375
      Left            =   10320
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   360
      Top             =   9960
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAMA\Desktop\DATABASE.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAMA\Desktop\DATABASE.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"Form2.frx":0015
      Caption         =   "DATABASE AKTIVITAS PENGISIAN BBM"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":00B4
      Height          =   6135
      Left            =   360
      TabIndex        =   0
      Top             =   3600
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   10821
      _Version        =   393216
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
      Caption         =   "TABEL_BBM"
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
            LCID            =   1057
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
            LCID            =   1057
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
   Begin MSCommLib.MSComm MSComm2 
      Left            =   9120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   6
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm3 
      Left            =   9840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   7
      DTREnable       =   -1  'True
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   6480
      Top             =   3120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAMA\Desktop\DATABASE.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAMA\Desktop\DATABASE.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT RFID FROM TABEL_BBM"
      Caption         =   "QUERY - CEK DATA RFID"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   12000
      Top             =   3120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAMA\Desktop\DATABASE.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAMA\Desktop\DATABASE.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TABEL_BBMX"
      Caption         =   "QUERY - PRINTER"
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
   Begin VB.Label Label5 
      Caption         =   "INPUT SERIAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label KET2 
      Caption         =   "KEMUDIAN LAKUKAN PENGISIAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label KET1 
      Caption         =   "MASUKKAN DATA RFID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13800
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DATA_RFID As String
Dim TERIMA_RFID As String
Dim TERIMA_FLOW As String

Dim ADA_RFID As Boolean
Dim PERTAMAX As Boolean
Dim KEDIP As Boolean

Dim ISI_TEMP As Single
Dim TOTAL_TEMP As Single
Dim SISA_TEMP As Single

Dim L_QUOTA As Integer
Dim L_ISI_TEMP As Integer
Dim L_TOTAL_TEMP As Integer


Private Sub COM1_Click()
    
    ADA_RFID = True
    DATA_RFID = Text1.Text
    Adodc3.RecordSource = "SELECT RFID,NO_POLISI,QUOTA FROM TABEL_RFID where RFID like'" & "%" & DATA_RFID & "%" & "'"
    Adodc3.Refresh

    If Adodc3.Recordset.EOF Then
        Text1.Text = ""
        ADA_RFID = False
        KET1.Caption = "DATA RFID DITOLAK"
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> SIMULASI LED2
        
        LED2.Text = "* RFID DITOLAK *"
        LED4.Text = " 0L          0L "
        Form3.LED2.Text = LED2.Text
        Form3.LED4.Text = LED4.Text
        
    Else
        Text1.Text = ""
        KET1.Caption = "DATA RFID DITERIMA"
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> SIMULASI LED2
        
        
        L_QUOTA = Len(Adodc3.Recordset.Fields("QUOTA"))
        
        If L_QUOTA = 3 Then
            LED2.Text = Adodc3.Recordset.Fields("NO_POLISI") + " " + Adodc3.Recordset.Fields("QUOTA") + "L"
        ElseIf L_QUOTA = 2 Then
            LED2.Text = Adodc3.Recordset.Fields("NO_POLISI") + "  " + Adodc3.Recordset.Fields("QUOTA") + "L"
        ElseIf L_QUOTA = 1 Then
            LED2.Text = Adodc3.Recordset.Fields("NO_POLISI") + "   " + Adodc3.Recordset.Fields("QUOTA") + "L"
        Else
            LED2.Text = Adodc3.Recordset.Fields("NO_POLISI") + " " + Adodc3.Recordset.Fields("QUOTA") + "L"
        End If
                
                
        LED4.Text = " 0L          0L "
        Form3.LED2.Text = LED2.Text
        Form3.LED4.Text = LED4.Text
        
    End If
    
    Adodc4.RecordSource = "SELECT RFID FROM TABEL_BBM where RFID like'" & "%" & DATA_RFID & "%" & "' AND TANGGAL like'" & "%" & Label1.Caption & "%" & "'"
    Adodc4.Refresh
    
    If Adodc4.Recordset.EOF Then
        PERTAMAX = True
        KET2.Caption = "PERTAMA KALI PENGISIAN"
    Else
        PERTAMAX = False
        KET2.Caption = "LEBIH DARI 1X PENGISIAN"
    End If

    
    
End Sub


Private Sub COM2_Click()

    If ADA_RFID = True Then
        Adodc2.Refresh
        With Adodc2.Recordset
        
            .AddNew
            .Fields("RFID") = DATA_RFID
            .Fields("NO_POLISI") = Adodc3.Recordset.Fields("NO_POLISI")
            .Fields("QUOTA") = Adodc3.Recordset.Fields("QUOTA")
            
            ISI_TEMP = Val(Text2.Text) / 1000
            STR_ISI_TEMP = Str(ISI_TEMP)
            
            .Fields("ISI_BBM") = STR_ISI_TEMP
            
            '''' bug
            
            If PERTAMAX = True Then
                TOTAL_TEMP = ISI_TEMP
                SISA_TEMP = Val(Adodc3.Recordset.Fields("QUOTA")) - ISI_TEMP
            ElseIf PERTAMAX = False Then
                TOTAL_TEMP = TOTAL_TEMP + ISI_TEMP
                SISA_TEMP = SISA_TEMP - ISI_TEMP
            End If
                
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> SIMULASI LED4
            
            L_ISI_TEMP = Len(STR_ISI_TEMP)
        
            If L_ISI_TEMP = 5 Then
                LED4.Text = Str(ISI_TEMP) + "L   " + Str(TOTAL_TEMP) + "L"
            ElseIf L_ISI_TEMP = 4 Then
                LED4.Text = Str(ISI_TEMP) + "L     " + Str(TOTAL_TEMP) + "L"
            ElseIf L_ISI_TEMP = 3 Then
                LED4.Text = Str(ISI_TEMP) + "L       " + Str(TOTAL_TEMP) + "L"
            ElseIf L_ISI_TEMP = 2 Then
                LED4.Text = Str(ISI_TEMP) + "L         " + Str(TOTAL_TEMP) + "L"
            ElseIf L_ISI_TEMP = 1 Then
                LED4.Text = Str(ISI_TEMP) + "L           " + Str(TOTAL_TEMP) + "L"
            Else
                LED4.Text = Str(ISI_TEMP) + "L " + Str(TOTAL_TEMP) + "L"
            End If
            
            Form3.LED4.Text = LED4.Text
                            
            STR_TOTAL_TEMP = Str(TOTAL_TEMP)
            .Fields("TOTAL_BBM") = STR_TOTAL_TEMP
            
            STR_SISA_TEMP = Str(SISA_TEMP)
            .Fields("SISA_BBM") = STR_SISA_TEMP
            
            .Fields("TANGGAL") = Label1.Caption
            .Fields("JAM") = Label2.Caption
            
            .Update
            .Sort = "JAM"
            DATA_RFID = ""
            ADA_RFID = False

        End With
        Adodc2.Refresh
        Adodc2.Refresh
        
        KET1.Caption = "PENGISIAN BBM DENGAN RFID"
        KET2.Caption = "DATA TERSIMPAN DI DATABASE"
        
        
    Else
        Adodc2.Refresh
        With Adodc2.Recordset
            .AddNew
            ISI_TEMP = Val(Text2.Text) / 1000
    
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> SIMULASI LED2
        
            LED2.Text = " * TANPA RFID * "
            LED4.Text = " 0L          0L "
            Form3.LED2.Text = LED2.Text
            Form3.LED4.Text = LED4.Text
            
            
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> SIMULASI LED4
            
            LED4.Text = Str(ISI_TEMP) + "L"
            Form3.LED4.Text = LED4.Text
            
            .Fields("RFID") = " * TANPA RFID * "
            .Fields("NO_POLISI") = "0"
            .Fields("QUOTA") = "0"
            
            ISI_TEMP = Val(Text2.Text) / 1000
            STR_ISI_TEMP = Str(ISI_TEMP)
            
            .Fields("ISI_BBM") = STR_ISI_TEMP
            
            .Fields("TOTAL_BBM") = "0"
            .Fields("SISA_BBM") = "0"
            .Fields("TANGGAL") = Label1.Caption
            .Fields("JAM") = Label2.Caption
            
            .Update
            .Sort = "JAM"
            DATA_RFID = ""
        End With
        Adodc2.Refresh
        Adodc2.Refresh
        
        KET1.Caption = "PENGISIAN BBM TANPA RFID"
        KET2.Caption = "DATA TERSIMPAN DI DATABASE"
        
        
        
    End If
            
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' TINDIH DATA PRINT
            
    Adodc5.Refresh
    Adodc5.Recordset.Fields("RFID") = Adodc2.Recordset.Fields("RFID")
    Adodc5.Recordset.Fields("NO_POLISI") = Adodc2.Recordset.Fields("NO_POLISI")
    Adodc5.Recordset.Fields("QUOTA") = Adodc2.Recordset.Fields("QUOTA")
    Adodc5.Recordset.Fields("ISI_BBM") = Adodc2.Recordset.Fields("ISI_BBM")
    Adodc5.Recordset.Fields("TOTAL_BBM") = Adodc2.Recordset.Fields("TOTAL_BBM")
    Adodc5.Recordset.Fields("SISA_BBM") = Adodc2.Recordset.Fields("SISA_BBM")
    Adodc5.Recordset.Fields("TANGGAL") = Adodc2.Recordset.Fields("TANGGAL")
    Adodc5.Recordset.Fields("JAM") = Adodc2.Recordset.Fields("JAM")
    Adodc5.Recordset.Update
    Adodc5.Refresh
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' NGEPRINT
    
    'DataReport1.Show
    'DataReport1.PrintReport False
    
    Text2.Text = ""
    Text1.BackColor = &H80000005
    Text1.Enabled = True
    
    Adodc2.RecordSource = "SELECT TANGGAL,JAM,RFID,NO_POLISI,QUOTA,ISI_BBM,TOTAL_BBM,SISA_BBM FROM TABEL_BBM where TANGGAL like'" & "%" & Label1.Caption & "%" & "' ORDER BY JAM DESC"
    Adodc2.Refresh

End Sub


Private Sub COM3_Click()
    Form3.Visible = True
End Sub

Private Sub EDIT_Click()
        
    Timer1.Enabled = False
        
    'Form1.Visible = True
    'Form1.Enabled = True
    Form2.Visible = False
    'Form2.Enabled = False
    
End Sub


Private Sub Form_Load()
    
    MSComm1.PortOpen = True              '''''''''''''''''' COM 5
    MSComm2.PortOpen = True              '''''''''''''''''' COM 6
    'MSComm3.PortOpen = True             '''''''''''''''''' COM ???
    
    Timer1.Enabled = True
    Timer2.Enabled = True
    
    KEDIP = True
    
    Dim xDate, xTime As Date
    xDate = DateValue(Now)
    xTime = TimeValue(Now)
    Label1.Caption = xDate
    Label2.Caption = xTime
    
    PERTAMAX = True
    
    'Adodc2.ConnectionString = "C:\Users\PAMA\Desktop\DATABASE.mdb"
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "SELECT TANGGAL,JAM,RFID,NO_POLISI,QUOTA,ISI_BBM,TOTAL_BBM,SISA_BBM FROM TABEL_BBM where TANGGAL like'" & "%" & Label1.Caption & "%" & "' ORDER BY JAM DESC"
    Adodc2.Refresh
    
    'Adodc3.ConnectionString = "C:\Users\PAMA\Desktop\DATABASE.mdb"
    'Adodc3.CommandType = adCmdText
    'Adodc3.RecordSource = "SELECT RFID,NO_POLISI,QUOTA FROM TABEL_RFID ORDER BY RFID"
    Adodc3.Refresh
    
    'Adodc4.ConnectionString = "C:\Users\PAMA\Desktop\DATABASE.mdb"
    'Adodc4.CommandType = adCmdText
    'Adodc4.RecordSource = "SELECT RFID FROM TABEL_BBM"
    Adodc4.Refresh
    
End Sub


Private Sub PRINT_Click()
    DataReport1.Show
    DataReport1.PrintReport True
End Sub

Private Sub Timer0_Timer()
  
    xDate = DateValue(Now)
    xTime = TimeValue(Now)
    
    Label1.Caption = xDate
    Label2.Caption = xTime
    
    If Label2.Caption = "6:00:00" Then
        PERTAMAX = True
        
        KET1.Caption = "SHIFT PAGI"
        KET2.Caption = "TOTAL & SISA = NOL"
        
        TOTAL_TEMP = 0
        SISA_TEMP = 0
        
    ElseIf Label2.Caption = "18:00:00" Then
        PERTAMAX = True
        
        KET1.Caption = "SHIFT MALAM"
        KET2.Caption = "TOTAL & SISA = NOL"
        
        TOTAL_TEMP = 0
        SISA_TEMP = 0
    End If
    
End Sub


Private Sub Timer1_Timer()
    
    TERIMA_RFID = MSComm1.Input                 '''''''''''''''''' COM
    TERIMA_RFID = Left(TERIMA_RFID, 9)
    
    If TERIMA_RFID = "" Then
        'Text1.Text = ""
    Else
        Text1.Text = TERIMA_RFID
        Call COM1_Click
    End If
    
End Sub

Private Sub Timer2_Timer()

    TERIMA_FLOW = MSComm2.Input                 '''''''''''''''''' COM
    'TERIMA_FLOW = Mid(TERIMA_FLOW, 8, 2)
    
    If TERIMA_FLOW = "" Then
        'Text1.Text = ""
    Else
        Text2.Text = TERIMA_FLOW
        Call COM2_Click
    End If

End Sub

Private Sub TimerX_Timer()
    
    If KEDIP = True Then
        Form3.LED2.Visible = True
        Form3.LED4.Visible = True
        KEDIP = False
    ElseIf KEDIP = False Then
        Form3.LED2.Visible = False
        Form3.LED4.Visible = False
        KEDIP = True
    End If
    
End Sub
