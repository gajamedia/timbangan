VERSION 5.00
Begin VB.Form frmlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SILAHKAN ANDA LOGIN"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6240
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6240
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdlogin 
      Caption         =   "Login"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2520
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2295
      Begin VB.Image Image1 
         Height          =   1710
         Left            =   480
         Picture         =   "frmlogin.frx":57E2
         Top             =   600
         Width           =   1305
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "PT. GRESIK JASATAMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2550
      TabIndex        =   7
      Top             =   240
      Width           =   2115
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "2004 - 2012"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   2880
      Width           =   3795
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Sistem Informasi Timbangan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
      Width           =   3795
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      X1              =   2280
      X2              =   6240
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      X1              =   2280
      X2              =   6240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password :"
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
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama User :"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   1155
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   5040
      Picture         =   "frmlogin.frx":70F6
      Stretch         =   -1  'True
      Top             =   60
      Width           =   855
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rsLogin As New ADODB.Recordset
Dim strSql As String
Dim fEncrypt As New Encrypt

Option Explicit

'Prosedur Cek Nama User Dan Password User --------------------------------------------------------------
Private Sub cmdlogin_Click()
 Dim sPass As String
 
 If Trim(txtUserName.Text) <> vbNullString Then
  strSql = "select * from tbUser " & _
           "where user_name='" & Trim(txtUserName.Text) & "'"
           
  Set rsLogin = con.Execute(strSql)
  If Not rsLogin.EOF Then
    sPass = fEncrypt.ChgPass(rsLogin.Fields("user_pass").Value, 2)
    UserID = rsLogin.Fields("user_name").Value
    UserPass = Trim(sPass)
    UserGroup = rsLogin.Fields("user_grup").Value
  End If
  Set rsLogin = Nothing
         
  If (Trim(txtUserName.Text) = UserID) Then
   If (Trim(txtPassword.Text) = UserPass) Then
    frmmain.sb.Panels(1).Text = "Operator : " & UserID & " (" & UserGroup & ")"
    
    Logon = True
    Menu_Visible True, True
    Unload Me
   Else
    MsgBox "Password User Salah !", vbOKOnly + vbInformation, "Login - JASATAMA"
    txtPassword.SetFocus
   End If
  Else
   MsgBox "Nama User Tidak Terdaftar di Sistem !", vbOKOnly + vbInformation, "Login - JASATAMA"
   txtUserName.SetFocus
  End If
 Else
   MsgBox "Nama User Boleh Kosong", vbOKOnly + vbInformation, "Login - JASATAMA"
   txtUserName.SetFocus
 End If
End Sub
'---------------------------------------------------------------------------------------------------------

Private Sub Form_Load()
 'Open Database --------------------
  con.CursorLocation = adUseClient
  con.ConnectionTimeout = 0
  con.Open "DSN=dstimbang2"
 '----------------------------------

 Label4.Caption = Chr(169) & " " & Label4.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
 con.Close
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then cmdlogin_Click
End Sub
