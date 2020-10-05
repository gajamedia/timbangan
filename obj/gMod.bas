Attribute VB_Name = "gMod"
Global GForm(18) As Form
Global cODBCName
Global iDxFrm As Integer
Global Scatter_Code As Variant, Scatter_Code1 As Variant
Global Scatter_Code2 As Variant, Scatter_Code3 As Variant
Global Scatter_Code4 As Variant, Scatter_Code5 As Variant
Global Field_No As Integer, Field_No1 As Integer
Global Field_No2 As Integer, Field_No3 As Integer
Global Field_No4 As Integer, noRPT As Integer
Global CnString As String, RsString As String
Global strSelectCritera As String, noBuk As String
Global tampungtiket As String, Logon As Boolean
Global jmltiket As Integer, nScat As Integer
Global UserID As String, UserPass As String, UserGroup As String
Global cMemvar As String, cMemKey As String, vLoadHelp As Boolean
Global Kas As String, Piutang As String, Hutang As String
Global PotPenj As String
Global appon As Boolean

'Initialisasi Port -----------------------------
 Global ncom As Integer, baudrate As Integer
 Global parity As String, databit As Integer
 Global stopbit As Integer, sPathData As String
'-----------------------------------------------

'Inisialisasi Printer ------------------------
 Private Type DOCINFO
  pDocName As String
  pOutputFile As String
  pDatatype As String
 End Type
 Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
 Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
 Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
 Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
 Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
 Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
 Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
 
 Global gBruto, gNetto, gTara
 Global gLambung, gNopol, gRFID
 Global gMasuk, gKeluar
 Global gBarang, gPemilik
 Global gNomer, gDermaga
'----------------------------------------------

Public Declare Function GetTickCount Lib "kernel32" () As Long

Option Explicit

Private Sub Cetak(Title As String, DataPrint As String, sizeFont As Long)
 Dim lhPrinter As Long
 Dim lReturn As Long
 Dim lpcWritten As Long
 Dim lDoc As Long
 Dim sWrittenData As String
 Dim MyDocInfo As DOCINFO

 lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
 If lReturn = 0 Then
  MsgBox "Tidak Ada Printer Yang Tersedia !!!", vbOKOnly + vbCritical, "JASATAMA"
  Exit Sub
 End If

 MyDocInfo.pDocName = Title
 MyDocInfo.pOutputFile = vbNullString
 MyDocInfo.pDatatype = "RAW"
 lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
 Call StartPagePrinter(lhPrinter)
 
 'sWrittenData = vbCrLf & DataPrint & vbCrLf
 sWrittenData = DataPrint & vbCrLf
 ' vbFormFeed
 Printer.FontSize = sizeFont
 lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, Len(sWrittenData), lpcWritten)

 lReturn = EndPagePrinter(lhPrinter)
 lReturn = EndDocPrinter(lhPrinter)
 lReturn = ClosePrinter(lhPrinter)
End Sub

Public Function Cetak_Struk2(cFile As String) As String
 Dim cData As String, fBaris As Long
 Dim i As Integer, nAw As Integer, nAk As Integer
 Dim cTemp As String, cField As String, sfont As Long
 
 cData = FileRead(cFile, False, fBaris)(1): cData = vbNullString
 For i = 1 To fBaris
  cTemp = FileRead(cFile, False)(i)
  sfont = 10

  'Membaca Apakah Ada Sebuah Field ? ------------------
   nAw = InStr(cTemp, "<<"): nAk = InStr(cTemp, ">>")
   If nAw <> 0 Then
    cField = Mid(cTemp, nAw + 2, nAk - (nAw + 2))
    If cField = "bruto" Then
     sfont = 14
     'cTemp = Mid(cTemp, 1, nAw - 1) & pvcur(1).Text & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
     cTemp = Mid(cTemp, 1, nAw - 1) & gBruto & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
    ElseIf cField = "tara" Then
     sfont = 14
     cTemp = Mid(cTemp, 1, nAw - 1) & gTara & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
     'If cMode <> "1" Then
     ' cTemp = Mid(cTemp, 1, nAw - 1) & pvcur(0).Text & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
     'Else
     ' cTemp = Mid(cTemp, 1, nAw - 1) & pvcur(2).Text & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
     'End If
    ElseIf cField = "netto" Then
     sfont = 14
     'cTemp = Mid(cTemp, 1, nAw - 1) & pvcur(3).Text & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
     cTemp = Mid(cTemp, 1, nAw - 1) & gNetto & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
    ElseIf cField = "nolambung" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & txtData(0).Text
     cTemp = Mid(cTemp, 1, nAw - 1) & gLambung
    ElseIf cField = "nopol" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & txtData(1).Text
     cTemp = Mid(cTemp, 1, nAw - 1) & gNopol
    ElseIf cField = "wmasuk" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & Format(dMasuk, "dd-mm-yyyy / HH:MM:SS")
     cTemp = Mid(cTemp, 1, nAw - 1) & Format(gMasuk, "dd-mm-yyyy / HH:MM:SS")
    ElseIf cField = "wkeluar" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & Format(dKeluar, "dd-mm-yyyy / HH:MM:SS")
     cTemp = Mid(cTemp, 1, nAw - 1) & Format(gKeluar, "dd-mm-yyyy / HH:MM:SS")
    ElseIf cField = "barang" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & txtData(3).Text
     cTemp = Mid(cTemp, 1, nAw - 1) & gBarang
    ElseIf cField = "pemilik" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & cPemilik
     cTemp = Mid(cTemp, 1, nAw - 1) & gPemilik
    ElseIf cField = "nomer" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & cNomer
     cTemp = Mid(cTemp, 1, nAw - 1) & gNomer
    ElseIf cField = "nodermaga" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & cnDmg
     cTemp = Mid(cTemp, 1, nAw - 1) & gDermaga
    ElseIf cField = "nmoperator" Then
     cTemp = Mid(cTemp, 1, nAw - 1) & UserID
    ElseIf cField = "norfid" Then
     cTemp = Mid(cTemp, 1, nAw - 1) & gRFID
    End If
   End If
  '----------------------------------------------------
  
  'Cetak "Cetak Struk", cTemp, sfont
  cData = cData & cTemp
  If i <> fBaris Then cData = cData & vbCrLf
 Next
  Cetak_Struk2 = cData
End Function

Public Sub Cetak_Struk(cFile As String)
 Dim cData As String, fBaris As Long
 Dim i As Integer, nAw As Integer, nAk As Integer
 Dim cTemp As String, cField As String, sfont As Long
 
 cData = FileRead(cFile, False, fBaris)(1): cData = vbNullString
 For i = 1 To fBaris
  cTemp = FileRead(cFile, False)(i)
  sfont = 10

  'Membaca Apakah Ada Sebuah Field ? ------------------
   nAw = InStr(cTemp, "<<"): nAk = InStr(cTemp, ">>")
   If nAw <> 0 Then
    cField = Mid(cTemp, nAw + 2, nAk - (nAw + 2))
    If cField = "bruto" Then
     sfont = 14
     'cTemp = Mid(cTemp, 1, nAw - 1) & pvcur(1).Text & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
     cTemp = Chr(27) & Chr(33) & Chr(2) & Mid(cTemp, 1, nAw - 1) & gBruto & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1)) & Chr(27) & Chr(33) & Chr(4)
    ElseIf cField = "tara" Then
     sfont = 14
     cTemp = Chr(27) & Chr(33) & Chr(2) & Mid(cTemp, 1, nAw - 1) & gTara & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1)) & Chr(27) & Chr(33) & Chr(4)
     'If cMode <> "1" Then
     ' cTemp = Mid(cTemp, 1, nAw - 1) & pvcur(0).Text & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
     'Else
     ' cTemp = Mid(cTemp, 1, nAw - 1) & pvcur(2).Text & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
     'End If
    ElseIf cField = "netto" Then
     sfont = 14
     'cTemp = Mid(cTemp, 1, nAw - 1) & pvcur(3).Text & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1))
     cTemp = Chr(27) & Chr(33) & Chr(2) & Mid(cTemp, 1, nAw - 1) & gNetto & Mid(cTemp, nAk + 2, Len(cTemp) - (nAk + 1)) & Chr(27) & Chr(33) & Chr(4)
    ElseIf cField = "nolambung" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & txtData(0).Text
     cTemp = Mid(cTemp, 1, nAw - 1) & gLambung
    ElseIf cField = "nopol" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & txtData(1).Text
     cTemp = Mid(cTemp, 1, nAw - 1) & gNopol
    ElseIf cField = "wmasuk" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & Format(dMasuk, "dd-mm-yyyy / HH:MM:SS")
     cTemp = Mid(cTemp, 1, nAw - 1) & Format(gMasuk, "dd-mm-yyyy / HH:MM:SS")
    ElseIf cField = "wkeluar" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & Format(dKeluar, "dd-mm-yyyy / HH:MM:SS")
     cTemp = Mid(cTemp, 1, nAw - 1) & Format(gKeluar, "dd-mm-yyyy / HH:MM:SS")
    ElseIf cField = "barang" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & txtData(3).Text
     cTemp = Mid(cTemp, 1, nAw - 1) & gBarang
    ElseIf cField = "pemilik" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & cPemilik
     cTemp = Mid(cTemp, 1, nAw - 1) & gPemilik
    ElseIf cField = "nomer" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & cNomer
     cTemp = Mid(cTemp, 1, nAw - 1) & gNomer
    ElseIf cField = "nodermaga" Then
     'cTemp = Mid(cTemp, 1, nAw - 1) & cnDmg
     cTemp = Mid(cTemp, 1, nAw - 1) & gDermaga
    ElseIf cField = "nmoperator" Then
     cTemp = Mid(cTemp, 1, nAw - 1) & UserID
    ElseIf cField = "norfid" Then
     cTemp = Mid(cTemp, 1, nAw - 1) & gRFID
    End If
   End If
  '----------------------------------------------------
  
  Cetak "Cetak Struk", cTemp, sfont
  'cData = cData & cTemp
  'If i <> fBaris Then cData = cData & vbCrLf
 Next
 'MsgBox cData
End Sub

Public Sub Save_Code(vString As Variant, vString1 As Variant, _
         vString2 As Variant, vString3 As Variant, _
         vString4 As Variant, vString5 As Variant)
    Scatter_Code = vString
    Scatter_Code1 = vString1
    Scatter_Code2 = vString2
    Scatter_Code3 = vString3
    Scatter_Code4 = vString4
    Scatter_Code5 = vString5
End Sub

Public Sub ShowFind(vCn As String, vRs As String, sCap As String, nKey As Integer, _
        Optional nField As Integer, Optional nField1 As Integer, _
        Optional nField2 As Integer, Optional nField3 As Integer, _
        Optional nField4 As Integer)
    CnString = vCn
    RsString = vRs
    nScat = nKey
    Field_No = nField
    Field_No1 = nField1
    Field_No2 = nField2
    Field_No3 = nField3
    Field_No4 = nField4
    
    FrmScatter.Caption = sCap
    FrmScatter.Show 1
End Sub

Public Sub DatagridColumnAutoResize(ByRef oDataGrid As DataGrid, _
    ByRef oForm As Form)
Dim i As Integer, iMax As Integer
Dim t As Integer, tMax As Integer
Dim iWidth As Integer
Dim vBMark As Variant
Dim aWidth As Variant
Dim cText As String
Dim oFont As Font

'    On Error Resume Next

    'need this to make TextWidth()
    'work with prossibly different font in DG
    Set oFont = oForm.Font
    oForm.Font = oDataGrid.Font

    iMax = oDataGrid.Columns.Count - 1
    ReDim aWidth(iMax)

    For i = 0 To iMax   'init maxwidth holder
       aWidth(i) = 0
    Next

    'one visible page to get to an estimate
    tMax = oDataGrid.VisibleRows - 1
    If tMax > 0 Then
        For t = 0 To tMax   'number of rows
            vBMark = oDataGrid.GetBookmark(t)
            For i = 0 To iMax   'number of columns
                cText = oDataGrid.Columns(i).CellText(vBMark)
                iWidth = oForm.TextWidth(cText)
                If iWidth + ((12 * Len(cText)) + 220) > aWidth(i) Then
                    'the font is right, the stringlength too, but
                    'still some misalignment on long stings. So we
                    'have to fiddle this a bit by hand...
                    aWidth(i) = iWidth + ((12 * Len(cText)) + 220)
                End If
                If t = 0 Then   'take care of the headers
                    iWidth = oForm.TextWidth(oDataGrid.Columns( _
                        i).Caption)
                    If iWidth + ((12 * Len(cText)) + 220) > aWidth( _
                        i) Then
                        aWidth(i) = iWidth + ((12 * Len(cText)) + 220)
                    End If
                End If
            Next
        Next
        For i = 0 To iMax   ' finally set the new column width
            oDataGrid.Columns(i).Width = aWidth(i)
        Next
    End If
    oForm.Font = oFont
End Sub

Public Sub Menu_Visible(bFlag As Boolean, bRights As Boolean)
 Dim i As Integer
 
 If Not bRights Then
  frmmain.mdiMenu(1).Visible = bFlag
  frmmain.mdiMenu(2).Visible = bFlag
  frmmain.mdiMenu(3).Visible = bFlag
 Else
  User_Rights
 End If
 
 If Logon Then
  frmmain.mnApp(0).Caption = "&Logout"
  frmmain.mnApp(5).Visible = True
 Else
  frmmain.mnApp(0).Caption = "&Login"
  For i = 2 To frmmain.mnApp.UBound - 1
   frmmain.mnApp(i).Visible = False
  Next
 End If
End Sub

Private Sub User_Rights()
    Dim cn As New ADODB.Connection
    Dim Menu1 As String, Menu2 As String
    Dim Menu3 As String, Menu4 As String
    Dim cStatus As String, i As Integer
    Dim rsMenu As New ADODB.Recordset
    Dim Menu5 As String
    
    Set cn = New ADODB.Connection
 
   'Open Database --------------------
    cn.CursorLocation = adUseClient
    cn.Open "DSN=dstimbang2"
   '----------------------------------
    
    Set rsMenu = cn.Execute("select * from tbgrup where cKode='" & UserGroup & "'")
    If Not rsMenu.EOF Then
     Menu1 = Trim(rsMenu.Fields("cMenu1"))
     Menu2 = Trim(rsMenu.Fields("cMenu2"))
     Menu3 = Trim(rsMenu.Fields("cMenu3"))
     Menu4 = Trim(rsMenu.Fields("cMenu4"))
     Menu5 = Trim(rsMenu.Fields("cMenu5"))

        cStatus = Mid(Menu1, 1, 1)
        If cStatus = "0" Then
            frmmain.mdiMenu(0).Visible = False
        Else
            frmmain.mdiMenu(0).Visible = True
            For i = 1 To Len(Menu1) - 1
                cStatus = Mid(Menu1, i + 1, 1)
                If cStatus = "0" Then
                    frmmain.mnApp(i + 1).Visible = False
                Else
                    frmmain.mnApp(i + 1).Visible = True
                End If
            Next i
        End If

        cStatus = Mid(Menu2, 1, 1)
        If cStatus = "0" Or cStatus = vbNullString Then
            frmmain.mdiMenu(1).Visible = False
        Else
            frmmain.mdiMenu(1).Visible = True
            For i = 1 To Len(Menu2) - 1
                cStatus = Mid(Menu2, i + 1, 1)
                If cStatus = "0" Then
                    frmmain.mnMaster(i - 1).Visible = False
                Else
                    frmmain.mnMaster(i - 1).Visible = True
                End If
            Next i
        End If
        
        cStatus = Mid(Menu3, 1, 1)
        If cStatus = "0" Or cStatus = vbNullString Then
            frmmain.mdiMenu(2).Visible = False
        Else
            frmmain.mdiMenu(2).Visible = True
            For i = 1 To Len(Menu3) - 1
                cStatus = Mid(Menu3, i + 1, 1)
                If cStatus = "0" Then
                    frmmain.mnTrans(i - 1).Visible = False
                Else
                    frmmain.mnTrans(i - 1).Visible = True
                End If
            Next i
        End If
        
        cStatus = Mid(Menu4, 1, 1)
        If cStatus = "0" Or cStatus = vbNullString Then
            frmmain.mdiMenu(3).Visible = False
        Else
            frmmain.mdiMenu(3).Visible = True
            For i = 1 To Len(Menu4) - 1
                cStatus = Mid(Menu4, i + 1, 1)
                If cStatus = "0" Then
                    frmmain.mnLap(i - 1).Visible = False
                Else
                    frmmain.mnLap(i - 1).Visible = True
                End If
            Next i
        End If
    
        cStatus = Mid(Menu5, 1, 1)
        If cStatus = "0" Or cStatus = vbNullString Then
            frmmain.mdiMenu(4).Visible = False
        Else
            frmmain.mdiMenu(4).Visible = True
            For i = 1 To Len(Menu5) - 1
                cStatus = Mid(Menu5, i + 1, 1)
                If cStatus = "0" Then
                    frmmain.mnUtil(i - 1).Visible = False
                Else
                    frmmain.mnUtil(i - 1).Visible = True
                End If
            Next i
        End If
    
    End If
    Set rsMenu = Nothing
  cn.Close
End Sub

Public Sub FileWriteBinary(vData As Variant, sFileName As String, Optional bAppendToFile As Boolean = True)
    Dim iFileNum As Integer, lWritePos As Long
    
    On Error GoTo ErrFailed
    If bAppendToFile = False Then
        If Len(Dir$(sFileName)) > 0 And Len(sFileName) > 0 Then
            'Delete the existing file
            VBA.Kill sFileName
        End If
    End If
    
    iFileNum = FreeFile
    Open sFileName For Binary Access Write As #iFileNum
    
    If bAppendToFile = False Then
        'Write to first byte
        lWritePos = 1
    Else
        'Write to last byte + 1
        lWritePos = LOF(iFileNum) + 1
    End If
    
    Put #iFileNum, lWritePos, vData
    Close iFileNum
    
    'FileWriteBinary = True
    Exit Sub

ErrFailed:
    'FileWriteBinary = False
    Close iFileNum
    Debug.Print Err.Description
End Sub

Public Function FileRead(sFileName As String, BinaryFile As Boolean, Optional ByRef nLine As Long) As Variant
    Dim iFileNum As Integer, lFileLen As Long
    Dim vThisBlock As Variant, lThisBlock As Long, vFileData As Variant
    
    On Error GoTo ErrFailed
    
    If Len(Dir$(sFileName)) > 0 And Len(sFileName) > 0 Then
        iFileNum = FreeFile
        If BinaryFile Then
         Open sFileName For Binary Access Read As #iFileNum
        Else
         Open sFileName For Input As #iFileNum
        End If
        
        lFileLen = LOF(iFileNum)
        
        Do
            lThisBlock = lThisBlock + 1
            If BinaryFile Then
             Get #iFileNum, , vThisBlock
            Else
             Line Input #iFileNum, vThisBlock
            End If
            If IsEmpty(vThisBlock) = False Then
                If lThisBlock = 1 Then
                    ReDim vFileData(1 To 1)
                Else
                    ReDim Preserve vFileData(1 To lThisBlock)
                End If
                vFileData(lThisBlock) = vThisBlock
            End If
        Loop While EOF(iFileNum) = False
        Close iFileNum
        
        FileRead = vFileData
        nLine = lThisBlock
    End If

    Exit Function
    
ErrFailed:
    Close iFileNum
    Debug.Print Err.Description
End Function

Public Function FileExists(Fname As String) As Boolean

 If Fname = "" Or Right(Fname, 1) = "\" Then
  FileExists = False: Exit Function
 End If

 FileExists = (Dir(Fname) <> "")

End Function

Public Function NumberToRomawi(ByVal cNumber As Long) As String
    Dim n As Long
    Dim IntIdy As Integer, intIdx As Integer
    Dim Rom, Latin
    Dim Romawi As String

    Rom = Array("I", "IV", "V", "IX", "X", "XL", "L", "XC", "C", "CD", "D", "CM", "M")
    Latin = Array(1, 4, 5, 9, 10, 40, 50, 90, 100, 400, 500, 900, 1000)

    Romawi = ""

    For IntIdy = 12 To 0 Step -1
        n = Int(cNumber / Latin(IntIdy))
        If n <> 0 Then
            For intIdx = 1 To n
                Romawi = Romawi & Rom(IntIdy)
            Next intIdx
        End If
        cNumber = cNumber Mod Latin(IntIdy)
    Next
    NumberToRomawi = Romawi
End Function

Public Function RomawiToNumber(cRomawi As String) As Long
    Dim n As Long
    Dim intIdx As Integer
    Dim Rom, Latin

    Rom = Array("I", "1", "V", "2", "X", "3", "L", "4", "C", "5", "D", "6", "M")
    Latin = Array(1, 4, 5, 9, 10, 40, 50, 90, 100, 400, 500, 900, 1000)

    cRomawi = Replace(cRomawi, "IV", "1")
    cRomawi = Replace(cRomawi, "IX", "2")
    cRomawi = Replace(cRomawi, "XL", "3")
    cRomawi = Replace(cRomawi, "XC", "4")
    cRomawi = Replace(cRomawi, "CD", "5")
    cRomawi = Replace(cRomawi, "CM", "6")

    n = 0
    For intIdx = 12 To 0 Step -1
        While Rom(intIdx) = Mid(cRomawi, 1, 1) And Len(cRomawi) > 0
            n = n + Latin(intIdx)
            cRomawi = Mid(cRomawi, 2)
        Wend
    Next intIdx
    RomawiToNumber = n
End Function
