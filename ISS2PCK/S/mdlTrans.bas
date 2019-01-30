Attribute VB_Name = "mdlTrans"
'=============================ISS2PCK AOG(QT,SO) ExVat ======================================
'conERP  =ExcludeVAT ,path\PCK61
'conERP2 =ExcludeVAT ,path\WAT61

'24/10/2018 baseline version
'29/10/2018 change update express.quotation_no to iss.DocNumber
'07/11/2018 combine aog_quote + aog_so
'29/01/2019 change code for crq 199 and crq 229-1

Option Explicit
 
Global conERP As New ADODB.Connection, conERP2 As New ADODB.Connection
Global conISS As ADODB.Connection, conISSA As ADODB.Connection
Global rsISS As ADODB.Recordset, rsISSA As ADODB.Recordset
Global rsISSH As ADODB.Recordset, rsISSI As ADODB.Recordset, rsISSD As ADODB.Recordset, rsISSC As ADODB.Recordset
Global rsERP As ADODB.Recordset
Global rsERPH As ADODB.Recordset, rsERPHR As ADODB.Recordset, rsERPD As ADODB.Recordset
Global strEmailTo, strEmailCC, strEmailAdmin1, strEmailAdmin2, strEmailSubject As String
Global strEmailOrder, strEmailQuata, strEmailReturn, strEmailStock As String
Global strISSServer, strISSUser, strISSPWD, strISSDBS, strISSDBA As String
Global strERPServer, strERPUser, strERPPWD, strERPPath, strERP2Path, strERP1UPath As String
Global strsql, strWorkingDir, strISSCompany As String
Global fn, intCurrMonth, intCurrYear, intNoOfRow, intpaytrm As Integer
Global strComcode, strCompany, strRuntype, strThaidoc, strEngdoc  As String
Global strSOno, strDoctype, strPrefix, strSlmlist, strSOlist As String
Global gVat_rate As Currency

        
Sub Main()
On Error GoTo Err_Handler

    Dim strISSConnectionString, strERPConnectionString As String
'    Dim dteDate, thaidate As Date
    Dim FSO As New Scripting.FileSystemObject
    
    
    'use for server date set thai date
'    thaidate = Date
'    dteDate = Date
'    intCurrMonth = Month(thaidate)
'    intCurrYear = Year(thaidate)
    
    fn = FreeFile
    GetiniFile
       
    'Initialize all Related Database
    strISSServer = "203.151.94.173"
    strISSUser = "iconnectmkt"
    strISSPWD = "tkmtcennoci2@!!"

    'ISS Database Connection
    strISSConnectionString = "Driver={SQL Server};Server=" + strISSServer + ";Database=" + strISSDBS + ";Uid=" + strISSUser + ";Pwd=" + strISSPWD + ";"
    Set conISS = New ADODB.Connection
    conISS.ConnectionString = strISSConnectionString
    conISS.Open

    Write_Log (": connected: " + strISSDBS)

    strISSConnectionString = "Driver={SQL Server};Server=" + strISSServer + ";Database=" + strISSDBA + ";Uid=" + strISSUser + ";Pwd=" + strISSPWD + ";"
    Set conISSA = New ADODB.Connection
    conISSA.ConnectionString = strISSConnectionString
    conISSA.Open
        
    Write_Log (": connected: " + strISSDBA)

    strComcode = strISSCompany
    strCompany = "PCK"

    Write_Log (": Company: " + strComcode)
    

    If FSO.FolderExists(strERP1UPath) = False Then
        MsgBox "พบข้อผิดพลาด!!! ในการเชื่อมต่อระบบบัญชี" & vbCrLf & "กรุณาตรวจสอบการเชื่อมต่อไปยังแฟ้มต่อไปนี้" & vbCrLf & " Drive:" & Mid(strERPPath, 1, 1) & " Folder=" & strERP1UPath & vbCrLf & "กรุณาตรวจสอบ Server Express.", vbOKOnly
        Write_Log ("พบข้อผิดพลาด!!! ในการเชื่อมต่อระบบบัญชี" & vbCrLf & "กรุณาตรวจสอบการเชื่อมต่อไปยังแฟ้มต่อไปนี้ Drive:" & Mid(strERPPath, 1, 1) & " Folder=" & strERP1UPath & vbCrLf & "กรุณาตรวจสอบ Server Express." & " :Call Main")
        Call SendEmailErr2("ISS-PCK: พบข้อผิดพลาด!!! ในการเชื่อมต่อระบบบัญชี", "กรุณาตรวจสอบการเชื่อมต่อไปยังแฟ้มต่อไปนี้ Drive:" & Mid(strERPPath, 1, 1) & " Folder=" & strERP1UPath, ":Call Main")
        End
    End If
        
    Write_Log (": connected1: " + strERPPath)
        
    Write_Log ("conERP.Provider = VFPOLEDB.1")
    conERP.Provider = "VFPOLEDB.1"
    
    Write_Log ("conERP.Properties(Data Source) = strERPPath")
    conERP.Properties("Data Source") = strERPPath
    
    Write_Log ("conERP.Properties(Collating Sequence) = THAI")
    'conERP.Properties("Jet OLEDB:Database Password") = strERPPWD
    conERP.Properties("Collating Sequence") = "THAI"
    
    Write_Log ("conERP.CursorLocation = adUseClient")
    conERP.CursorLocation = adUseClient
    
    Write_Log ("conERP.Open")
    conERP.Open
        
    Write_Log (": connected2: " + strERP2Path)
     
    conERP2.Provider = "VFPOLEDB.1"
    conERP2.Properties("Data Source") = strERP2Path
    conERP2.Properties("Collating Sequence") = "THAI"
    conERP2.CursorLocation = adUseClient
    conERP2.Open
    

    Call GenQuote
'    Call CheckSlmCode 'notuse
            
    
    Set rsERP = Nothing
    Set rsERPH = Nothing
    Set rsERPHR = Nothing
    Set rsERPD = Nothing
    Set rsISS = Nothing
    Set rsISSA = Nothing
    Set rsISSH = Nothing
    Set rsISSI = Nothing
    Set rsISSD = Nothing
    Set rsISSC = Nothing
    Set conERP = Nothing
    Set conERP2 = Nothing
    Set conISS = Nothing
    Set conISSA = Nothing
    
    Write_Log (": Program Run Successfully. . .")
    MsgBox "Program Run Successfully", vbOKOnly
    
    
    Exit Sub
Err_Handler:
    MsgBox "พบข้อผิดพลาด!!! ในการเชื่อมต่อระบบบัญชี" & vbCrLf & "กรุณาตรวจสอบการเชื่อมต่อ Server Express" & vbCrLf & vbCrLf & "Error : " & Err.Number & " " & Err.Description & vbCrLf & " :Call Main", vbOKOnly
    Write_Log ("Error : " & Err.Number & " " & Err.Description & " :Call Main")
    Call SendEmailErr2("ISS-PCK: พบข้อผิดพลาด!!! ในการเชื่อมต่อระบบบัญชี", "กรุณาตรวจสอบการเชื่อมต่อ Server Express", "Error : " & Err.Number & " " & Err.Description & vbCrLf & "   :Call Main")
    End
End Sub

Public Sub GetiniFile()
On Error GoTo Err_Handler

    Dim strLine As String

    
    Open "C:\WINDOWS\system32\ISS2PCK.ini" For Input As #1
    Do While Not EOF(1) ' Check for end of file.
        Input #1, strLine  ' Read data.
           
           If Mid$(strLine, 1, 11) = "WorkingDir=" Then
              strWorkingDir = Mid$(strLine, 12, Len(Trim(strLine)) - 11)
           End If
           
           If Mid$(strLine, 1, 7) = "ISSDBS=" Then
              strISSDBS = Mid$(strLine, 8, Len(Trim(strLine)) - 7)
           End If

           If Mid$(strLine, 1, 7) = "ISSDBA=" Then
              strISSDBA = Mid$(strLine, 8, Len(Trim(strLine)) - 7)
           End If

           If Mid$(strLine, 1, 8) = "ERP1Dir=" Then
              strERPPath = Mid$(strLine, 9, Len(Trim(strLine)) - 8)
           End If
           
           If Mid$(strLine, 1, 8) = "ERP2Dir=" Then
              strERP2Path = Mid$(strLine, 9, Len(Trim(strLine)) - 8)
           End If
           
           If Mid$(strLine, 1, 10) = "ERP1UPath=" Then
              strERP1UPath = Mid$(strLine, 11, Len(Trim(strLine)) - 10)
           End If
           
           If Mid$(strLine, 1, 8) = "EmailTo=" Then
              strEmailTo = Mid$(strLine, 9, Len(Trim(strLine)) - 8)
           End If
           
           If Mid$(strLine, 1, 8) = "EmailCC=" Then
              strEmailCC = Mid$(strLine, 9, Len(Trim(strLine)) - 8)
           End If
           
           If Mid$(strLine, 1, 12) = "EmailAdmin1=" Then
              strEmailAdmin1 = Mid$(strLine, 13, Len(Trim(strLine)) - 12)
           End If
           
           If Mid$(strLine, 1, 13) = "EmailSubject=" Then
              strEmailSubject = Mid$(strLine, 14, Len(Trim(strLine)) - 13)
           End If
           
           If Mid$(strLine, 1, 15) = "EmailSendOrder=" Then
              strEmailOrder = Mid$(strLine, 16, Len(Trim(strLine)) - 15)
           End If

           If Mid$(strLine, 1, 15) = "EmailSendQuata=" Then
              strEmailQuata = Mid$(strLine, 16, Len(Trim(strLine)) - 15)
           End If

           If Mid$(strLine, 1, 16) = "EmailSendReturn=" Then
              strEmailReturn = Mid$(strLine, 17, Len(Trim(strLine)) - 16)
           End If

           If Mid$(strLine, 1, 15) = "EmailSendStock=" Then
              strEmailStock = Mid$(strLine, 16, Len(Trim(strLine)) - 15)
           End If
 
           If Mid$(strLine, 1, 11) = "ISSCompany=" Then
              strISSCompany = Mid$(strLine, 12, Len(Trim(strLine)) - 11)
           End If
            
    Loop
    Close #1
        
    Write_Log (": Program Start. . .")
    Write_Log (": Database: " + strISSDBS)
    Write_Log (": WorkingDir: " + strWorkingDir)
    Write_Log (": ERPDBS: " + strERPPath)

    Exit Sub
Err_Handler:
    MsgBox "พบข้อผิดพลาด!!! กรุณาแจ้งทีม Support iSmartSales" & vbCrLf & "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ" & vbCrLf & vbCrLf & "Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call GetiniFile", vbOKOnly
    Write_Log ("Error : " & Err.Number & " " & Err.Description & " :Call GetiniFile")
    Call SendEmailErr2("ISS-PCK: พบข้อผิดพลาด!!! กรุณาแจ้งทีม Support iSmartSales", "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ", "Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call GetiniFile")
    End
End Sub

Public Sub GenQuote()
On Error GoTo Err_Handler

    Dim strDiscFormula As String
'    Dim Disc_amt, SumDisc_amt, Tot_amt As Currency
'    Dim lintNoOfRow As Integer
    Dim lTotaldisc As Currency
    Dim lMaxDiscSeq As Integer
    Dim lDocNumber As String
    
    
    Write_Log (": GenQuote Company=" & strComcode & ", Runtype=" & strRuntype & " Start. . .")

'=====BEGIN Admin1,2 by company=================
Set rsISSA = New ADODB.Recordset

strsql = "          SELECT Username,ID"
strsql = strsql & "  From WebUser"
strsql = strsql & "  where Username in ('pck')"
strsql = strsql & "  and division ='" & strComcode & "'"
strsql = strsql & "  order by Username"

rsISSA.Open strsql, conISSA, 1, 1
    
rsISSA.MoveFirst
Do While Not rsISSA.EOF

'    Set rsISS = New ADODB.Recordset
'
'    strsql = "          SELECT Application_ID "
'    strsql = strsql & "  FROM  WebUser_Application "
'    strsql = strsql & "  Where webuser_id = " & rsISSA.Fields("ID")
'    strsql = strsql & "  order by Application_ID"
'    rsISS.Open strsql, conISSA, 1, 1
'
'    rsISS.MoveFirst
'    strSlmlist = ""
'    Do While Not rsISS.EOF
'
'        strSlmlist = strSlmlist & "'" & rsISS.Fields("Application_ID") & "',"
'        rsISS.MoveNext
'    Loop
'    strSlmlist = Mid(strSlmlist, 1, Len(strSlmlist) - 1)
'    rsISS.Close
'    Set rsISS = Nothing
    

    '==Head BEGIN================================
    'Filename,DocNumber is Primarykey

    Set rsISSH = New ADODB.Recordset
    
    strsql = "         SELECT CustomerCode, SalesmanCode, CONVERT(VARCHAR(10),CAST(DocDate as date),103) as DocDate1, ShipCode "
    strsql = strsql & "     ,CONVERT(VARCHAR(10),CAST(ShipDate as date),103) as ShipDate1, CONVERT(VARCHAR(10),CAST(DocRefDate as date),103) as DocRefDate1 "
    strsql = strsql & "     ,Total_Sum, Total_Discount, Totla_HeadDiscount , Total_Vat , Total_Total , DocRefNumber "
    strsql = strsql & "     ,DocNumber , ShipName , DocRemark, Filename, DepartmentCode, CreditType, Other1, ContactPerson"
    If strRuntype = "QT" Then
    strsql = strsql & " From QuotationTransactionLoad"
    Else
    strsql = strsql & " From OrderTransactionLoad"
    End If
    strsql = strsql & " Where Company = '" & strComcode & "'"
    If strRuntype = "QT" Then
    strsql = strsql & " and status IS NULL"
    Else
    strsql = strsql & " and SOstatus IS NULL"
    End If
    strsql = strsql & " and customercode not like 'NewCus%'"
'    strsql = strsql & " and SalesmanCode in (" & strSlmlist & ")"
    strsql = strsql & " and DocCreateTime >= Convert(char(8),dateadd(day,-30,GETDATE()), 112)+'000000'"
'    strsql = strsql & " and Docdate >= '20181024' "
    
    'TEST
    'SO
'    strsql = strsql & " and filename = 'PCKPPV_transaction_20190124175054.txt'"
'    strsql = strsql & " and filename IN ('PCKPPV_transaction_20181106144459.txt')" 'so1
'    strsql = strsql & " and filename IN ('PCKPPV_transaction_20181106144459.txt','PCKPPV_transaction_20181106175653.txt')" 'so2-8
'    strsql = strsql & " and filename = 'PCKPPV_transaction_20181108091950.txt'"  'so9,transportby
    
    'QT
'    strsql = strsql & " and filename = 'PCKPPV_transaction_20181126173218.txt'" 'remark60char
'    strsql = strsql & " and filename = 'PCKPPV_transaction_20181112140203.txt'" 'ex1-6
'    strsql = strsql & " and filename = 'PCKPPV_transaction_20181022161417.txt'" 'beer2
'    strsql = strsql & " and filename = 'PCKKNC_transaction_20181022153209.txt'" 'beer1
'    strsql = strsql & " and Filename IN ('PCKPPV_transaction_20181019131711.txt','PCKPPV_transaction_20181019163148.txt')"  'ex1 to 5,6
'    strsql = strsql & " and filename = 'PCKPPV_transaction_20181009165121.txt'" 'PCKPPV-000003.pck 4.WAT
'    strsql = strsql & " and filename = 'PCKPPV_transaction_20181018154722.txt'" 'PCK+WAT err
'    strsql = strsql & " and DocNumber = ''"
'    strsql = strsql & " and DocCreateTime >= '20170906000000'"
'    strsql = strsql & " and SalesmanCode =''"

    strsql = strsql & " Order by DocDate,DocNumber"

    rsISSH.Open strsql, conISS, 1, 1
    
    If Not rsISSH.EOF Then
        strSOlist = ""
        rsISSH.MoveFirst
        intNoOfRow = rsISSH.RecordCount
        
        '*use index of table
        'run 2 step don't change code.
        strsql = " USE OESO.DBF INDEX OESO.CDX"
        conERP.Execute strsql
        strsql = " USE OESO.DBF INDEX OESO.CDX"
        conERP2.Execute strsql
        
        strsql = " USE OESOIT.DBF INDEX OESOIT.CDX"
        conERP.Execute strsql
        strsql = " USE OESOIT.DBF INDEX OESOIT.CDX"
        conERP2.Execute strsql
        
    Else
        Write_Log (": GenQuote Not found " & strEngdoc & " . . .")
        MsgBox "Not found " & strEngdoc & "!", vbOKOnly
        strSOlist = ""
        intNoOfRow = 0
    End If

    Do While Not rsISSH.EOF
    
        '=====COMPANY BEGIN======================================
        Set rsISSC = New ADODB.Recordset
        
        strsql = "          SELECT Distinct LEFT(ProductCode,3) as strcom"
        If strRuntype = "QT" Then
        strsql = strsql & " From QuotationTransactionItem t"
        Else
        strsql = strsql & " From OrderTransactionItem t"
        End If
        strsql = strsql & " where Filename ='" & rsISSH.Fields("Filename") & "'"
        strsql = strsql & " and DocNumber ='" & rsISSH.Fields("DocNumber") & "'"
        strsql = strsql & " Order by strcom"
        
        rsISSC.Open strsql, conISS, 1, 1
        
        rsISSC.MoveFirst
        
        Do While Not rsISSC.EOF
        
            If rsISSC.RecordCount > 1 Then
                Call Update_StatusError
                MsgBox "ข้อผิดพลาด: ไม่สามารถสร้าง" & strThaidoc & " เนื่องจากมี สินค้า 2 บริษัทปนกัน ให้แจ้งพนักงานขาย ส่งมาใหม่" & vbCrLf & strEngdoc & ": " & rsISSH.Fields("DocNumber") & vbCrLf & vbCrLf & "Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call GenQuote ,Filename=" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix, vbOKOnly
                Write_Log ("Error ไม่สามารถสร้าง" & strThaidoc & " เนื่องจากมี สินค้า 2 บริษัทปนกัน ให้แจ้งพนักงานขาย ส่งมาใหม่" & vbCrLf & strEngdoc & ": " & rsISSH.Fields("DocNumber") & vbCrLf & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call GenQuote ,Filename=" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
                Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! ไม่สามารถสร้าง" & strThaidoc & " เนื่องจากมี สินค้า 2 บริษัทปนกัน ให้แจ้งพนักงานขาย ส่งมาใหม่", strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ข้อผิดพลาด:ไม่สามารถสร้าง" & strThaidoc & " เนื่องจากมี สินค้า 2 บริษัทปนกัน ให้แจ้งพนักงานขาย ส่งมาใหม่", "Error : " & Err.Number & " " & Err.Description & " :Call GenQuote ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Doctype=" & strPrefix)
                Exit Do
            End If
            
            Set rsERPH = New ADODB.Recordset
            strsql = "SELECT * FROM OESO WHERE 1=2 "
            
            If rsISSC.Fields("strcom") = "PCK" Then
                rsERPH.Open strsql, conERP, adOpenDynamic, adLockOptimistic
                
                conERP.BeginTrans
            Else
                rsERPH.Open strsql, conERP2, adOpenDynamic, adLockOptimistic
                
                conERP2.BeginTrans
            End If

            'log begin trans order iss
            Write_Log (": BeginTrans DocNumber=" & rsISSH.Fields("DocNumber") & ", Runtype=" & strRuntype)
    
            rsERPH.AddNew
    
            Call GetQTno_id
            
            Call PrepareQTHD
            
            rsERPH.Update
            rsERPH.Close
            Set rsERPH = Nothing
            
            
            '=====ITEM BEGIN======================================
            If strRuntype = "QT" Then
            Write_Log (": QuotationTransactionItem " & ", Runtype=" & strRuntype & " Start. . .")
            Else
            Write_Log (": OrderTransactionItem " & ", Runtype=" & strRuntype & " Start. . .")
            End If
    
            
            Set rsERPD = New ADODB.Recordset
            strsql = "SELECT * FROM OESOIT WHERE 1=2"
            
            If rsISSC.Fields("strcom") = "PCK" Then
                rsERPD.Open strsql, conERP, adOpenDynamic, adLockOptimistic
            Else
                rsERPD.Open strsql, conERP2, adOpenDynamic, adLockOptimistic
            End If
                        
            Set rsISSI = New ADODB.Recordset
            
            strsql = "          SELECT "
            strsql = strsql & "        Filename, DocNumber , ItemNumber , DiscountRefNo , ProductCode "
            strsql = strsql & "      , ProductQuan , ProductUnit , ProductPrice , ItemRemark "
            If strRuntype = "QT" Then
            strsql = strsql & " From QuotationTransactionItem"
            Else
            strsql = strsql & " From OrderTransactionItem"
            End If
            strsql = strsql & " where Filename ='" & rsISSH.Fields("Filename") & "'"
            strsql = strsql & " and DocNumber ='" & rsISSH.Fields("DocNumber") & "'"
            strsql = strsql & " Order by ItemNumber"
            
            rsISSI.Open strsql, conISS, 1, 1
            
            rsISSI.MoveFirst
    '        lintNoOfRow = rsISSI.RecordCount
            lMaxDiscSeq = 0
            
            Do While Not rsISSI.EOF
                rsERPD.AddNew
                
                '=====Item DISCOUNT BEGIN======================================
                
                strDiscFormula = ""
                If rsISSI.Fields("ProductPrice") > 0 Then
                
                    Set rsISSD = New ADODB.Recordset
                    
                    strsql = "          SELECT "
                    strsql = strsql & "        Filename, DocNumber , DiscountRefNo , DiscountSequence "
                    strsql = strsql & "      , DiscountCode , DiscountType , DiscountAmount "
                    strsql = strsql & "      , DiscountUnit "
                    If strRuntype = "QT" Then
                    strsql = strsql & " From QuotationTransactionDiscount"
                    Else
                    strsql = strsql & " From OrderTransactionDiscount"
                    End If
                    strsql = strsql & " Where Filename ='" & rsISSI.Fields("Filename") & "'"
                    strsql = strsql & " and DocNumber ='" & rsISSI.Fields("DocNumber") & "'"
                    strsql = strsql & " and DiscountRefNo =" & rsISSI.Fields("DiscountRefNo")
                    strsql = strsql & " and DiscountRefNo <> 0"
                    strsql = strsql & " Order by DiscountRefNo,DiscountSequence"
                    
                    rsISSD.Open strsql, conISS, 1, 1
                    
                    If Not rsISSD.EOF Then
                        'Write_Log (": QuotationTransactionDiscount" & ", Runtype=" & strRuntype & " Start. . .")
                        rsISSD.MoveFirst
    '                    Disc_amt = 0
    '                    SumDisc_amt = 0
    '                    Tot_amt = 0
                        lTotaldisc = 0
                    End If
                    
                    Do While Not rsISSD.EOF
        
                        Select Case rsISSD.Fields("DiscountType")
                            Case "P"
                                strDiscFormula = rsISSD.Fields("DiscountType")
                                lTotaldisc = lTotaldisc + rsISSD.Fields("DiscountAmount")
                                
    '                            If rsISSD.Fields("DiscountSequence") = 1 Then
    '                                Disc_amt = Round(rsISSI.Fields("ProductQuan") * rsISSI.Fields("ProductPrice") * rsISSD.Fields("DiscountAmount") / 100, 2)
    '                                Tot_amt = (rsISSI.Fields("ProductQuan") * rsISSI.Fields("ProductPrice")) - Disc_amt
    '                                SumDisc_amt = SumDisc_amt + Disc_amt
    '                            Else
    '                                Disc_amt = Round(Tot_amt * rsISSD.Fields("DiscountAmount") / 100, 2)
    '                                Tot_amt = Tot_amt - Disc_amt
    '                                SumDisc_amt = SumDisc_amt + Disc_amt
    '                            End If
                                
                            Case "Q"
                                strDiscFormula = rsISSD.Fields("DiscountType")
                                lTotaldisc = lTotaldisc + (rsISSD.Fields("DiscountAmount") * rsISSI.Fields("ProductQuan"))
                                
    '                            If rsISSD.Fields("DiscountSequence") = 1 Then
    '                                Disc_amt = rsISSI.Fields("ProductQuan") * rsISSD.Fields("DiscountAmount")
    '                                Tot_amt = (rsISSI.Fields("ProductQuan") * rsISSI.Fields("ProductPrice")) - Disc_amt
    '                                SumDisc_amt = SumDisc_amt + Disc_amt
    '                            Else
    '                                Disc_amt = rsISSI.Fields("ProductQuan") * rsISSD.Fields("DiscountAmount")
    '                                Tot_amt = Tot_amt - Disc_amt
    '                                SumDisc_amt = SumDisc_amt + Disc_amt
    '                            End If
                                
                        End Select
                        If rsISSD.Fields("DiscountSequence") > lMaxDiscSeq Then lMaxDiscSeq = rsISSD.Fields("DiscountSequence")
                        
                        rsISSD.MoveNext
                    Loop
                    rsISSD.Close
                    Set rsISSD = Nothing
                    
    '                Write_Log (": QuotationTransactionDiscount" & ", Runtype=" & strRuntype & " Finished. . .")
                End If
    
                '====Item DISCOUNT END========================================
                
                If Len(Trim(strDiscFormula)) > 0 Then
                
                    'Discount%P or Q
                    If strDiscFormula = "Q" Then
                        strDiscFormula = Format(lTotaldisc, "#######.00")
                        rsERPD.Fields("DISC") = Space(10 - Len(strDiscFormula)) & strDiscFormula
                        rsERPD.Fields("DISCAMT") = lTotaldisc
                    
                    ElseIf strDiscFormula = "P" Then
                        strDiscFormula = lTotaldisc & "%"
                        rsERPD.Fields("DISC") = Space(10 - Len(strDiscFormula)) & strDiscFormula
                        rsERPD.Fields("DISCAMT") = Round(rsISSI.Fields("ProductQuan") * rsISSI.Fields("ProductPrice") * lTotaldisc / 100, 2)
                    End If
                    rsERPD.Fields("TRNVAL") = (rsISSI.Fields("ProductQuan") * rsISSI.Fields("ProductPrice")) - rsERPD.Fields("DISCAMT")
                    
                Else
                    'Free or No_discount
                    rsERPD.Fields("DISC") = " "
                    rsERPD.Fields("DISCAMT") = 0
                    rsERPD.Fields("TRNVAL") = rsISSI.Fields("ProductQuan") * rsISSI.Fields("ProductPrice")
                End If
                
                
                Call PrepareQTDT
                
                rsERPD.Update
                
                rsISSI.MoveNext
                    
            Loop
            
            rsISSI.Close
            rsERPD.Close
            Set rsISSI = Nothing
            Set rsERPD = Nothing
            
            If strRuntype = "QT" Then
            Write_Log (": QuotationTransactionItem" & ", Runtype=" & strRuntype & " Finished. . .")
            Else
            Write_Log (": OrderTransactionItem" & ", Runtype=" & strRuntype & " Finished. . .")
            End If
            
            '====ITEM END========================================
            
            If rsISSH.Fields("Total_Discount") > 0 And lMaxDiscSeq > 1 Then
                Call Update_TotalDiscount
            End If
            
            If Len(Trim(rsISSH.Fields("DocRemark"))) > 0 Then
                Call PrepareQTHDRemark
            End If
            
            Call CheckHeaderDetail
    
    
            If rsISSC.Fields("strcom") = "PCK" Then
                conERP.CommitTrans
            Else
                conERP2.CommitTrans
            End If
               
            'log commit tran issorder
            Write_Log (": CommitTrans DocNumber=" & rsISSH.Fields("DocNumber") & " SOno=" & strSOno & " ,Doctype=" & strPrefix)
            
            If strRuntype = "QT" Then
            
                lDocNumber = rsISSH.Fields("DocNumber")
                'UPdate status ='C' complete ,DocNumber =Expressno
                strsql = "          Update QuotationTransactionLoad"
                strsql = strsql & " set status = 'C', DocRefNumber = DocNumber, DocNumber ='" & strSOno & "'"
                strsql = strsql & "    ,DocRemark = 'ISS Number: " & lDocNumber & " '+DocRemark"
                strsql = strsql & " where Company ='" & strComcode & "'"
                strsql = strsql & " and Filename ='" & rsISSH.Fields("Filename") & "'"
                strsql = strsql & " and DocNumber ='" & lDocNumber & "'"
                conISS.Execute strsql
    
                strsql = "          Update QuotationTransaction"
                strsql = strsql & " set DocRefNumber = DocNumber, DocNumber ='" & strSOno & "'"
                strsql = strsql & "    ,DocRemark = 'ISS Number: " & lDocNumber & " '+DocRemark"
                strsql = strsql & " where Company ='" & strComcode & "'"
                strsql = strsql & " and Filename ='" & rsISSH.Fields("Filename") & "'"
                strsql = strsql & " and DocNumber ='" & lDocNumber & "'"
                conISS.Execute strsql
    
                strsql = "          Update QuotationTransactionItem"
                strsql = strsql & " set DocNumber ='" & strSOno & "'"
                strsql = strsql & " where Filename ='" & rsISSH.Fields("Filename") & "'"
                strsql = strsql & " and DocNumber ='" & lDocNumber & "'"
                conISS.Execute strsql
    
                strsql = "          Update QuotationTransactionDiscount"
                strsql = strsql & " set DocNumber ='" & strSOno & "'"
                strsql = strsql & " where Filename ='" & rsISSH.Fields("Filename") & "'"
                strsql = strsql & " and DocNumber ='" & lDocNumber & "'"
                conISS.Execute strsql
            
            Else   'SO
                'UPdate status ='C' complete
                strsql = "          Update OrderTransactionLoad"
                strsql = strsql & " set sostatus = 'C'"
                strsql = strsql & " where Company ='" & strComcode & "'"
                strsql = strsql & " and Filename ='" & rsISSH.Fields("Filename") & "'"
                strsql = strsql & " and DocNumber ='" & rsISSH.Fields("DocNumber") & "'"
                conISS.Execute strsql
            End If


            'fast but everyone get out program Express
            '*Unlock Record
            strsql = "          SET EXCLUSIVE OFF;"
    
            If rsISSC.Fields("strcom") = "PCK" Then
                conERP.Execute strsql
            Else
                conERP2.Execute strsql
            End If
    
            Write_Log (": SOno=" & strSOno & " ,Filename=" & rsISSH.Fields("Filename") & " ,DocNumber=" & rsISSH.Fields("DocNumber") & " ,SalesmanCode=" & rsISSH.Fields("SalesmanCode") & " ,CustomerCode=" & rsISSH.Fields("CustomerCode") & ", Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix & " ,Status=Complete")
            strSOlist = strSOlist & "บริษัท:" & rsISSC.Fields("strcom") & "   " & strEngdoc & " No: " & rsISSH.Fields("DocNumber") & "=>" & strThaidoc & "เลขที่: " & strSOno & vbCrLf

            rsISSC.MoveNext
        Loop
        rsISSC.Close
        Set rsISSC = Nothing
        '=====COMPANY END======================================
        
        rsISSH.MoveNext
    Loop
    rsISSH.Close
    Set rsISSH = Nothing

    If Len(Trim(strSOlist)) > 0 And intNoOfRow > 0 Then Call SendEmailSOlist

    
    '==Head END===============================
'    Write_Log (": GenQuote Finished. . .")
    
    
    '======SendEmailOf Quatation,Return,Stock
'    If strEmailQuata = "Y" Then Call SendEmailOf("QUATA") 'NotUse
'    If strEmailReturn = "Y" Then Call SendEmailOf("RETURN")  'TEST
'    If strEmailStock = "Y" Then Call SendEmailOf("STOCK")

    rsISSA.MoveNext
Loop
rsISSA.Close
Set rsISSA = Nothing

'=====END Admin1,2 by company=================


    Exit Sub
Err_Handler:
    If Len(Trim(strSOlist)) > 0 Then Call SendEmailSOlist
    If Err.Number <> -2147217865 Then
        Call Update_StatusError
    End If
    MsgBox "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ" & vbCrLf & vbCrLf & "Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call GenQuote ,Filename=" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix, vbOKOnly
    Write_Log ("Error : " & Err.Number & " " & Err.Description & " :Call GenQuote ,Filename=" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
    Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! กรุณาแจ้งทีม Support iSmartSales", strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ", "Error : " & Err.Number & " " & Err.Description & " :Call GenQuote ,Filename=" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & ", Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
    If rsISSC.Fields("strcom") = "PCK" Then
        conERP.RollbackTrans
    Else
        conERP2.RollbackTrans
    End If
    'log rollback
    Write_Log (": RollbackTrans DocNumber=" & rsISSH.Fields("DocNumber") & ", Runtype=" & strRuntype)
    End
End Sub

Public Sub CheckSlmCode()

Dim Comstrslmlist As String

  
'    Write_Log (": CheckOrder salemancode isn't link with admin user PCK Company=" & strComcode & " Start. . .")

    '=====Get slmlist by company=================
    Set rsISSA = New ADODB.Recordset
    
    strsql = "          SELECT Username,ID,Application_ID"
    strsql = strsql & "  From WebUser u, WebUser_Application a"
    strsql = strsql & "  Where u.Username in ('pck')"
    strsql = strsql & "    and u.division ='" & strComcode & "'"
    
    strsql = strsql & "    and u.id = a.webuser_id "
    strsql = strsql & "  order by Username,Application_ID"
    
    rsISSA.Open strsql, conISSA, 1, 1
        
    rsISSA.MoveFirst
    Comstrslmlist = ""
    Do While Not rsISSA.EOF
    
        Comstrslmlist = Comstrslmlist & "'" & rsISSA.Fields("Application_ID") & "',"
        rsISSA.MoveNext
    Loop
    Comstrslmlist = Mid(Comstrslmlist, 1, Len(Comstrslmlist) - 1)
    rsISSA.Close
    Set rsISSA = Nothing

    '==Check salesmancode isn't link with Admin_user pck ================================
    'Filename,DocNumber is Primarykey

    Set rsISSH = New ADODB.Recordset
    
    strsql = "         SELECT Distinct SalesmanCode  "
    If strRuntype = "QT" Then
    strsql = strsql & " From QuotationTransactionLoad"
    Else
    strsql = strsql & " From OrderTransactionLoad"
    End If
    
    strsql = strsql & " where Company = '" & strComcode & "'"
    strsql = strsql & " and status IS NULL"
    strsql = strsql & " and customercode not like 'NewCus%'"
    strsql = strsql & " and DocCreateTime >= Convert(char(8),dateadd(day,-30,GETDATE()), 112)+'000000'"
    strsql = strsql & " and SalesmanCode not in (" & Comstrslmlist & ")"
    strsql = strsql & " Order by SalesmanCode"

    rsISSH.Open strsql, conISS, 1, 1
    
    If Not rsISSH.EOF Then
        rsISSH.MoveFirst
    End If

    Do While Not rsISSH.EOF
        Call SendEmailErrSlm("ISS-PCK:พบข้อผิดพลาด!!! ตรวจพบSaleman " & rsISSH.Fields("SalesmanCode") & " ไม่พบAdmin ดูแล", "พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ตรวจพบข้อผิดพลาด กรุณาติดต่อ Support iSmartSales เพื่อกำหนด Admin ผู้ดูแลพนักงานขาย", "Error : SalesmanCode : " & rsISSH.Fields("SalesmanCode") & " isn't link with Admin_user PCK!")

        rsISSH.MoveNext
    Loop
    rsISSH.Close
    Set rsISSH = Nothing

End Sub

Public Sub CheckHeaderDetail()
On Error GoTo Err_Handler

    
    Write_Log (": CheckHeader" & " ,Doctype=" & strPrefix & " Start. . .")

    'Check Header
    Set rsERP = New ADODB.Recordset
    
    strsql = "        SELECT count(*) AS count1"
    strsql = strsql & " FROM OESO"
    strsql = strsql & " Where SONUM = '" & strSOno & "'"
    
    If rsISSC.Fields("strcom") = "PCK" Then
        rsERP.Open strsql, conERP, 1, 1
    Else
        rsERP.Open strsql, conERP2, 1, 1
    End If

    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
    
        If rsERP.Fields("count1") = 0 Then
        
            If Len(Trim(strSOlist)) > 0 Then Call SendEmailSOlist
            Call Update_StatusError
            MsgBox "ตรวจพบข้อผิดพลาด การสร้างข้อมูลไม่สมบูรณ์ กรุณารอรอบการโหลดออเดอร์ครั้งต่อไป" & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "Error ไม่พบข้อมูล Header data, " & strEngdoc & ":" & strSOno & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call CheckHeaderDetail ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix, vbOKOnly
            Write_Log ("Error ไม่พบข้อมูล Header data, " & strEngdoc & ":" & strSOno & vbCrLf & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call CheckHeaderDetail ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
            Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! กรุณารอรอบการโหลดออเดอร์ครั้งต่อไป", strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ตรวจพบข้อผิดพลาด การสร้างข้อมูลไม่สมบูรณ์ กรุณารอรอบการโหลดออเดอร์ครั้งต่อไป", "Error ไม่พบข้อมูล Header data, " & strEngdoc & "=" & strSOno & vbCrLf & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call CheckHeaderDetail ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & ", Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
            If rsISSC.Fields("strcom") = "PCK" Then
                conERP.RollbackTrans
            Else
                conERP2.RollbackTrans
            End If
            'log rollback
            Write_Log (": RollbackTrans DocNumber=" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode"))
            End
            
        End If
        
        Exit Do
'        rsERP.MoveNext
    Loop
    rsERP.Close
    Set rsERP = Nothing
        
        
    Write_Log (": CheckDetail" & " ,Doctype=" & strPrefix & " Start. . .")

    'Check Detail
    Set rsERP = New ADODB.Recordset
    
    strsql = "        SELECT count(*) AS count1"
    strsql = strsql & " FROM OESOIT"
    strsql = strsql & " Where SONUM = '" & strSOno & "'"
    
    If rsISSC.Fields("strcom") = "PCK" Then
        rsERP.Open strsql, conERP, 1, 1
    Else
        rsERP.Open strsql, conERP2, 1, 1
    End If

    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
    
        If rsERP.Fields("count1") = 0 Then
        
            If Len(Trim(strSOlist)) > 0 Then Call SendEmailSOlist
            Call Update_StatusError
            MsgBox "ตรวจพบข้อผิดพลาด การสร้างข้อมูลไม่สมบูรณ์ กรุณารอรอบการโหลดออเดอร์ครั้งต่อไป" & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "Error ไม่พบข้อมูล Detail data, " & strEngdoc & "=" & strSOno & vbCrLf & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call CheckHeaderDetail ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & ", Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix, vbOKOnly
            Write_Log ("Error ไม่พบข้อมูล Detail data, " & strEngdoc & "=" & strSOno & vbCrLf & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call CheckHeaderDetail ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & ", Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
            Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! กรุณารอรอบการโหลดออเดอร์ครั้งต่อไป", strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ตรวจพบข้อผิดพลาด การสร้างข้อมูลไม่สมบูรณ์ กรุณารอรอบการโหลดออเดอร์ครั้งต่อไป", "Error ไม่พบข้อมูล Detail data, " & strEngdoc & "=" & strSOno & vbCrLf & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call CheckHeaderDetail ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & ", Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
            If rsISSC.Fields("strcom") = "PCK" Then
                conERP.RollbackTrans
            Else
                conERP2.RollbackTrans
            End If
            'log rollback
            Write_Log (": RollbackTrans DocNumber=" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Doctype=" & strPrefix)
            End
            
        End If
        
        Exit Do
'        rsERP.MoveNext
    Loop
    rsERP.Close
    Set rsERP = Nothing
    
    
    Write_Log (": CheckHeaderDetail" & " ,Doctype=" & strPrefix & " Finish. . .")

    Exit Sub
Err_Handler:
    If Len(Trim(strSOlist)) > 0 Then Call SendEmailSOlist
    Call Update_StatusError
    MsgBox "ตรวจพบข้อผิดพลาด การสร้างข้อมูลไม่สมบูรณ์ กรุณารอรอบการโหลดออเดอร์ครั้งต่อไป" & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call CheckHeaderDetail ,Filename=" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Doctype=" & strPrefix, vbOKOnly
    Write_Log ("Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call CheckHeaderDetail ,Filename=" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Doctype=" & strPrefix)
    Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! กรุณารอรอบการโหลดออเดอร์ครั้งต่อไป", strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ตรวจพบข้อผิดพลาด การสร้างข้อมูลไม่สมบูรณ์ กรุณารอรอบการโหลดออเดอร์ครั้งต่อไป", "Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call CheckHeaderDetail ,Filename=" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & ", Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
    If rsISSC.Fields("strcom") = "PCK" Then
        conERP.RollbackTrans
    Else
        conERP2.RollbackTrans
    End If
    'log rollback
    Write_Log (": RollbackTrans DocNumber=" & rsISSH.Fields("DocNumber") & " ,Doctype=" & strPrefix)
    End
End Sub


Public Sub Update_TotalDiscount()
On Error GoTo Err_Handler

    Dim ltotalsum As Currency
    
'    Write_Log (": Update_TotalDiscount Start. . .")
        
        
    '== Update SUM(TRNVAL) = OrderTransaction.Total_Sum ==============================================
    ' If lMaxDiscSeq > 1 then
    '    user_key disc > 1 step and i will total many disc to 1 disc only.
    Set rsERP = New ADODB.Recordset
    
    strsql = "        SELECT SUM(TRNVAL) AS totalsum"
    strsql = strsql & " FROM OESOIT "
    strsql = strsql & " Where SONUM = '" & strSOno & "'"
    
    If rsISSC.Fields("strcom") = "PCK" Then
        rsERP.Open strsql, conERP, 1, 1
    Else
        rsERP.Open strsql, conERP2, 1, 1
    End If
    
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
    
        ltotalsum = rsERP.Fields("totalsum")
        rsERP.MoveNext
        
    Loop
    rsERP.Close
    Set rsERP = Nothing
    
    
    '==Update totalsum in SOHD
    'example: discount_line 2 step 2%+3%=5%
    strsql = "          UPDATE OESO"
    strsql = strsql & " SET AMOUNT = " & ltotalsum
    strsql = strsql & "    ,TOTAL = " & ltotalsum & " - DISCAMT "
    strsql = strsql & "    ,NETAMT = ROUND((" & ltotalsum & " - DISCAMT) + ((" & ltotalsum & " - DISCAMT) *" & gVat_rate & "/100),2)"
    strsql = strsql & "    ,VATAMT = ROUND((" & ltotalsum & " - DISCAMT) *" & gVat_rate & "/100,2)"
    strsql = strsql & "    ,NETVAL = TOTAL "
    strsql = strsql & " Where SONUM = '" & strSOno & "'"

    If rsISSC.Fields("strcom") = "PCK" Then
        conERP.Execute strsql
    Else
        conERP2.Execute strsql
    End If
    
    
'    Write_Log (": Update_TotalDiscount Finish. . .")

    Exit Sub
Err_Handler:
    If Len(Trim(strSOlist)) > 0 Then Call SendEmailSOlist
    Call Update_StatusError
    MsgBox "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ" & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call Update_TotalDiscount ,Filename=" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix, vbOKOnly
    Write_Log ("Error : " & Err.Number & " " & Err.Description & vbCrLf & " :Call Update_TotalDiscount ,Filename=" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
    Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! กรุณาแจ้งทีม Support iSmartSales", strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ", "Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call Update_TotalDiscount ,Filename=" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & ", Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
    If rsISSC.Fields("strcom") = "PCK" Then
        conERP.RollbackTrans
    Else
        conERP2.RollbackTrans
    End If
    'log rollback
    Write_Log (": RollbackTrans DocNumber=" & rsISSH.Fields("DocNumber"))
    End
End Sub

Public Sub PrepareQTHD()
On Error GoTo Err_Handler

    Dim strDiscFormula As String
    Dim lTotaldisc As Currency
    Dim lDiscountSequence As Integer
    Dim ldate1 As Date

    
    Write_Log (": PrepareQTHD DocNumber=" & rsISSH.Fields("DocNumber") & " SOno=" & strSOno & " ,Runtype=" & strRuntype)
    'log headqua issor, qano.

    '== ARMAS ==============================================
    Set rsERP = New ADODB.Recordset
    
'    Write_Log (": Read ARMAS . . .")
    
    strsql = ""
    strsql = strsql & " Select cuscod, areacod, paytrm, orgnum"
    strsql = strsql & " From ARMAS "
    strsql = strsql & " Where cuscod ='" & Mid(rsISSH.Fields("CustomerCode"), 4) & "'"
    
    If rsISSC.Fields("strcom") = "PCK" Then
        rsERP.Open strsql, conERP, 1, 1
    Else
        rsERP.Open strsql, conERP2, 1, 1
    End If

    If Not rsERP.EOF Then
        rsERP.MoveFirst
    Else
        If Len(Trim(strSOlist)) > 0 Then Call SendEmailSOlist
        Call Update_StatusError
        MsgBox "ข้อผิดพลาด: ไม่พบข้อมูลลูกค้า " & Mid(rsISSH.Fields("CustomerCode"), 4) & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call PrepareQTHD ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix, vbOKOnly
        Write_Log ("ข้อผิดพลาด: ไม่พบข้อมูลลูกค้า " & Mid(rsISSH.Fields("CustomerCode"), 4) & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call PrepareQTHD ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
        Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! ไม่พบข้อมูลลูกค้า " & Mid(rsISSH.Fields("CustomerCode"), 4) & " ในบริษัท " & rsISSC.Fields("strcom"), strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ข้อผิดพลาด:ไม่พบข้อมูลลูกค้า " & Mid(rsISSH.Fields("CustomerCode"), 4) & " ในบริษัท " & rsISSC.Fields("strcom"), "Error : " & Err.Number & " " & Err.Description & " :Call PrepareQTHD ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & ", Company=" & rsISSC.Fields("strcom") & ", Comcode=" & strComcode & " ,Doctype=" & strPrefix)
        If rsISSC.Fields("strcom") = "PCK" Then
            conERP.RollbackTrans
        Else
            conERP2.RollbackTrans
        End If
        'log rollback
        Write_Log (": RollbackTrans DocNumber=" & rsISSH.Fields("DocNumber"))
        End
    End If
    
    Do While Not rsERP.EOF
    
        rsERPH.Fields("CUSCOD") = rsERP.Fields("cuscod")
        If strRuntype = "QT" Then
            rsERPH.Fields("AREACOD") = " "
        Else  'SO
            rsERPH.Fields("AREACOD") = rsERP.Fields("areacod")
        End If
        rsERPH.Fields("PAYTRM") = rsERP.Fields("paytrm")
        intpaytrm = rsERP.Fields("paytrm")
        rsERPH.Fields("ORGNUM") = rsERP.Fields("orgnum")
                
        rsERP.MoveNext
    Loop
    rsERP.Close
    Set rsERP = Nothing
    
    
    '== ISINFO =============================
'    Write_Log (": Read ISINFO . . .")
    
    Set rsERP = New ADODB.Recordset
    
    strsql = ""
    strsql = strsql & " Select vatrat"
    strsql = strsql & " From ISINFO"
    
    If rsISSC.Fields("strcom") = "PCK" Then
        rsERP.Open strsql, conERP, 1, 1
    Else
        rsERP.Open strsql, conERP2, 1, 1
    End If
    
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
    
        rsERPH.Fields("VATRAT") = rsERP.Fields("vatrat")
        
'        If Mid(rsISSH.Fields("DepartmentCode"), 3, 1) = "Y" Then
            gVat_rate = rsERP.Fields("vatrat") 'Vat=7
'        Else
'            gVat_rate = 0                      'Vat=0
'        End If
        
        rsERP.MoveNext
    Loop
    rsERP.Close
    Set rsERP = Nothing
    
    
    '== ISTAB =============================
    '= transport by
    Write_Log (": Read ISTAB . . .")
    
    If Len(Trim(rsISSH.Fields("Other1"))) > 0 Then
        rsERPH.Fields("DLVBY") = Mid(rsISSH.Fields("Other1"), 4)
    Else
        rsERPH.Fields("DLVBY") = " "
    End If
'    Set rsERP = New ADODB.Recordset
'
'    strsql = ""
'    strsql = strsql & " Select typcod, typdes"
'    strsql = strsql & " From ISTAB"
'    strsql = strsql & " WHERE tabtyp ='41'"
'    strsql = strsql & "   AND typcod ='" & rsISSH.Fields("Other1") & "'"
'
'    rsERP.Open strsql, conERP, 1, 1
'
'    If Not rsERP.EOF Then
'        rsERP.MoveFirst
'    Else
'        rsERPH.Fields("DLVBY") = " "
'    End If
'
'    Do While Not rsERP.EOF
'
'        rsERPH.Fields("DLVBY") = Trim(rsERP.Fields("typcod"))
'
'        rsERP.MoveNext
'    Loop
'    rsERP.Close
'    Set rsERP = Nothing
    
    
    '==SOHD Other Field=============================
    
    rsERPH.Fields("SONUM") = strSOno
    'dd/mm/yyyy
    ldate1 = CDate(rsISSH.Fields("DocDate1"))
    ldate1 = DateAdd("yyyy", 543, ldate1)
    rsERPH.Fields("SODAT") = ldate1
    
'    If Mid(rsISSH.Fields("DepartmentCode"), 3, 1) = "Y" Then
        rsERPH.Fields("FLGVAT") = "2"  'ExcludeVat
'    Else
'        rsERPH.Fields("FLGVAT") = "0"  'NoVat
'    End If
    
    rsERPH.Fields("DEPCOD") = " "
    rsERPH.Fields("SLMCOD") = Mid(rsISSH.Fields("SalesmanCode"), 4)
    rsERPH.Fields("SHIPTO") = " "
    rsERPH.Fields("RFF") = " "
    
    If strRuntype = "QT" Then
    
        rsERPH.Fields("SORECTYP") = "5"
        rsERPH.Fields("YOUREF") = Mid(rsISSH.Fields("ShipName"), 1, 30)
        'dd/mm/yyyy
        ldate1 = CDate(rsISSH.Fields("DocDate1"))
        ldate1 = DateAdd("yyyy", 543, ldate1)
        ldate1 = DateAdd("d", intpaytrm, ldate1)
        rsERPH.Fields("DLVDAT") = ldate1
        
    Else   'SO
        rsERPH.Fields("SORECTYP") = "0"
        
        'ContactPerson =QuotationNo
        'DocRefNumber  =PONo
        'DocNumber =3 digit
        rsERPH.Fields("YOUREF") = Mid(rsISSH.Fields("ContactPerson") & "," & rsISSH.Fields("DocRefNumber") & "," & Mid(rsISSH.Fields("DocNumber"), 11, 3), 1, 30)

        'dd/mm/yyyy
        ldate1 = CDate(rsISSH.Fields("ShipDate1"))
        ldate1 = DateAdd("yyyy", 543, ldate1)
        rsERPH.Fields("DLVDAT") = ldate1
    End If
    
    rsERPH.Fields("DLVTIM") = " "
    rsERPH.Fields("DLVDAT_IT") = " "
    rsERPH.Fields("AMTRAT0") = 0
    ldate1 = Empty
    rsERPH.Fields("CMPLDAT") = ldate1
    rsERPH.Fields("DOCSTAT") = "N"
    rsERPH.Fields("USERID") = "ISS"
    rsERPH.Fields("CHGDAT") = Date
    rsERPH.Fields("USERPRN") = " "
    ldate1 = Empty
    rsERPH.Fields("PRNDAT") = ldate1
    rsERPH.Fields("PRNCNT") = 0
    rsERPH.Fields("PRNTIM") = " "
    rsERPH.Fields("AUTHID") = " "
    ldate1 = Empty
    rsERPH.Fields("APPROVE") = ldate1
    rsERPH.Fields("BILLTO") = " "



    '==Maxitem of QuotationTransactionItem =============================
    Set rsISS = New ADODB.Recordset
    
    strsql = "          SELECT MAX(ItemNumber) as maxitem"
    If strRuntype = "QT" Then
    strsql = strsql & " From QuotationTransactionItem"
    Else
    strsql = strsql & " From OrderTransactionItem"
    End If
    strsql = strsql & " where Filename ='" & rsISSH.Fields("Filename") & "'"
    strsql = strsql & " and DocNumber ='" & rsISSH.Fields("DocNumber") & "'"
    
    rsISS.Open strsql, conISS, 1, 1
    rsISS.MoveFirst
    
    Do While Not rsISS.EOF
        rsERPH.Fields("NXTSEQ") = Space(3 - Len(CStr(rsISS.Fields("maxitem")))) & rsISS.Fields("maxitem")
        rsISS.MoveNext
    Loop
    rsISS.Close
    Set rsISS = Nothing
    
    
    rsERPH.Fields("AMOUNT") = rsISSH.Fields("Total_Sum")
    rsERPH.Fields("DISCAMT") = rsISSH.Fields("Totla_HeadDiscount")
    rsERPH.Fields("TOTAL") = rsISSH.Fields("Total_Sum") - rsISSH.Fields("Totla_HeadDiscount")
    
'    If Mid(rsISSH.Fields("DepartmentCode"), 3, 1) = "Y" Then
        rsERPH.Fields("VATAMT") = rsISSH.Fields("Total_Vat")
        rsERPH.Fields("NETAMT") = rsISSH.Fields("Total_Total")
'    Else
'        rsERPH.Fields("VATAMT") = 0
'        rsERPH.Fields("NETAMT") = rsISSH.Fields("Total_Total") - rsISSH.Fields("Total_Vat")
'    End If
    
    rsERPH.Fields("NETVAL") = rsISSH.Fields("Total_Sum") - rsISSH.Fields("Totla_HeadDiscount")



    '==Header DISCOUNT====================================================
    strDiscFormula = ""
    lTotaldisc = 0
    
    If rsISSH.Fields("Totla_HeadDiscount") > 0 Then

        Set rsISSD = New ADODB.Recordset

        strsql = "          SELECT "
        strsql = strsql & "        d.DocNumber , d.DiscountRefNo , d.DiscountSequence "
        strsql = strsql & "      , d.DiscountCode , d.DiscountType , d.DiscountAmount "
        strsql = strsql & "      , d.DiscountUnit , t.DiscountTypeDescription"
        If strRuntype = "QT" Then
        strsql = strsql & " From QuotationTransactionDiscount d, DiscountType t"
        Else
        strsql = strsql & " From OrderTransactionDiscount d, DiscountType t"
        End If
        strsql = strsql & " Where d.Filename ='" & rsISSH.Fields("Filename") & "'"
        strsql = strsql & " and d.DocNumber ='" & rsISSH.Fields("DocNumber") & "'"
        strsql = strsql & " and d.DiscountRefNo = 0"
        strsql = strsql & " and d.DiscountCode = t.DiscountType"
        strsql = strsql & " and t.company = '" & strComcode & "'"
        strsql = strsql & " Order by d.DiscountRefNo, d.DiscountSequence"

        rsISSD.Open strsql, conISS, 1, 1

        If Not rsISSD.EOF Then
            rsISSD.MoveFirst
        End If

        Do While Not rsISSD.EOF

            Select Case rsISSD.Fields("DiscountType")
                'Discount%p
                Case "P"
                    strDiscFormula = rsISSD.Fields("DiscountType")
                    lTotaldisc = lTotaldisc + rsISSD.Fields("DiscountAmount")

                'Discount'A' amt
                Case "A"
                    strDiscFormula = rsISSD.Fields("DiscountType")
                    lTotaldisc = lTotaldisc + rsISSD.Fields("DiscountAmount")
                    
            End Select
            lDiscountSequence = rsISSD.Fields("DiscountSequence")

            rsISSD.MoveNext
        Loop
        rsISSD.Close
        Set rsISSD = Nothing
    End If

    If Len(strDiscFormula) > 0 Then
        'Discount%P or A
        If strDiscFormula = "A" Then
        
            strDiscFormula = Format(lTotaldisc, "#######.00")
            ' Express lock disc 1 step only but iss can create more than 1 step
            ' So i will total many disc to 1 discount
            If lDiscountSequence > 1 Then
                rsERPH.Fields("DISCAMT") = lTotaldisc
                rsERPH.Fields("TOTAL") = rsISSH.Fields("Total_Sum") - rsERPH.Fields("DISCAMT")
                rsERPH.Fields("VATAMT") = rsERPH.Fields("TOTAL") * (gVat_rate / 100)
                rsERPH.Fields("NETAMT") = rsERPH.Fields("TOTAL") + rsERPH.Fields("VATAMT")
                rsERPH.Fields("NETVAL") = rsERPH.Fields("TOTAL")
            End If
            
        ElseIf strDiscFormula = "P" Then
        
            strDiscFormula = lTotaldisc & "%"
            If lDiscountSequence > 1 Then
                rsERPH.Fields("DISCAMT") = rsISSH.Fields("Total_Sum") * (lTotaldisc / 100)
                rsERPH.Fields("TOTAL") = rsISSH.Fields("Total_Sum") - rsERPH.Fields("DISCAMT")
                rsERPH.Fields("VATAMT") = rsERPH.Fields("TOTAL") * (gVat_rate / 100)
                rsERPH.Fields("NETAMT") = rsERPH.Fields("TOTAL") + rsERPH.Fields("VATAMT")
                rsERPH.Fields("NETVAL") = rsERPH.Fields("TOTAL")
            End If
            
        End If
        rsERPH.Fields("DISC") = Space(10 - Len(strDiscFormula)) & strDiscFormula
                
    Else
        'Free or No_discount
        rsERPH.Fields("DISC") = " "
    End If
        

'    Write_Log (": PrepareQTHD Finished. . .")

    Exit Sub
Err_Handler:
    If Len(Trim(strSOlist)) > 0 Then Call SendEmailSOlist
    If Err.Number <> -2147217865 Then
        Call Update_StatusError
    End If
    MsgBox "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ" & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call PrepareQTHD ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix, vbOKOnly
    Write_Log ("Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call PrepareQTHD ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
    Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! กรุณาแจ้งทีม Support iSmartSales", strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ", "Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call PrepareQTHD ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & ", Company=" & rsISSC.Fields("strcom") & ", Comcode=" & strComcode & " ,Doctype=" & strPrefix)
    If rsISSC.Fields("strcom") = "PCK" Then
        conERP.RollbackTrans
    Else
        conERP2.RollbackTrans
    End If
    'log rollback
    Write_Log (": RollbackTrans DocNumber=" & rsISSH.Fields("DocNumber"))
    End
End Sub

Public Sub PrepareQTHDRemark()
On Error GoTo Err_Handler

    Dim lline, linex, countStr, lstart As Integer
    
    Write_Log (": QTHDRemark DocNumber=" & rsISSH.Fields("DocNumber") & " ,Doctype=" & strPrefix)
    
    'Change Max char =48 because fix black color in remark form of express
    
    'Max 50*10_line =500 char
    Set rsERPHR = New ADODB.Recordset
    strsql = "SELECT * FROM ARTRNRM Where 1=2 "
    
    If rsISSC.Fields("strcom") = "PCK" Then
        rsERPHR.Open strsql, conERP, adOpenDynamic, adLockOptimistic
    Else
        rsERPHR.Open strsql, conERP2, adOpenDynamic, adLockOptimistic
    End If
    
    lline = Len(Trim(rsISSH.Fields("DocRemark"))) \ 48
    If Len(Trim(rsISSH.Fields("DocRemark"))) Mod 48 > 0 Then
        lline = lline + 1
    End If
    
    countStr = 0
    
    If strRuntype = "QT" And rsISSC.Fields("strcom") = "PCK" Then
        'QT must start line2 only
        lstart = 2
        lline = lline + 1
        
        rsERPHR.AddNew
        rsERPHR.Fields("DOCNUM") = strSOno
        rsERPHR.Fields("SEQNUM") = "@1"
        rsERPHR.Fields("REMARK") = " "
        rsERPHR.Update
    Else
        'SO Normal
        lstart = 1
    End If
    
    For linex = lstart To lline
    
        If linex > 9 Then Exit For
        rsERPHR.AddNew
        
        rsERPHR.Fields("DOCNUM") = strSOno
        rsERPHR.Fields("SEQNUM") = "@" & linex
        rsERPHR.Fields("REMARK") = Mid(Trim(rsISSH.Fields("DocRemark")), countStr + 1, 48)
        
        rsERPHR.Update
        
        If strRuntype = "QT" And rsISSC.Fields("strcom") = "PCK" Then
            countStr = (linex - 1) * 48
        Else
            countStr = linex * 48
        End If
        
    Next linex
    
    rsERPHR.Close
    Set rsERPHR = Nothing
    
    
'    Write_Log (": QTHDRemark Finished. . .")

    Exit Sub
Err_Handler:
    If Len(Trim(strSOlist)) > 0 Then Call SendEmailSOlist
    If Err.Number <> -2148217865# Then
        Call Update_StatusError
    End If
    MsgBox "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ" & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "Error : " & Err.Number & " " & Err.Description & vbCrLf & " :Call PrepareQTHDRemark ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix, vbOKOnly
    Write_Log ("Error : " & Err.Number & " " & Err.Description & vbCrLf & " :Call PrepareQTHDRemark ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
    Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! กรุณาแจ้งทีม Support iSmartSales", strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ", "Error : " & Err.Number & " " & Err.Description & vbCrLf & " :Call PrepareQTHDRemark ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & ", Company=" & rsISSC.Fields("strcom") & ", Comcode=" & strComcode & " ,Doctype=" & strPrefix)
    If rsISSC.Fields("strcom") = "PCK" Then
        conERP.RollbackTrans
    Else
        conERP2.RollbackTrans
    End If
    'log rollback
    Write_Log (": RollbackTrans DocNumber=" & rsISSH.Fields("DocNumber"))
    End
End Sub

Public Sub PrepareQTDT()
On Error GoTo Err_Handler

    Dim ldate1 As Date
    Dim ltfactor As Currency
    
'    Write_Log (": PrepareQTDT Start. . .")


    '== ISTAB for find tqucod(UnitName) ==============================================
    Set rsERP = New ADODB.Recordset

    strsql = ""
    strsql = strsql & " SELECT typcod"
    strsql = strsql & " From ISTAB"
    strsql = strsql & " Where tabtyp ='20'"
    strsql = strsql & "   AND typdes ='" & rsISSI.Fields("ProductUnit") & "'"

    If rsISSC.Fields("strcom") = "PCK" Then
        rsERP.Open strsql, conERP, 1, 1
    Else
        rsERP.Open strsql, conERP2, 1, 1
    End If

    If Not rsERP.EOF Then
        rsERP.MoveFirst
    Else
        If Len(Trim(strSOlist)) > 0 Then Call SendEmailSOlist
        Call Update_StatusError
        MsgBox "ข้อผิดพลาด: ไม่พบข้อมูลหน่วยสินค้า Unit=" & rsISSI.Fields("ProductUnit") & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call PrepareQTDT ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Doctype=" & strPrefix & " ,Company=" & rsISSC.Fields("strcom") & " ,ProductCode=" & rsISSI.Fields("ProductCode"), vbOKOnly
        Write_Log ("Error Notfound Unit=" & rsISSI.Fields("ProductUnit") & vbCrLf & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call PrepareQTDT ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Doctype=" & strPrefix & " ,Company=" & rsISSC.Fields("strcom") & " ,ProductCode=" & rsISSI.Fields("ProductCode"))
        Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! ไม่พบข้อมูลหน่วยสินค้า " & rsISSI.Fields("ProductUnit") & " ในบริษัท " & rsISSC.Fields("strcom"), strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ข้อผิดพลาด:ไม่พบข้อมูลหน่วยสินค้า Unit=" & rsISSI.Fields("ProductUnit") & " ในบริษัท " & rsISSC.Fields("strcom"), "Error : " & Err.Number & " " & Err.Description & " :Call PrepareQTDT ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Comcode=" & strComcode & " ,ProductCode=" & rsISSI.Fields("ProductCode") & " ,Doctype=" & strPrefix)
        If rsISSC.Fields("strcom") = "PCK" Then
            conERP.RollbackTrans
        Else
            conERP2.RollbackTrans
        End If
        'log rollback
        Write_Log (": RollbackTrans DocNumber=" & rsISSH.Fields("DocNumber") & " ProductCode=" & rsISSI.Fields("ProductCode") & " ,Doctype=" & strPrefix)
        End
    End If
    
    Do While Not rsERP.EOF

        rsERPD.Fields("TQUCOD") = Mid(rsERP.Fields("typcod"), 1, 2)
        rsERP.MoveNext
    Loop
    rsERP.Close
    Set rsERP = Nothing


    '== STMAS ==============================================
    Set rsERP = New ADODB.Recordset
    
    strsql = ""
    strsql = strsql & " SELECT stkdes, qucod, cqucod, cfactor"
    strsql = strsql & " From STMAS"
    strsql = strsql & " Where stkcod = '" & Mid(rsISSI.Fields("ProductCode"), 4) & "'"

    If rsISSC.Fields("strcom") = "PCK" Then
        rsERP.Open strsql, conERP, 1, 1
    Else
        rsERP.Open strsql, conERP2, 1, 1
    End If
    
    If Not rsERP.EOF Then
        rsERP.MoveFirst
    Else
        If Len(Trim(strSOlist)) > 0 Then Call SendEmailSOlist
        Call Update_StatusError
        MsgBox "ข้อผิดพลาด: ไม่พบข้อมูลสินค้า " & rsISSI.Fields("ProductCode") & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call PrepareQTDT ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix, vbOKOnly
        Write_Log ("Error ไม่พบข้อมูลสินค้า " & rsISSI.Fields("ProductCode") & vbCrLf & vbCrLf & "Error : " & Err.Number & " " & Err.Description & " :Call PrepareQTDT ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
        Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! ไม่พบข้อมูลสินค้า " & rsISSI.Fields("ProductCode") & " ในบริษัท " & rsISSC.Fields("strcom"), strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ข้อผิดพลาด:ไม่พบข้อมูลสินค้า " & rsISSI.Fields("ProductCode") & " ในบริษัท " & rsISSC.Fields("strcom"), "Error : " & Err.Number & " " & Err.Description & " :Call PrepareQTDT ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,CustCode=" & rsISSH.Fields("CustomerCode") & " ,Doctype=" & strPrefix)
        If rsISSC.Fields("strcom") = "PCK" Then
            conERP.RollbackTrans
        Else
            conERP2.RollbackTrans
        End If
        'log rollback
        Write_Log (": RollbackTrans DocNumber=" & rsISSH.Fields("DocNumber") & " ProductCode=" & rsISSI.Fields("ProductCode") & " ,Doctype=" & strPrefix)
        End
    End If
    
    Do While Not rsERP.EOF
    
        rsERPD.Fields("STKDES") = rsERP.Fields("stkdes")
        
        If rsERPD.Fields("TQUCOD") = rsERP.Fields("qucod") Then
        
            ltfactor = 1
            rsERPD.Fields("TFACTOR") = 1

        Else 'If rsERPD.Fields("TQUCOD") = rsERP.Fields("cqucod") Then
        
            ltfactor = rsERP.Fields("cfactor")
            rsERPD.Fields("TFACTOR") = rsERP.Fields("cfactor")
            
        End If
        
        rsERP.MoveNext
    Loop
    rsERP.Close
    Set rsERP = Nothing
    
    
    If strRuntype = "QT" Then
        '=Quotation do not use Location
        rsERPD.Fields("LOCCOD") = " "
    
    Else   'SO
        '==Location =plant
        If strPrefix = "SB" Then
            rsERPD.Fields("LOCCOD") = "01" 'PCK
        ElseIf strPrefix = "SO" Then
            rsERPD.Fields("LOCCOD") = "04" 'WAT
        End If
        
        '==Update Totres+Quantity_SO
        strsql = "          Update STMAS"
        strsql = strsql & " SET TOTRES = TOTRES+" & (rsISSI.Fields("ProductQuan") * ltfactor)
        strsql = strsql & " Where STKCOD = '" & Mid(rsISSI.Fields("ProductCode"), 4) & "'"
        If rsISSC.Fields("strcom") = "PCK" Then
            conERP.Execute strsql
        Else
            conERP2.Execute strsql
        End If
    
        '==Update Locres+Quantity_SO
        strsql = "          Update STLOC"
        strsql = strsql & " SET LOCRES = LOCRES+" & (rsISSI.Fields("ProductQuan") * ltfactor)
        strsql = strsql & " Where STKCOD = '" & Mid(rsISSI.Fields("ProductCode"), 4) & "'"
        strsql = strsql & "   And LOCCOD = '" & rsERPD.Fields("LOCCOD") & "'"
        If rsISSC.Fields("strcom") = "PCK" Then
            conERP.Execute strsql
        Else
            conERP2.Execute strsql
        End If
        
    End If


    '== ISTAB ==============================================
'    Set rsERP = New ADODB.Recordset
'
'    strsql = ""
'    strsql = strsql & " SELECT typcod"
'    strsql = strsql & " From ISTAB"
'    strsql = strsql & " Where tabtyp ='21'"
'    strsql = strsql & "   AND typcod ='01'"
'
'    rsERP.Open strsql, conERP, 1, 1
'
'    rsERP.MoveFirst
'
'    Do While Not rsERP.EOF
'
'        rsERPD.Fields("LOCCOD") = rsERP.Fields("typcod")
'        rsERP.MoveNext
'    Loop
'    rsERP.Close
'    Set rsERP = Nothing

    

    '==SODT Other Field===================================================
    
    rsERPD.Fields("SONUM") = strSOno
    rsERPD.Fields("SEQNUM") = Space(3 - Len(CStr(rsISSI.Fields("ItemNumber")))) & rsISSI.Fields("ItemNumber")
    
    'dd/mm/yyyy
    ldate1 = CDate(rsISSH.Fields("DocDate1"))
    ldate1 = DateAdd("yyyy", 543, ldate1)
    rsERPD.Fields("SODAT") = ldate1
    
    If strRuntype = "QT" Then
        
        rsERPD.Fields("SORECTYP") = "5"
        'dd/mm/yyyy
        ldate1 = CDate(rsISSH.Fields("DocDate1"))
        ldate1 = DateAdd("yyyy", 543, ldate1)
        ldate1 = DateAdd("d", intpaytrm, ldate1)
        rsERPD.Fields("DLVDAT") = ldate1
        
    Else   'SO
        rsERPD.Fields("SORECTYP") = "0"
        'dd/mm/yyyy
        ldate1 = CDate(rsISSH.Fields("ShipDate1"))
        ldate1 = DateAdd("yyyy", 543, ldate1)
        rsERPD.Fields("DLVDAT") = ldate1
    End If
    
    rsERPD.Fields("CUSCOD") = Mid(rsISSH.Fields("CustomerCode"), 4)
    rsERPD.Fields("STKCOD") = Mid(rsISSI.Fields("ProductCode"), 4)
    rsERPD.Fields("DEPCOD") = " "
    rsERPD.Fields("VATCOD") = " "
    
    If rsISSI.Fields("ProductPrice") = 0 Then
        rsERPD.Fields("FREE") = "Y"
    Else
        'If rsISSI.Fields("ProductPrice") > 0 Then
        rsERPD.Fields("FREE") = " "
    End If
    
    rsERPD.Fields("ORDQTY") = rsISSI.Fields("ProductQuan")
    rsERPD.Fields("CANCELQTY") = 0
    rsERPD.Fields("CANCELTYP") = " "
    ldate1 = Empty
    rsERPD.Fields("CANCELDAT") = ldate1
    rsERPD.Fields("REMQTY") = rsISSI.Fields("ProductQuan")
    rsERPD.Fields("UNITPR") = rsISSI.Fields("ProductPrice")
    If strRuntype = "QT" Then
        rsERPD.Fields("PACKING") = rsISSH.Fields("DocNumber")
    Else
        rsERPD.Fields("PACKING") = " "
    End If
        
        
'    Write_Log (": PrepareQTDT Finished. . .")

    Exit Sub
Err_Handler:
    If Len(Trim(strSOlist)) > 0 Then Call SendEmailSOlist
    If Err.Number <> -2147217865 Then
        Call Update_StatusError
    End If
    MsgBox "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ" & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call PrepareQTDT ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Item:" & rsISSI.Fields("ItemNumber") & " ,Product:" & rsISSI.Fields("ProductCode") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix, vbOKOnly
    Write_Log ("Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call PrepareQTDT ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Item:" & rsISSI.Fields("ItemNumber") & " ,Product:" & rsISSI.Fields("ProductCode") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
    Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! กรุณาแจ้งทีม Support iSmartSales", strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ", "Error : " & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & " :Call PrepareQTDT ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Item:" & rsISSI.Fields("ItemNumber") & " ,Product:" & rsISSI.Fields("ProductCode") & " ,Comcode=" & strComcode & " ,Doctype=" & strPrefix)
    If rsISSC.Fields("strcom") = "PCK" Then
        conERP.RollbackTrans
    Else
        conERP2.RollbackTrans
    End If
    'log rollback
    Write_Log (": RollbackTrans DocNumber=" & rsISSH.Fields("DocNumber") & " ProductCode=" & rsISSI.Fields("ProductCode") & " ,Doctype=" & strPrefix)
    End
End Sub

Public Sub GetQTno_id()
On Error GoTo Err_Handler
    
    Dim lintSOid As Long
    
    '=================== VAT ======================================================================
    '    Write_Log (": GetQTno_id Start. . .")
        
    '*Lock Record
    strsql = "          SET EXCLUSIVE ON;"
    strsql = strsql & " USE ISRUN;"
    
    If rsISSC.Fields("strcom") = "PCK" Then
        conERP.Execute strsql
    Else
        conERP2.Execute strsql
    End If

    '===Get NextSOno==========================
    Set rsERP = New ADODB.Recordset
    
    strsql = "         SELECT doctyp, prefix, docnum AS docnum2"
    strsql = strsql & "  From ISRUN"
    
    If strRuntype = "QT" Then
    
        If rsISSC.Fields("strcom") = "PCK" Then
            strsql = strsql & " WHERE doctyp = 'QT'"
            strsql = strsql & "   AND prefix = 'QB'"
            
            rsERP.Open strsql, conERP, 1, 3
        Else
            strsql = strsql & " WHERE doctyp = 'QT'"
            strsql = strsql & "   AND prefix = 'QT'"
            
            rsERP.Open strsql, conERP2, 1, 3
        End If
        
    Else   'SO
        If rsISSC.Fields("strcom") = "PCK" Then
            strsql = strsql & " WHERE doctyp = 'SO'"
            strsql = strsql & "   AND prefix = 'SB'"
            
            rsERP.Open strsql, conERP, 1, 3
        Else
            strsql = strsql & " WHERE doctyp = 'SO'"
            strsql = strsql & "   AND prefix = 'SO'"
            
            rsERP.Open strsql, conERP2, 1, 3
        End If
    End If
    
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        strSOno = Trim(rsERP.Fields("prefix")) & Trim(rsERP.Fields("docnum2"))
        lintSOid = rsERP.Fields("docnum2")
        strDoctype = rsERP.Fields("doctyp")
        strPrefix = rsERP.Fields("prefix")
        
        rsERP.MoveNext
    Loop
    rsERP.Close
    Set rsERP = Nothing
    
    
    'Format PCK='QB1800350'  WAT='QT6100150'
    strsql = "          Update ISRUN"
    strsql = strsql & " Set docnum = '" & Format(lintSOid + 1, "0000000") & "'"
    strsql = strsql & " Where doctyp ='" & strDoctype & "'"
    strsql = strsql & "   AND prefix ='" & strPrefix & "'"
    
    If rsISSC.Fields("strcom") = "PCK" Then
        conERP.Execute strsql
    Else
        conERP2.Execute strsql
    End If

'    Write_Log (": GetQTno_id Finished. . .")
    

    Exit Sub
Err_Handler:
    If Len(Trim(strSOlist)) > 0 Then Call SendEmailSOlist
    If Err.Number <> -2147217865 Then
        Call Update_StatusError
    End If
    MsgBox "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ" & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "Error : " & Err.Number & " " & Err.Description & vbCrLf & " :Call GetQTno_id ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix, vbOKOnly
    Write_Log ("Error : " & Err.Number & " " & Err.Description & vbCrLf & " :Call GetQTno_id ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & " ,Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
    Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! ในการเชื่อมต่อระบบบัญชี", strEngdoc & ":" & rsISSH.Fields("DocNumber") & "   พนักงานขาย:" & rsISSH.Fields("SalesmanCode") & vbCrLf & "ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ", "Error : " & Err.Number & " " & Err.Description & vbCrLf & " :Call GetQTno_id ,Filename:" & rsISSH.Fields("Filename") & " ,DocNumber:" & rsISSH.Fields("DocNumber") & ", Company=" & rsISSC.Fields("strcom") & " ,Doctype=" & strPrefix)
    If rsISSC.Fields("strcom") = "PCK" Then
        conERP.RollbackTrans
    Else
        conERP2.RollbackTrans
    End If
    'log rollback
    Write_Log (": RollbackTrans DocNumber=" & rsISSH.Fields("DocNumber"))
    End
End Sub

Public Sub Update_StatusError()

    'UPdate status ='E' Error
    If strRuntype = "QT" Then
        strsql = "          Update QuotationTransactionLoad"
        strsql = strsql & " set status = 'E'"
    Else
        strsql = "          Update OrderTransactionLoad"
        strsql = strsql & " set SOstatus = 'E'"
    End If
    strsql = strsql & " where Company ='" & strComcode & "'"
    strsql = strsql & " and Filename ='" & rsISSH.Fields("Filename") & "'"
    strsql = strsql & " and DocNumber ='" & rsISSH.Fields("DocNumber") & "'"
    
    conISS.Execute strsql

End Sub

Public Sub Update_StatusErrorOf(ByVal lDoctype As String)

    'UPdate status ='E' Error
    
    If lDoctype = "QUATA" Then
'        strsql = "      UPDATE QuotationTransactionLoad"
    ElseIf lDoctype = "RETURN" Then
        strsql = "      UPDATE ReturnTransactionLoad"
    ElseIf lDoctype = "STOCK" Then
        strsql = "      UPDATE StockTransactionLoad"
    End If
    strsql = strsql & " Set status = 'E'"
    If lDoctype <> "STOCK" Then
        strsql = strsql & " Where Company ='" & strComcode & "'"
    Else
        strsql = strsql & " Where 1=1"
    End If
    strsql = strsql & " And Filename ='" & rsISSH.Fields("Filename") & "'"
    strsql = strsql & " And DocNumber ='" & rsISSH.Fields("DocNumber") & "'"
    
    conISS.Execute strsql

End Sub

Public Sub SendEmailSOlist()

    Dim retval As String
    Dim lstrEmailSubject, lstrEmailbody As String
    Dim lstrEmailTo, lstrEmailCC As String
'    Dim mypos As Integer

'    mypos = InStr(strCompany, "0")
    lstrEmailSubject = "ISS-PCK: การโหลด" & strThaidoc & "เรียบร้อยแล้ว"
 
    lstrEmailbody = "รายการ" & strThaidoc & "ที่โหลดแล้ว มีดังนี้" & vbCrLf & strSOlist


    Select Case UCase(rsISSA.Fields("Username"))
        Case "PCK"
           lstrEmailTo = strEmailAdmin1  'ar2@jeedjard.com
    End Select
    
    lstrEmailCC = strEmailTo   '"thamarat.v@iconnectmkt.com;"
    
    retval = SendMail(Trim$(lstrEmailTo), Trim$(lstrEmailCC), _
                                    Trim$(lstrEmailSubject), _
                                    Trim$("iSmartSale Support Team") & "<" & Trim$("support@ismartsale.com") & ">", _
                                    Trim$(lstrEmailbody), _
                                    Trim$("smtp.ismartsale.com"), _
                                    587, _
                                    Trim$("support@ismartsale.com"), _
                                    Trim$("Sup2@11"), _
                                    Trim$(""), _
                                    0)

End Sub

Public Sub SendEmailOf(ByVal lDoctype As String)
On Error GoTo Err_Handler

    Dim retval As String
    Dim lstrEmailSubject, lstrEmailbody As String
    Dim lstrEmailTo, lstrEmailCC As String
    Dim mypos As Integer
    Dim lstrSOlist As String
    

    Set rsISS = New ADODB.Recordset
    
    strsql = "         SELECT SalesmanCode ,  Cast(DocDate AS datetime) as DocDate1 "
    strsql = strsql & "     , DocNumber , Filename "
    
    If lDoctype = "QUATA" Then
'        strsql = strsql & " From QuotationTransactionLoad"
    ElseIf lDoctype = "RETURN" Then
        strsql = strsql & " From ReturnTransactionLoad"
    ElseIf lDoctype = "STOCK" Then
        strsql = strsql & " From StockTransactionLoad"
    End If
    strsql = strsql & " Where status IS NULL"
    If lDoctype <> "STOCK" Then
        strsql = strsql & " AND Company = '" & strComcode & "'"
    End If
    
    strsql = strsql & " and DocCreateTime >= Convert(char(8),dateadd(day,-30,GETDATE()), 112)+'000000'"
'    strsql = strsql & " and DocCreateTime >= '20171030000000'"
'    strsql = strsql & " and SalesmanCode in (" & strSlmlist & ")"
    strsql = strsql & " Order by DocDate,DocNumber"
    
    rsISS.Open strsql, conISS, 1, 3
    
    lstrSOlist = ""
    If Not rsISS.EOF Then
    
        rsISS.MoveFirst
    Else
    
'        If lDoctype = "QUATA" Then
'            Write_Log (": SendEmail Not found Quatation. . .")
'        ElseIf lDoctype = "RETURN" Then
'            Write_Log (": SendEmail Not found Return. . .")
'        ElseIf lDoctype = "STOCK" Then
'            Write_Log (": SendEmail Not found Stock. . .")
'        End If
        
        rsISS.Close
        Set rsISS = Nothing
        Exit Sub
    End If
    
    Do While Not rsISS.EOF
    
        lstrSOlist = lstrSOlist & "'" & rsISS.Fields("DocNumber") & "',"
        
        'UPdate status ='C' complete
        If lDoctype = "QUATA" Then
'            strsql = "      UPDATE QuotationTransactionLoad"
        ElseIf lDoctype = "RETURN" Then
            strsql = "      UPDATE ReturnTransactionLoad"
        ElseIf lDoctype = "STOCK" Then
            strsql = "      UPDATE StockTransactionLoad"
        End If
        
        strsql = strsql & " SET status = 'C'"
        If lDoctype <> "STOCK" Then
            strsql = strsql & " Where Company ='" & strComcode & "'"
        Else
            strsql = strsql & " Where 1=1"
        End If
        strsql = strsql & " And Filename ='" & rsISS.Fields("Filename") & "'"
        strsql = strsql & " And DocNumber ='" & rsISS.Fields("DocNumber") & "'"
        
        conISS.Execute strsql
        
        rsISS.MoveNext
    Loop
    lstrSOlist = Mid(lstrSOlist, 1, Len(lstrSOlist) - 1)
    
    rsISS.Close
    Set rsISS = Nothing


    '---------SendEmail Start.--------------------------'

'    mypos = InStr(strCompany, "0")
    If lDoctype = "QUATA" Then
'        lstrEmailSubject = "ISS-PCK: ISS2PCK AOG Quataion Company:" & strComcode
'        lstrEmailbody = rsISSA.Fields("Username") & " has Quatation = " & lstrSOlist

    ElseIf lDoctype = "RETURN" Then
        lstrEmailSubject = "ISS-PCK: ISS2PCK AOG Return Company:" & strComcode
        lstrEmailbody = rsISSA.Fields("Username") & " has Return = " & lstrSOlist

    ElseIf lDoctype = "STOCK" Then
        lstrEmailSubject = "ISS-PCK: ISS2PCK AOG Stock Company:" & strComcode
        lstrEmailbody = rsISSA.Fields("Username") & " has Stock = " & lstrSOlist
        
    End If
    
    Select Case UCase(rsISSA.Fields("Username"))
        Case "PCK"
           lstrEmailTo = strEmailAdmin1  'wanphilast@force.co.th
    End Select
    
'    lstrEmailTo = strEmailTo   '"thamarat.v@iconnectmkt.com;"
    lstrEmailCC = strEmailTo   '"thamarat.v@iconnectmkt.com;"
    
    retval = SendMail(Trim$(lstrEmailTo), Trim$(lstrEmailCC), _
                                    Trim$(lstrEmailSubject), _
                                    Trim$("iSmartSale Support Team") & "<" & Trim$("support@ismartsale.com") & ">", _
                                    Trim$(lstrEmailbody), _
                                    Trim$("smtp.ismartsale.com"), _
                                    587, _
                                    Trim$("support@ismartsale.com"), _
                                    Trim$("Sup2@11"), _
                                    Trim$(""), _
                                    0)
                                    
    Write_Log (": Call SendEmailof " & lDoctype & " ,strcompany:" & strCompany & " ,username:" & rsISSA.Fields("Username") & " ,lstrEmailTo:" & lstrEmailTo & " ,lstrSOlist:" & lstrSOlist & " ,strComcode:" & strComcode)
    
    '---------SendEmail Finished.--------------------------'

    Exit Sub
Err_Handler:
    Call Update_StatusErrorOf(lDoctype)
    Write_Log ("Error : " & Err.Number & " " & Err.Description & " :Call SendEmailof " & lDoctype & " ,Filename=" & rsISS.Fields("Filename") & " ,DocNumber:" & rsISS.Fields("DocNumber"))
    Call SendEmailErr("ISS-PCK: พบข้อผิดพลาด!!! กรุณาแจ้งทีม Support iSmartSales", "Quotation:" & rsISS.Fields("DocNumber") & "   พนักงานขาย:" & rsISS.Fields("SalesmanCode") & vbCrLf & " ตรวจพบข้อผิดพลาด กรุณาแจ้งทีม Support iSmartSales เพื่อดำเนินการ", "Error : " & Err.Number & " " & Err.Description & " :Call SendEmailof " & lDoctype & " ,Filename=" & rsISS.Fields("Filename") & " ,DocNumber:" & rsISS.Fields("DocNumber"))
    End
End Sub

Public Sub SendEmailErr(ByVal MsgSubject As String, ByVal MsgErr As String, ByVal MsgTec As String)

    Dim retval As String
    Dim lstrEmailSubject, lstrEmailbody As String
    Dim lstrEmailTo, lstrEmailCC As String
    

    lstrEmailSubject = MsgSubject
    lstrEmailbody = MsgErr & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "   Database: " & strISSDBS & ", ERP Database: " & strERPPath & vbCrLf & "   " & MsgTec

'    Select Case UCase(rsISSA.Fields("Username"))
'        Case "PCK"
           lstrEmailTo = strEmailAdmin1  'ar2@jeedjard.com

'    End Select
    
'    lstrEmailTo = strEmailTo                        'thamarat.v@iconnectmkt.com
    lstrEmailCC = strEmailTo & ";" & strEmailCC    '"thamarat.v@iconnectmkt.com;setthapong.k@iconnectmkt.com"
    
    retval = SendMail(Trim$(lstrEmailTo), Trim$(lstrEmailCC), _
                                    Trim$(lstrEmailSubject), _
                                    Trim$("iSmartSale Support Team") & "<" & Trim$("support@ismartsale.com") & ">", _
                                    Trim$(lstrEmailbody), _
                                    Trim$("smtp.ismartsale.com"), _
                                    587, _
                                    Trim$("support@ismartsale.com"), _
                                    Trim$("Sup2@11"), _
                                    Trim$(""), _
                                    0)
    
End Sub

Public Sub SendEmailErr2(ByVal MsgSubject As String, ByVal MsgErr As String, ByVal MsgTec As String)

    Dim retval As String
    Dim lstrEmailSubject, lstrEmailbody As String
    Dim lstrEmailTo, lstrEmailCC As String
    

    lstrEmailSubject = MsgSubject
    lstrEmailbody = MsgErr & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "   Database: " & strISSDBS & ", ERP Database: " & strERPPath & vbCrLf & "   " & MsgTec

    lstrEmailTo = strEmailAdmin1  'wanphilast@force.co.th
    
    lstrEmailCC = strEmailTo & ";" & strEmailCC    '"thamarat.v@iconnectmkt.com;setthapong.k@iconnectmkt.com"
    
    retval = SendMail(Trim$(lstrEmailTo), Trim$(lstrEmailCC), _
                                    Trim$(lstrEmailSubject), _
                                    Trim$("iSmartSale Support Team") & "<" & Trim$("support@ismartsale.com") & ">", _
                                    Trim$(lstrEmailbody), _
                                    Trim$("smtp.ismartsale.com"), _
                                    587, _
                                    Trim$("support@ismartsale.com"), _
                                    Trim$("Sup2@11"), _
                                    Trim$(""), _
                                    0)
    
End Sub

Public Sub SendEmailErrSlm(ByVal MsgSubject As String, ByVal MsgErr As String, ByVal MsgTec As String)

    Dim retval As String
    Dim lstrEmailSubject, lstrEmailbody As String
    Dim lstrEmailTo, lstrEmailCC As String
    

    lstrEmailSubject = MsgSubject
    lstrEmailbody = MsgErr & vbCrLf & vbCrLf & "ข้อมูลทางเทคนิค" & vbCrLf & "   Database: " & strISSDBS & ", ERP Database: " & strERPPath & vbCrLf & "   " & MsgTec
 
    lstrEmailTo = strEmailAdmin1 & ";"              'wanphilast@force.co.th
    lstrEmailCC = strEmailTo & ";" & strEmailCC     '"thamarat.v@iconnectmkt.com;setthapong.k@iconnectmkt.com"
    
    retval = SendMail(Trim$(lstrEmailTo), Trim$(lstrEmailCC), _
                                    Trim$(lstrEmailSubject), _
                                    Trim$("iSmartSale Support Team") & "<" & Trim$("support@ismartsale.com") & ">", _
                                    Trim$(lstrEmailbody), _
                                    Trim$("smtp.ismartsale.com"), _
                                    587, _
                                    Trim$("support@ismartsale.com"), _
                                    Trim$("Sup2@11"), _
                                    Trim$(""), _
                                    0)
    
End Sub

Public Sub Write_Log(pstrLogMsg As String)

    Open strWorkingDir & "\ISS2PCK_Log.txt" For Append As #fn
    Write #fn, Now & pstrLogMsg
    Close #fn
    
End Sub

Public Function SendMail(sTo As String, sCC As String, sSubject As String, sFrom As String, _
    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
    sSmtpUser As String, sSmtpPword As String, _
    sFilePath As String, bSmtpSSL As Boolean) As String
      
    On Error GoTo SendMail_Error:
    Dim lobj_cdomsg      As CDO.Message
    Set lobj_cdomsg = New CDO.Message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.BodyPart.Charset = "utf-8"
    lobj_cdomsg.To = sTo
    lobj_cdomsg.CC = sCC
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.Subject = sSubject
    lobj_cdomsg.TextBody = sBody
    If Trim$(sFilePath) <> vbNullString Then
        lobj_cdomsg.AddAttachment (sFilePath)
    End If
    lobj_cdomsg.Send
    Set lobj_cdomsg = Nothing
    SendMail = "ok"
    Exit Function
          
SendMail_Error:
    SendMail = Err.Description
End Function

