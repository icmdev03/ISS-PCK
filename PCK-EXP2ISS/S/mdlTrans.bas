Attribute VB_Name = "mdlTrans"
Option Explicit
 
Global conERP As New ADODB.Connection
Global conERP1 As New ADODB.Connection
Global conISS As ADODB.Connection
Global rsISS As ADODB.Recordset
Global rsERP As ADODB.Recordset
Global rsERP1 As ADODB.Recordset
Global strERPPath, strERP1Path, Mgrcode, Mgrcode1 As String
Global strEmailTo, strEmailSubject As String
Global strISSServer, strISSUser, strISSPWD, strISSDBS As String
Global strSQL, strERPPWD, strWorkingDir As String
Global mCmd As ADODB.Command
Global mCmd1 As ADODB.Command
Global idx, fn As Integer
Global intCurrMonth, intCurrYear, intLastYear, intLastMonth, intNoOfRow  As Integer

Sub Main()

    Dim ldbISS As Database
    Dim ldbISS2 As Database
    Dim strDBS As String
    Dim retval As String
    Dim strISSConnectionString As String
    Dim dteSalesLoadingDate As Date
    
    fn = FreeFile
    GetiniFile


    'Initialize all Related Database
    strISSServer = "203.151.94.173"
    strISSUser = "iconnectmkt"
    strISSPWD = "tkmtcennoci2@!!"
'    'strISSDBS = "ISSMASTER"
'    strERPPWD = ""
'
'    strISSServer = "NB-H2o\SQLEXPRESS"
'    strISSUser = "iss_admin"
'    strISSPWD = "Icm@2011"
'
    'ISS Database Connection
    strISSConnectionString = "Driver={SQL Server};Server=" + strISSServer + ";Database=" + strISSDBS + ";Uid=" + strISSUser + ";Pwd=" + strISSPWD + ";"
    Set conISS = New ADODB.Connection
    conISS.ConnectionString = strISSConnectionString
    conISS.Open
    
    
    conERP.Provider = "VFPOLEDB.1"
    conERP.Properties("Data Source") = strERPPath
    'conERP.Properties("Jet OLEDB:Database Password") = strERPPWD
    conERP.Properties("Collating Sequence") = "THAI"
    conERP.Open
    
    conERP1.Provider = "VFPOLEDB.1"
    conERP1.Properties("Data Source") = strERP1Path
    'conERP1.Properties("Jet OLEDB:Database Password") = strERPPWD
    conERP1.Properties("Collating Sequence") = "THAI"
    conERP1.Open
    
    'if last day of the month then recreate last month data
    dteSalesLoadingDate = Date - 1
    intCurrMonth = Month(dteSalesLoadingDate)
    intCurrYear = Year(dteSalesLoadingDate)
    
    intLastMonth = intCurrMonth - 1
    If intLastMonth = 0 Then
        intLastMonth = 12
        intLastYear = intCurrYear - 1
    Else
        intLastYear = intCurrYear
    End If
    
'    Call ARReport
'    Call AvailableStockTransaction
'    Call Customer
'    Call CustomerSalesman
'    Call CustomerPricing
    Call Daily_SalesSummaries
'    Call GroupPricing
'    Call SalesAnalysis
    Call SalesHistory
    Call SalesSummaries
'    Call SingleMaster
    Call SalesAnalysis
    
    retval = SendMail(Trim$(strEmailTo), _
                                        Trim$(strEmailSubject), _
                                        Trim$("iSmartSale Support Team") & "<" & Trim$("support@ismartsales.com") & ">", _
                                        Trim$("Database: " + strISSDBS + ", WorkingDir: " + strWorkingDir + ", ERPDir: " + strERPPath), _
                                        Trim$("smtp.ismartsales.com"), _
                                        587, _
                                        Trim$("support@ismartsales.com"), _
                                        Trim$("Sup2@11"), _
                                        Trim$(""), _
                                        0)
    
    Write_Log (": Program Run Successfully. . .")

End Sub
Public Sub GetiniFile()
Dim strLine As String

    
    Open "C:\WINDOWS\system32\PCK2ISS.ini" For Input As #1
    Do While Not EOF(1) ' Check for end of file.
        Input #1, strLine  ' Read data.
           If Mid$(strLine, 1, 11) = "WorkingDir=" Then
              strWorkingDir = Mid$(strLine, 12, Len(Trim(strLine)) - 11)
           End If
           
            If Mid$(strLine, 1, 7) = "ISSDBS=" Then
              strISSDBS = Mid$(strLine, 8, Len(Trim(strLine)) - 7)
           End If

           If Mid$(strLine, 1, 8) = "ERP1Dir=" Then
              strERPPath = Mid$(strLine, 9, Len(Trim(strLine)) - 8)
           End If
           
           If Mid$(strLine, 1, 8) = "ERP2Dir=" Then
              strERP1Path = Mid$(strLine, 9, Len(Trim(strLine)) - 8)
           End If
           
           If Mid$(strLine, 1, 8) = "EmailTo=" Then
              strEmailTo = Mid$(strLine, 9, Len(Trim(strLine)) - 8)
           End If
           
           If Mid$(strLine, 1, 13) = "EmailSubject=" Then
              strEmailSubject = Mid$(strLine, 14, Len(Trim(strLine)) - 13)
           End If
           
            If Mid$(strLine, 1, 8) = "Mgrcode=" Then
              Mgrcode = Mid$(strLine, 9, Len(Trim(strLine)) - 8)
            End If
            
            If Mid$(strLine, 1, 9) = "Mgrcode1=" Then
              Mgrcode1 = Mid$(strLine, 10, Len(Trim(strLine)) - 8)
            End If
    Loop
    Close #1
        
    Write_Log (": Program Start. . .")
    Write_Log (": Database: " + strISSDBS)
    Write_Log (": WorkingDir: " + strWorkingDir)
    Write_Log (": ERPDir: " + strERPPath)

End Sub


Public Sub ARReport()
    Dim intRecordAffected As Integer

    Write_Log (": ARReport Start. . .")

    'Clear Table before Insert new records
    conISS.Execute "DELETE FROM ARReport WHERE Company = '30'"

    'PCK
        Set rsISS = New ADODB.Recordset
        rsISS.Open "SELECT * FROM ARReport ", conISS, adOpenDynamic, adLockOptimistic
    
    
        Set rsERP = New ADODB.Recordset
        Set mCmd = New ADODB.Command
        mCmd.ActiveConnection = conERP1
    
        'Select invoice remain amount greater than zero
        mCmd.CommandText = " SELECT DISTINCT '30'AS Company, ARTRN.slmcod AS SalesmanCode, 'PCK'+ARTRN.cuscod AS CustomerCode, " _
                                                            + " ARTRN.docnum AS BillNo, DTOS(ARTRN.docdat) AS BillDate, ARRCPIT.rcpnum AS TaxInvNo, " _
                                                            + " DTOS(ARTRN.duedat) AS DueDate_Txt, " _
                                                            + " ARTRN.remamt AS BalanceAmount, 'Not Due' AS BillStatus, '' AS PostDateCheque, " _
                                                            + " '' AS PostDateChequeDate, '' AS PostDateChequeAmount, 'THB' " _
                                                            + " From " + strERPPath + "\ARTRN LEFT OUTER JOIN " + strERPPath + "\ARRCPIT ON ARTRN.docnum = ARRCPIT.docnum " _
                                                                         + " LEFT OUTER JOIN " + strERPPath + "\BKTRN ON BKTRN.voucher = ARRCPIT.rcpnum " _
                                                            + " WHERE ARTRN.slmcod <> ' ' " _
                                                            + " AND ARTRN.remamt <> 0 " _
                                                            + " AND SUBSTR(ARTRN.docnum, 1, 2) IN ('IB', 'IS') " _
                                                            + " UNION " _
                                                            + " SELECT DISTINCT '30'AS Company, ARTRN.slmcod AS SalesmanCode, 'PCK'+ARTRN.cuscod AS CustomerCode, " _
                                                        + " ARTRN.docnum AS BillNo, DTOS(ARTRN.docdat) AS BillDate, ARRCPIT.rcpnum AS TaxInvNo, " _
                                                        + " DTOS(ARTRN.duedat) AS DueDate_Txt, " _
                                                        + " ARTRN.remamt AS BalanceAmount, 'Not Due' AS BillStatus, '' AS PostDateCheque, " _
                                                        + " '' AS PostDateChequeDate, '' AS PostDateChequeAmount, 'THB' " _
                                                        + " From " + strERP1Path + "\ARTRN LEFT OUTER JOIN " + strERP1Path + "\ARRCPIT ON ARTRN.docnum = ARRCPIT.docnum " _
                                                                     + " LEFT OUTER JOIN " + strERP1Path + "\BKTRN ON BKTRN.voucher = ARRCPIT.rcpnum " _
                                                        + " WHERE ARTRN.slmcod <> ' ' " _
                                                        + " AND ARTRN.remamt <> 0 " _
                                                        + " AND SUBSTR(ARTRN.docnum, 1, 2) IN ('IV') "

        mCmd.CommandType = adCmdText
    
    
        Set rsERP = mCmd.Execute
        If rsERP.EOF Then
            Write_Log (": No row found for Invoice remain amout > 0 ")
    
            rsISS.Close
            rsERP.Close
    
            Write_Log (": ARReport Finished. . .")
            Exit Sub
        End If
        rsERP.MoveFirst
    
        Do While Not rsERP.EOF
    
                rsISS.AddNew
    
                For idx = 0 To 12
    
                    If idx = 1 Then
                        rsISS.Fields(idx) = "PCK" & Trim(rsERP.Fields(idx))
                    ElseIf idx = 11 And rsERP.Fields(idx) = " " Then
                    Else
                        rsISS.Fields(idx) = rsERP.Fields(idx)
                    End If
                Next idx
    
                rsISS.Update
    
            rsERP.MoveNext
    
        Loop
    
        rsERP.Close
    
        Set rsERP = New ADODB.Recordset
        Set mCmd = New ADODB.Command
        mCmd.ActiveConnection = conERP1
    
        'Select invoice paid with PDC
        mCmd.CommandText = " SELECT DISTINCT '30'AS Company, ARTRN.slmcod AS SalesmanCode, 'PCK'+ARTRN.cuscod AS CustomerCode, " _
                                                            + " ARTRN.docnum AS BillNo, DTOS(ARTRN.docdat) AS BillDate, ARRCPIT.rcpnum AS TaxInvNo, " _
                                                            + " DTOS(ARTRN.duedat) AS DueDate_Txt, " _
                                                            + " ARTRN.remamt AS BalanceAmount, 'Not Due' AS BillStatus, BKTRN.chqnum AS PostDateCheque, " _
                                                            + " DTOS(BKTRN.chqdat) AS PostDateChequeDate, BKTRN.amount AS PostDateChequeAmount, 'THB' " _
                                                            + " From " + strERPPath + "\ARTRN INNER JOIN " + strERPPath + "\ARRCPIT ON ARTRN.docnum = ARRCPIT.docnum " _
                                                                         + " INNER JOIN " + strERPPath + "\BKTRN ON BKTRN.voucher = ARRCPIT.rcpnum " _
                                                            + " WHERE ARTRN.slmcod <> ' ' " _
                                                            + " AND ARTRN.remamt = 0 " _
                                                            + " AND BKTRN.chqdat > DATE() " _
                                                            + " UNION " _
                                                            + " SELECT DISTINCT '30'AS Company, ARTRN.slmcod AS SalesmanCode, 'WAT'+ARTRN.cuscod AS CustomerCode, " _
                                                            + " ARTRN.docnum AS BillNo, DTOS(ARTRN.docdat) AS BillDate, ARRCPIT.rcpnum AS TaxInvNo, " _
                                                            + " DTOS(ARTRN.duedat) AS DueDate_Txt, " _
                                                            + " ARTRN.remamt AS BalanceAmount, 'Not Due' AS BillStatus, BKTRN.chqnum AS PostDateCheque, " _
                                                            + " DTOS(BKTRN.chqdat) AS PostDateChequeDate, BKTRN.amount AS PostDateChequeAmount, 'THB' " _
                                                            + " From " + strERP1Path + "\ARTRN INNER JOIN " + strERP1Path + "\ARRCPIT ON ARTRN.docnum = ARRCPIT.docnum " _
                                                                         + " INNER JOIN " + strERP1Path + "\BKTRN ON BKTRN.voucher = ARRCPIT.rcpnum " _
                                                            + " WHERE ARTRN.slmcod <> ' ' " _
                                                            + " AND ARTRN.remamt = 0 " _
                                                            + " AND BKTRN.chqdat > DATE() "
                                                            
        mCmd.CommandType = adCmdText
    
    
        Set rsERP = mCmd.Execute
        If rsERP.EOF Then
            Write_Log (": No row found for Invoice paid by PDC ")
    
            rsISS.Close
            rsERP.Close
    
            GoTo UpdateStatus
    
        End If
    
        rsERP.MoveFirst
    
        Do While Not rsERP.EOF
    
                rsISS.AddNew
    
                For idx = 0 To 12
    
                    If idx = 1 Then
                        rsISS.Fields(idx) = "PCK" & Trim(rsERP.Fields(idx))
                    ElseIf idx = 11 And rsERP.Fields(idx) = " " Then
                    Else
                        rsISS.Fields(idx) = rsERP.Fields(idx)
                    End If
                Next idx
    
                rsISS.Update
    
            rsERP.MoveNext
    
        Loop
    
        rsERP.Close
    
        rsISS.Close
    
'
      
UpdateStatus:
    conISS.Execute "UPDATE ARReport SET SalesmanCode = RTRIM(SalesmanCode)"
    conISS.Execute "UPDATE ARReport SET BillStatus = 'Overdue' " _
                                    + " WHERE CONVERT(INTEGER,DueDate,0) < CONVERT(VARCHAR,GETDATE(),112) " _
                                    + " AND Company = '30'"
    
    conISS.Execute "UPDATE ARReport SET BillStatus = 'Due' " _
                                        + " WHERE CONVERT(INTEGER,DueDate,0) = CONVERT(VARCHAR,GETDATE(),112) " _
                                        + " AND Company = '30'"
    
     Write_Log (": ARReport Copy (Slmcode+.) for MgrCode. . .")
    
    strSQL = "         INSERT INTO ARReport"
    strSQL = strSQL & " SELECT Company,SalesmanCode+'.',CustomerCode,BillNo,BillDate,"
    strSQL = strSQL & "         TaxInvNo,DueDate,BalanceAmount,BillStatus,PostDateCheque,"
    strSQL = strSQL & "         PostDateChequeDate , PostDateChequeAmount, CurrencyCode"
    strSQL = strSQL & "   From ARReport"
    strSQL = strSQL & "   Where Company ='30'"
    strSQL = strSQL & "   and SalesmanCode <> '" & Mgrcode & "'"
    conISS.Execute strSQL
    
    Write_Log (": ARReport Copy (Slmcode+..) for MgrCode1. . .")
    
    strSQL = "         INSERT INTO ARReport"
    strSQL = strSQL & " SELECT Company,SalesmanCode+'..',CustomerCode,BillNo,BillDate,"
    strSQL = strSQL & "         TaxInvNo,DueDate,BalanceAmount,BillStatus,PostDateCheque,"
    strSQL = strSQL & "         PostDateChequeDate , PostDateChequeAmount, CurrencyCode"
    strSQL = strSQL & "   From ARReport"
    strSQL = strSQL & "   Where Company ='30'"
    strSQL = strSQL & "   and SalesmanCode <> '" & Mgrcode1 & "'"
    strSQL = strSQL & "   and SalesmanCode not like '%.'"
    conISS.Execute strSQL
    
    Write_Log (": ARReport Finished. . .")
    
End Sub

Public Sub AvailableStockTransaction()

    Write_Log (": AvailableStockTransaction Start. . .")
                                            
    'Clear Table before Insert new records
    conISS.Execute "DELETE FROM AvailableStockTransaction WHERE Company = '30'"

    Set rsISS = New ADODB.Recordset
    rsISS.Open "SELECT * FROM AvailableStockTransaction ", conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP1
    
    mCmd.CommandText = "SELECT '30'AS Company, 'PCK'+RTRIM(STLOC.stkcod) AS ProductCode, STLOC.loccod AS Plant, T1.typdes AS PlantName, " _
                                                    + " locbal AS AvaiQty, T2.typdes AS Unit, 'NOR', '', '', '' " _
                                                    + " FROM " + strERPPath + "\STMAS, " + strERPPath + "\STLOC, " + strERPPath + "\ISTAB T1, " + strERPPath + "\ISTAB T2 " _
                                                    + " WHERE STLOC.stkcod = STMAS.stkcod " _
                                                    + " AND STLOC.loccod = T1.typcod " _
                                                    + " AND T1.tabtyp = '21' " _
                                                    + " AND STMAS.qucod = T2.typcod " _
                                                    + " AND T2.tabtyp = '20' " _
                                                    + " AND locbal > 0 " _
                                                    + " UNION " _
                                                    + " SELECT '30'AS Company, 'WAT'+RTRIM(STLOC.stkcod) AS ProductCode, STLOC.loccod AS Plant, T1.typdes AS PlantName, " _
                                                    + " locbal AS AvaiQty, T2.typdes AS Unit, 'NOR', '', '', '' " _
                                                    + " FROM " + strERP1Path + "\STMAS, " + strERP1Path + "\STLOC, " + strERP1Path + "\ISTAB T1, " + strERP1Path + "\ISTAB T2 " _
                                                    + " WHERE STLOC.stkcod = STMAS.stkcod " _
                                                    + " AND STLOC.loccod = T1.typcod " _
                                                    + " AND T1.tabtyp = '21' " _
                                                    + " AND STMAS.qucod = T2.typcod " _
                                                    + " AND T2.tabtyp = '20' " _
                                                    + " AND locbal > 0 "
                                                    
    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 9
            rsISS.Fields(idx) = rsERP.Fields(idx)
        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
    
    rsISS.Close
    rsERP.Close
    
    
    Write_Log (": AvailableStockTransaction Finished. . .")
    
End Sub

Public Sub Customer()

    Write_Log (": Customer Start. . .")
    Dim a1 As String
    
    
    'Clear Table before Insert new records
    conISS.Execute "DELETE FROM Customer WHERE Company = '30'"
                                           
    Set rsISS = New ADODB.Recordset
    rsISS.Open "SELECT * FROM Customer ", conISS, adOpenDynamic, adLockOptimistic

    'Bill To
    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
    
    mCmd.CommandText = "SELECT '30', 'PCK'+TRIM(cuscod), 'PCK'+TRIM(cuscod), LTRIM(RTRIM(prenam))+' '+LTRIM(RTRIM(cusnam)), TRIM(addr01)+ ' ' +TRIM(addr02)+' '+TRIM(addr03), areacod, contact, " _
                                                                        + " telnum, taxid, crline, balance, crline-balance, 0, '', 0, paycond, paytrm, remark, custyp, '', custyp, tabpr, '', 'PCK', '', '30', 1, 7 " _
                                                                        + " From " + strERPPath + "\ARMAS " _
                                                                        + " WHERE cuscod <> ' ª√‘ßø‘≈¥' " _
                                                                        + " UNION " _
                                                                        + " SELECT '30', 'WAT'+TRIM(cuscod), 'WAT'+TRIM(cuscod), TRIM(prenam)+' '+TRIM(cusnam), TRIM(addr01)+ ' ' +TRIM(addr02)+' '+TRIM(addr03), areacod, contact, " _
                                                                        + " telnum, taxid, crline, balance, crline-balance, 0, '', 0, paycond, paytrm, remark, custyp, '', custyp, tabpr, '', 'PCK', '', '30', 1, 7 " _
                                                                        + " From " + strERP1Path + "\ARMAS " _
                                                                        + " WHERE cuscod <> ' ª√‘ßø‘≈¥' " _
                                                                        + " UNION " _
                                                                        + " SELECT '30', 'PWA'+TRIM(cuscod), 'PWA'+TRIM(cuscod), TRIM(prenam)+' '+TRIM(cusnam), TRIM(addr01)+ ' ' +TRIM(addr02)+' '+TRIM(addr03), areacod, contact, " _
                                                                        + " telnum, taxid, crline, balance, crline-balance, 0, '', 0, paycond, paytrm, remark, custyp, '', custyp, tabpr, '', 'PCK', '', '30', 1, 7 " _
                                                                        + " From " + strERPPath + "\ARMAS " _
                                                                        + " WHERE cuscod <> ' ª√‘ßø‘≈¥' " _
                                                                        + " AND cuscod in (SELECT cuscod From " + strERP1Path + "\ARMAS )"
                                                                       
    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        
        rsISS.AddNew
        
        For idx = 0 To 27

            If (rsISS.Properties.Item(idx).Type = adInteger Or adNumeric) And rsERP.Fields(idx) = "" Then
                rsISS.Fields(idx) = 0
            Else

                rsISS.Fields(idx) = rsERP.Fields(idx)
            End If

        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
    
    rsERP.Close
    
    Write_Log (": Customer Bill To Finished. . .")
    
    'Ship To
    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP

    mCmd.CommandText = "SELECT '30'AS Company, 'PCK'+TRIM(ARSHIP.cuscod)+'-'+TRIM(ARSHIP.shipto) AS CustomerCode, 'PCK'+ARSHIP.cuscod AS ParentCustomerCode, " _
                                                                        + " LTRIM(RTRIM(ARMAS.prenam))+' '+LTRIM(RTRIM(ARMAS.cusnam)) AS CustomerName, " _
                                                                        + " TRIM(ARSHIP.addr01)+' '+TRIM(ARSHIP.addr02)+' '+TRIM(ARSHIP.addr03) AS Address, ARMAS.areacod AS ProvinceCode, " _
                                                                        + " ARSHIP.contact AS ContactPerson, ARSHIP.telnum AS TelephoneNumber, " _
                                                                        + " ARMAS.taxid AS TaxNumber, 0, 0, 0, 0, '', 0, ARMAS.paycond, ARMAS.paytrm, ARMAS.remark, ARMAS.custyp, '', '', 0, '', 'PCK', '', '30', 1, 7 " _
                                                                        + " From " + strERPPath + "\ARSHIP, " + strERPPath + "\ARMAS " _
                                                                        + " WHERE ARSHIP.cuscod = ARMAS.cuscod " _
                                                                        + " AND slmcod <> ' ' " _
                                                                        + " UNION " _
                                                                        + " SELECT '30'AS Company, 'WAT'+RTRIM(ARSHIP.cuscod)+'-'+TRIM(ARSHIP.shipto) AS CustomerCode, 'PCK'+ARSHIP.cuscod AS ParentCustomerCode, " _
                                                                        + " LTRIM(RTRIM(ARMAS.prenam))+' '+LTRIM(RTRIM(ARMAS.cusnam)) AS CustomerName, " _
                                                                        + " TRIM(ARSHIP.addr01)+' '+TRIM(ARSHIP.addr02)+' '+TRIM(ARSHIP.addr03) AS Address, ARMAS.areacod AS ProvinceCode, " _
                                                                        + " ARSHIP.contact AS ContactPerson, ARSHIP.telnum AS TelephoneNumber, " _
                                                                        + " ARMAS.taxid AS TaxNumber, 0, 0, 0, 0, '', 0, ARMAS.paycond, ARMAS.paytrm, ARMAS.remark, ARMAS.custyp, '', '', 0, '', 'PCK', '', '30', 1, 7 " _
                                                                        + " From " + strERP1Path + "\ARSHIP, " + strERP1Path + "\ARMAS " _
                                                                        + " WHERE ARSHIP.cuscod = ARMAS.cuscod " _
                                                                        + " AND slmcod <> ' ' "
    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    
    If rsERP.EOF Then
        Write_Log (": No row found for Customer Ship To ")
    Else
        rsERP.MoveFirst
        
        Do While Not rsERP.EOF
            rsISS.AddNew
            
            For idx = 0 To 27
                If (rsISS.Properties.Item(idx).Type = adInteger Or adNumeric) And rsERP.Fields(idx) = "" Then
                    rsISS.Fields(idx) = 0
                Else
                    rsISS.Fields(idx) = rsERP.Fields(idx)
                End If
    
            Next idx
            
            rsISS.Update
            rsERP.MoveNext
                
        Loop
        
    End If
    
    rsISS.Close
    rsERP.Close
    
    conISS.Execute "UPDATE c " _
                            + " SET c.CreditLimitUsed = a.sumA " _
                            + " FROM Customer As c " _
                            + " INNER JOIN " _
                            + " (Select CustomerCode, SUM(BalanceAmount) sumA FROM ARReport WHERE company = '30' AND SalesmanCode not like '%.' GROUP BY CustomerCode) a" _
                            + " ON a.CustomerCode = c.CustomerCode" _
                            + " WHERE Company = '30'"
                            
    'UPDATE credit used,credit outstanding
     Write_Log (": UPDATE Customer.credit_used,credit_outstanding Start. . .")

    Set rsISS = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conISS

    strSQL = "          SELECT Company, customercode, sum(BalanceAmount )"
    strSQL = strSQL & " From ARReport"
    strSQL = strSQL & " Where Company ='30'"
    strSQL = strSQL & " and balanceamount <> 0"
    strSQL = strSQL & " and SalesmanCode not like '%.'"
    strSQL = strSQL & " Group by Company, CustomerCode"

    mCmd.CommandText = strSQL
    mCmd.CommandType = adCmdText

    Set rsISS = mCmd.Execute

    If Not rsISS.EOF Then
        rsISS.MoveFirst
    End If

    Do While Not rsISS.EOF

        conISS.Execute " UPDATE Customer " _
                   + "    SET CreditLimitUsed = " & rsISS.Fields(2) & " " _
                   + "  WHERE Company = '" & rsISS.Fields(0) & "'" _
                   + "   AND CustomerCode = '" & rsISS.Fields(1) & "'"

        rsISS.MoveNext

    Loop

    rsISS.Close
    Set rsISS = Nothing

    conISS.Execute " UPDATE Customer " _
                + "    SET CreditLimitOutstanding = CreditLimit - CreditLimitUsed  " _
                + "  WHERE Company = '30'"

    Write_Log (": UPDATE Customer.credit_used,credit_outstanding Finished. . .")


                                        
    conISS.Execute "UPDATE Customer SET ProvinceCode = 'UAS' WHERE ProvinceCode = '' OR ProvinceCode IS NULL "
    
    Write_Log (": Customer Ship To Finished. . .")
    
End Sub

Public Sub CustomerSalesman()

    Write_Log (": CustomerSalesman Start. . .")
    
    conISS.Execute "DELETE FROM CustomerSaleman WHERE Company = '30'"
    
    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM CustomerSaleman "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic

    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
    
    'For Salesman - Bill To
    mCmd.CommandText = "SELECT '30', 'PCK'+TRIM(cuscod), 'PCK'+TRIM(slmcod) " _
                                                    + " From " + strERPPath + "\ARMAS " _
                                                    + " WHERE slmcod IS NOT NULL AND slmcod <> ' ' " _
                                                    + " AND cuscod <> ' ª√‘ßø‘≈¥' " _
                                                    + " UNION " _
                                                    + " SELECT '30', 'WAT'+TRIM(cuscod), 'PCK'+TRIM(slmcod) " _
                                                    + " From " + strERP1Path + "\ARMAS " _
                                                    + " WHERE slmcod IS NOT NULL AND slmcod <> ' ' " _
                                                    + " AND cuscod <> ' ª√‘ßø‘≈¥' "


    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 2
            rsISS.Fields(idx) = rsERP.Fields(idx)
        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
    
    rsERP.Close
    
    'For Salesman - Ship To
    mCmd.CommandText = "SELECT '30', 'PCK'+TRIM(ARSHIP.cuscod)+'-'+TRIM(ARSHIP.shipto) AS sh, TRIM(slmcod) " _
                                                    + " From ARSHIP, ARMAS " _
                                                    + " WHERE ARSHIP.cuscod = ARMAS.cuscod " _
                                                    + " AND slmcod IS NOT NULL AND slmcod <> ' ' "
    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    
    If rsERP.EOF Then
        Write_Log (": No row found for SalesmanCustomer Salesman-Ship To ")
        
        rsERP.Close
    
        Write_Log (": SalesmanCustomer Salesman-Ship To Finished. . .")
    Else
    
        rsERP.MoveFirst
        
        Do While Not rsERP.EOF
            rsISS.AddNew
            
            For idx = 0 To 2
                If idx = 2 Then
                    rsISS.Fields(idx) = "PCK" & Trim(rsERP.Fields(idx))
                Else
                    rsISS.Fields(idx) = rsERP.Fields(idx)
                End If
            Next idx
            
            rsISS.Update
            rsERP.MoveNext
                
        Loop
        
        rsERP.Close
    
    End If
    
    'For Manager - Bill To
'    mCmd.CommandText = "SELECT DISTINCT '30', 'PCK'+TRIM(ARMAS.cuscod), TRIM(positn) " _
'                                                    + " From ARMAS, OESLM " _
'                                                    + " WHERE ARMAS.slmcod = OESLM.slmcod " _
'                                                    + " AND OESLM.positn IS NOT NULL " _
'                                                    + " AND OESLM.positn <> ' ' "
'
'    mCmd.CommandType = adCmdText
'
'    Set rsERP = mCmd.Execute
'
'    If rsERP.EOF Then
'        Write_Log (": No row found for SalesmanCustomer Manager-Bill To ")
'
'        rsERP.Close
'
'        Write_Log (": SalesmanCustomer Manager-Bill To Finished. . .")
'    Else
'        rsERP.MoveFirst
'
'        Do While Not rsERP.EOF
'            rsISS.AddNew
'
'            For idx = 0 To 2
'                If idx = 2 Then
'                    rsISS.Fields(idx) = "PCK" & Trim(rsERP.Fields(idx))
'                Else
'                    rsISS.Fields(idx) = rsERP.Fields(idx)
'                End If
'            Next idx
'
'            rsISS.Update
'            rsERP.MoveNext
'
'        Loop
'
'        rsERP.Close
'    End If
'
'    For Manager - Ship To
'    mCmd.CommandText = "SELECT DISTINCT '30', 'PCK'+TRIM(ARSHIP.cuscod)+'-'+TRIM(ARSHIP.shipto), TRIM(positn) " _
'                                                    + " From ARMAS, OESLM, ARSHIP " _
'                                                    + " WHERE ARMAS.slmcod = OESLM.slmcod " _
'                                                    + " AND OESLM.positn IS NOT NULL " _
'                                                    + " AND OESLM.positn <> ' ' " _
'                                                    + " AND ARMAS.cuscod = ARSHIP.cuscod "
'
'    mCmd.CommandType = adCmdText
'
'    Set rsERP = mCmd.Execute
'    If rsERP.EOF Then
'        Write_Log (": No row found for SalesmanCustomer Manager-Ship To ")
'
'        rsISS.Close
'        rsERP.Close
'
'        Write_Log (": SalesmanCustomer Manager-Ship To Finished. . .")
'    Else
'        rsERP.MoveFirst
'
'        Do While Not rsERP.EOF
'            rsISS.AddNew
'
'            For idx = 0 To 2
'                If idx = 2 Then
'                    rsISS.Fields(idx) = "PCK" & Trim(rsERP.Fields(idx))
'                Else
'                    rsISS.Fields(idx) = rsERP.Fields(idx)
'                End If
'            Next idx
'
'            rsISS.Update
'            rsERP.MoveNext
'
'        Loop
'
'        rsERP.Close
'        rsISS.Close
'    End If
    
    conISS.Execute "UPDATE CustomerSaleman SET SalesmanCode = RTRIM(SalesmanCode) "
    
    
    'for manager
    strSQL = "    INSERT INTO CustomerSaleman"
    strSQL = strSQL & " SELECT DISTINCT Company , CustomerCode ,'" & Mgrcode & "'"
    strSQL = strSQL & " From CustomerSaleman"
    strSQL = strSQL & " Where Company ='30'"
    strSQL = strSQL & " and SalesmanCode <> '" & Mgrcode & "'"
    strSQL = strSQL & " and CustomerCode NOT IN (Select customercode From CustomerSaleman"
    strSQL = strSQL & "                          Where Company ='30'"
    strSQL = strSQL & "                          and SalesmanCode = '" & Mgrcode & "')"

    conISS.Execute strSQL
    
    
    'for manager1
    strSQL = "    INSERT INTO CustomerSaleman"
    strSQL = strSQL & " SELECT DISTINCT Company , CustomerCode ,'" & Mgrcode1 & "'"
    strSQL = strSQL & " From CustomerSaleman"
    strSQL = strSQL & " Where Company ='30'"
    strSQL = strSQL & " and SalesmanCode <> '" & Mgrcode1 & "'"
    strSQL = strSQL & " and SalesmanCode <> '" & Mgrcode & "'"
    strSQL = strSQL & " and CustomerCode NOT IN (Select customercode From CustomerSaleman"
    strSQL = strSQL & "                          Where Company ='30'"
    strSQL = strSQL & "                          and SalesmanCode = '" & Mgrcode1 & "')"

    conISS.Execute strSQL
                                          
    Write_Log (": CustomerSaleman Finished. . .")
    
End Sub

'Public Sub CustomerPricing()
''On Error GoTo Err_Handler
''
''    Dim product1 As String
''    Dim qty(4) As Integer
''    Dim subcnt As Integer, i As Integer
''
''    Write_Log (": Tmp_CustomerPricing Start. . .")
''
''    'Clear Table before Insert new records
''    conISS.Execute "DELETE FROM Tmp_CustomerPricing WHERE Company = '30' "
''    conISS.Execute "DELETE FROM CustomerPricing WHERE Company = '30' "
''
''
''    Set rsISS = New ADODB.Recordset
''    strSQL = "SELECT * FROM Tmp_CustomerPricing WHERE 1=2"
''    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic
''
''
''    Set rsERP = New ADODB.Recordset
''    Set mCmd = New ADODB.Command
''    mCmd.ActiveConnection = conERP
''
''    strSQL = "          SELECT CASE '30','PCK'+people, stkcod, 0, 0, 0, MAX(rdocnum), 'THB' "
''    strSQL = strSQL & " FROM From STCRD "
''    strSQL = strSQL & " WHERE slmcod <> ' '"
''    strSQL = strSQL & " And SUBSTR(docnum, 1, 2) IN ('IB', 'IS')"
''    strSQL = strSQL & " Group by people, stkcod"
''
''
''    mCmd.CommandText = strSQL
''    mCmd.CommandType = adCmdText
''
''    Set rsERP = mCmd.Execute
''    If Not rsERP.EOF Then
''        rsERP.MoveFirst
''        product1 = ""
''    End If
''
''    Do While Not rsERP.EOF
'''        If product1 <> rsERP.Fields(2) Then
'''            product1 = rsERP.Fields(2)
''            rsISS.AddNew
''
''            For idx = 0 To 7
''                rsISS.Fields(idx) = rsERP.Fields(idx)
''            Next idx
''
''            'rsERP.Fields(4)=d.goodprice
''
''
''
''            rsISS.Update
''            rsERP.MoveNext
'''        Else
'''            rsERP.MoveNext
'''        End If
''    Loop
''
''    rsERP.Close
''    rsISS.Close
''    Set rsERP = Nothing
''    Set rsISS = Nothing
''
''    Write_Log (": Tmp_CustomerPricing Finished. . .")
''
''
'''    'Tmp_CustomerPricing Update Price1,2,3,4
'''    Write_Log (": Tmp_CustomerPricing Update Price1,2,3,4 Start. . .")
'''
'''    Set rsERP = New ADODB.Recordset
'''    Set mCmd = New ADODB.Command
'''    mCmd.ActiveConnection = conERP
'''
'''    strSQL = "          SELECT DISTINCT g.GoodCode ,m.goodunitrate "
'''    strSQL = strSQL & " From EMSetPriceHD h"
'''    strSQL = strSQL & "     ,EMSetPriceDT d, EMGood g"
'''    strSQL = strSQL & "     ,EMGoodMultiUnit m"
'''    strSQL = strSQL & " Where"
'''    strSQL = strSQL & "     h.SetPriceID = d.SetPriceID"
'''    strSQL = strSQL & " and d.ListID = g.goodid"
'''    strSQL = strSQL & " and d.ListID = m.GoodID"
'''    strSQL = strSQL & " and g.Inactive = 'A'"
'''    strSQL = strSQL & " and h.custflag = 'C'"
'''    strSQL = strSQL & " and h.docuflag = 'Y'"
'''    strSQL = strSQL & " and GETDATE() between h.BeginDate and h.EndDate"
'''    strSQL = strSQL & " ORDER by g.GoodCode,m.GoodUnitRate DESC"
'''
'''    mCmd.CommandText = strSQL
'''    mCmd.CommandType = adCmdText
'''
'''    Set rsERP = mCmd.Execute
'''
'''    If Not (rsERP.EOF) Then
'''        rsERP.MoveFirst
'''
'''        'clear array
'''        For i = 1 To 4
'''            qty(i) = 0
'''        Next i
'''
'''        product1 = "START"
'''        subcnt = 0
'''        qty(1) = rsERP.Fields(1)
'''    End If
'''
'''    Do While Not rsERP.EOF
'''
'''        'rsERP.Fields(x)
'''        '0 = productcode
'''        '1 = rate_Qty
'''        If product1 <> rsERP.Fields(0) Then
'''            If product1 = "START" Then product1 = rsERP.Fields(0)
'''
'''            strSQL = "          UPDATE Tmp_CustomerPricing "
'''            strSQL = strSQL & " SET price1 = price1*" & qty(1)
'''            strSQL = strSQL & "    ,price2 = price1*" & qty(2)
'''            strSQL = strSQL & "    ,price3 = price1*" & qty(3)
'''            strSQL = strSQL & "    ,price4 = price1*" & qty(4)
'''            strSQL = strSQL & " WHERE Company = '30'"
'''            strSQL = strSQL & " AND productcode = '" & product1 & "'"
'''            conISS.Execute strSQL
'''
'''            'clear array
'''            For i = 1 To 4
'''                qty(i) = 0
'''            Next i
'''            product1 = rsERP.Fields(0)
'''            subcnt = 1
'''
'''        Else
'''            subcnt = subcnt + 1
'''        End If
'''
'''        If subcnt >= 1 And subcnt <= 4 Then
'''            qty(subcnt) = rsERP.Fields(1)
'''        End If
'''
'''        rsERP.MoveNext
'''    Loop
''
''    strSQL = "          UPDATE Tmp_CustomerPricing "
''    strSQL = strSQL & " SET price1 = price1*" & qty(1)
''    strSQL = strSQL & "    ,price2 = price1*" & qty(2)
''    strSQL = strSQL & "    ,price3 = price1*" & qty(3)
''    strSQL = strSQL & "    ,price4 = price1*" & qty(4)
''    strSQL = strSQL & " WHERE Company = '30'"
''    strSQL = strSQL & " AND productcode = '" & product1 & "'"
''    conISS.Execute strSQL
''
''    rsERP.Close
''    Set rsERP = Nothing
''
''    Write_Log (": Tmp_CustomerPricing Update Price1,2,3,4 Finished. . .")
''
''
''    Write_Log (": CustomerPricing Start. . .")
''
''    conISS.Execute " INSERT INTO CustomerPricing SELECT * FROM Tmp_CustomerPricing WHERE Company = '30'"
''
''    Write_Log (": CustomerPricing Finished. . .")
''
''    Exit Sub
''Err_Handler:
''    Write_Log ("Error : " & Err.Number & " " & Err.Description)
''    MsgBox "Error : " & Err.Number & " " & Err.Description
''    End


Public Sub CustomerPricing()

    Write_Log (": Tmp_CustomerPricing Start. . .")
    
                        
    'Clear Table before Insert new records
'    conISS.Execute "DELETE FROM Tmp_CustomerPricing WHERE Company = '30'"
    conISS.Execute "DELETE FROM CustomerPricing WHERE Company = '30' "
    
    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM CustomerPricing"
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
    
   
  'THB
    strSQL = "SELECT DISTINCT '30','PCK'+people, 'PCK'+RTRIM(stkcod), max(unitpr), 0, 0, 0, 'THB' " _
                        + "  FROM " + strERPPath + "\STCRD " _
                        + "  WHERE slmcod <> ' '" _
                        + "  And SUBSTR(docnum, 1, 2) IN ('IB', 'IS')" _
                        + "  And rdocnum in (SELECT MAX(rdocnum)From STCRD WHERE slmcod <> ' ' And SUBSTR(docnum, 1, 2) IN ('IB', 'IS') Group by people, stkcod)" _
                        + "  Group by people, stkcod" _
                        + "  UNION " _
                        + " SELECT DISTINCT '30','WAT'+people, 'WAT'+RTRIM(stkcod), max(unitpr), 0, 0, 0, 'THB' " _
                        + "  FROM " + strERP1Path + "\STCRD " _
                        + "  WHERE slmcod <> ' '" _
                        + "  And SUBSTR(docnum, 1, 2) IN ('IV')" _
                        + "  And rdocnum in (SELECT MAX(rdocnum)From STCRD WHERE slmcod <> ' ' And SUBSTR(docnum, 1, 2) IN ('IB', 'IS') Group by people, stkcod)" _
                        + "  Group by people, stkcod"
                        
    mCmd.CommandText = strSQL
                                                    
    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 7
            rsISS.Fields(idx) = rsERP.Fields(idx)

        Next idx
    
         
        rsISS.Update
        rsERP.MoveNext
        
        
            
    Loop
    
    rsERP.Close
    
    'USD
    strSQL = "SELECT DISTINCT '30','PCK'+people, 'PCK'+RTRIM(stkcod), max(unitpr), 0, 0, 0, 'USD' " _
                        + "  FROM " + strERPPath + "\STCRD " _
                        + "  WHERE slmcod <> ' '" _
                        + "  And SUBSTR(docnum, 1, 2) IN ('IB', 'IS')" _
                        + "  And rdocnum in (SELECT MAX(rdocnum)From STCRD WHERE slmcod <> ' ' And SUBSTR(docnum, 1, 2) IN ('IB', 'IS') Group by people, stkcod)" _
                        + "  Group by people, stkcod" _
                        + "  UNION " _
                        + " SELECT DISTINCT '30','WAT'+people, 'WAT'+RTRIM(stkcod), max(unitpr), 0, 0, 0, 'USD' " _
                        + "  FROM " + strERP1Path + "\STCRD " _
                        + "  WHERE slmcod <> ' '" _
                        + "  And SUBSTR(docnum, 1, 2) IN ('IV')" _
                        + "  And rdocnum in (SELECT MAX(rdocnum)From STCRD WHERE slmcod <> ' ' And SUBSTR(docnum, 1, 2) IN ('IB', 'IS') Group by people, stkcod)" _
                        + "  Group by people, stkcod"
                        
    mCmd.CommandText = strSQL
                                                    
    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 7
            rsISS.Fields(idx) = rsERP.Fields(idx)

        Next idx
    
         
        rsISS.Update
        rsERP.MoveNext
        
        
            
    Loop
    
    rsERP.Close
    rsISS.Close
      Write_Log (": Tmp_CustomerPricing Finished. . .")
       Write_Log (": CustomerPricing Start. . .")

'    conISS.Execute " INSERT INTO CustomerPricing SELECT * FROM Tmp_CustomerPricing WHERE Company = '30'"
    

    Write_Log (": GroupPricing Finished. . .")
    

End Sub


Public Sub Daily_SalesSummaries()
    
    Write_Log (": Daily_SalesSummaries Start. . .")
    
    'Clear Table before Insert new records
    conISS.Execute "DELETE FROM Report_Daily_SalesSummaries " _
                                        + " WHERE Company = '30'" _
'                                        + " AND YEAR(BillDate) = " + CStr(intCurrYear) _
'                                        + " AND MONTH(BillDate) = " + CStr(intCurrMonth)
                                        
                                        
    conISS.Execute "DELETE FROM Tmp_Report_Daily_SalesSummaries " _
                                        + " WHERE Company = '30'"
    'pck
    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM Tmp_Report_Daily_SalesSummaries "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
                                                        
      mCmd.CommandText = "SELECT '30', slmcod, 'PCK'+people, 'PCK'+RTRIM(stkcod), DTOS(docdat), 'NOR',  " _
                                                        + " SUM(trnqty), SUM(netval), 'THB'  " _
                                                        + " FROM " + strERPPath + "\STCRD " _
                                                        + " WHERE SUBSTR(docnum, 1, 2) IN ('IB', 'IS') " _
                                                        + " AND pstkcod = ' ' " _
                                                        + " GROUP BY slmcod, people, stkcod, docdat " _
                                                        + " UNION " _
                                                        + " SELECT '30', slmcod, 'WAT'+people, 'WAT'+RTRIM(stkcod), DTOS(docdat), 'NOR',  " _
                                                        + " SUM(trnqty), SUM(netval)*-1, 'THB'  " _
                                                        + " FROM " + strERPPath + "\STCRD " _
                                                        + " WHERE pstkcod = ' ' " _
                                                        + " AND SUBSTR(docnum, 1, 2) IN ('SR') " _
                                                        + " GROUP BY slmcod, people, stkcod, docdat "
                                                        
    mCmd.CommandType = adCmdText
'
'    + " AND YEAR(docdat) = " + Str(intCurrYear) _
'                                                        + " AND MONTH(docdat) = " + Str(intCurrMonth) _

    Set rsERP = mCmd.Execute
    If rsERP.EOF Then
        Write_Log (": No row found for Period: " + Str(intCurrYear) + Format(intCurrMonth, "00"))
        
        rsISS.Close
        rsERP.Close
    
        Write_Log (": Daily_SalesSummaries Finished. . .")
        Exit Sub
    End If
    
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 8
            If idx = 1 Then
                rsISS.Fields(idx) = "PCK" & Trim(rsERP.Fields(idx))
            Else
                rsISS.Fields(idx) = rsERP.Fields(idx)
            End If
        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
        
    rsISS.Close
    
        'WAT
    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM Tmp_Report_Daily_SalesSummaries "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
                                                        
      mCmd.CommandText = "SELECT '30', slmcod, 'WAT'+people, 'WAT'+RTRIM(stkcod), DTOS(docdat), 'NOR',  " _
                                                        + " SUM(trnqty), SUM(netval), 'THB'  " _
                                                        + " FROM " + strERP1Path + "\STCRD " _
                                                        + " WHERE SUBSTR(docnum, 1, 2) IN ('IV') " _
                                                        + " AND pstkcod = ' ' " _
                                                        + " GROUP BY slmcod, people, stkcod, docdat " _
                                                        + " UNION " _
                                                        + " SELECT '30', slmcod, 'WAT'+people, 'WAT'+RTRIM(stkcod), DTOS(docdat), 'NOR',  " _
                                                        + " SUM(trnqty)*-1, +SUM(netval)*-1, 'THB'  " _
                                                        + " FROM " + strERP1Path + "\STCRD " _
                                                        + " WHERE pstkcod = ' ' " _
                                                        + " AND SUBSTR(docnum, 1, 2) IN ('SR') " _
                                                        + " GROUP BY slmcod, people, stkcod, docdat "
                                                        
    mCmd.CommandType = adCmdText
'
'    + " AND YEAR(docdat) = " + Str(intCurrYear) _
'                                                        + " AND MONTH(docdat) = " + Str(intCurrMonth) _

    Set rsERP = mCmd.Execute
    If rsERP.EOF Then
        Write_Log (": No row found for Period: " + Str(intCurrYear) + Format(intCurrMonth, "00"))
        
        rsISS.Close
        rsERP.Close
    
        Write_Log (": Daily_SalesSummaries Finished. . .")
        Exit Sub
    End If
    
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 8
            If idx = 1 Then
                rsISS.Fields(idx) = "PCK" & Trim(rsERP.Fields(idx))
            Else
                rsISS.Fields(idx) = rsERP.Fields(idx)
            End If
        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
        
    rsISS.Close
    rsERP.Close
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    strSQL = "   UPDATE Tmp_Report_Daily_SalesSummaries "
    strSQL = strSQL & " set Tmp_Report_Daily_SalesSummaries.SalesmanCode = CS.SalesmanCode "
    strSQL = strSQL & " from Tmp_Report_Daily_SalesSummaries DS inner join CustomerSaleman CS on DS.CustomerCode = CS.CustomerCode"
    strSQL = strSQL & " WHERE DS.Company = '30'"
    strSQL = strSQL & " AND CS.Company = '30'"
    conISS.Execute strSQL

       rsERP.Close
    Set rsERP = Nothing
    
    conISS.Execute " INSERT INTO Report_Daily_SalesSummaries " _
                                    + " SELECT Company, RTRIM(SalesmanCode), CustomerCode, ProductCode, BillDate, RowType, " _
                                    + " SUM(SalesQuantity), SUM(SalesAmount), CurrencyCode " _
                                    + " FROM Tmp_Report_Daily_SalesSummaries " _
                                    + " WHERE Company = '30' " _
                                    + " GROUP BY Company, SalesmanCode, CustomerCode, ProductCode, BillDate, RowType, CurrencyCode "
                                    
    'conISS.Execute "UPDATE Report_Daily_SalesSummaries SET SalesmanCode = RTRIM(SalesmanCode) "
    
    conISS.Execute " INSERT INTO Report_Daily_SalesSummaries " _
                                    + " SELECT Company, SalesmanCode+'.' , CustomerCode, ProductCode, " _
                                    + " BillDate , RowType, SalesQuantity, SalesAmount, CurrencyCode" _
                                    + " From Report_Daily_SalesSummaries" _
                                    + " Where Company ='30' " _
                                    + " And SalesmanCode <> '" & Mgrcode & "'"
    conISS.Execute strSQL
    
    Write_Log (": Daily_SalesSummaries15mth Insert (slmcode+.) for manager. . .")
    
    conISS.Execute " INSERT INTO Report_Daily_SalesSummaries " _
                                    + " SELECT Company, SalesmanCode+'..' , CustomerCode, ProductCode, " _
                                    + " BillDate , RowType, SalesQuantity, SalesAmount, CurrencyCode" _
                                    + " From Report_Daily_SalesSummaries" _
                                    + " Where Company ='30' " _
                                    + " And SalesmanCode <> '" & Mgrcode1 & "'" _
                                    + " And SalesmanCode not like '%.'"
    conISS.Execute strSQL
    
    Write_Log (": Daily_SalesSummaries15mth Insert (slmcode+..) for manager1. . .")
    
    Write_Log (": Daily_SalesSummaries Finished. . .")
    
End Sub

Public Sub GroupPricing()

    Write_Log (": GroupPricing Start. . .")
                                                
    'Clear Table before Insert new records
    conISS.Execute "DELETE FROM GroupPricing WHERE Company = '30'"
    
    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM GroupPricing"
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
    
    strSQL = "SELECT '30', '1', stkcod, sellpr1, 0, 0, 0, 'THB' FROM STMAS WHERE sellpr1 <> 0 AND squcod = qucod AND cqucod = ' ' " _
                        + " UNION SELECT '30', '1', stkcod, sellpr1, sellpr1/sfactor, 0, 0, 'THB' FROM STMAS WHERE sellpr1 <> 0 AND squcod = cqucod AND cqucod <> ' ' " _
                        + " UNION SELECT '30', '2', stkcod, sellpr2, 0, 0, 0, 'THB' FROM STMAS WHERE sellpr2 <> 0 AND squcod = qucod AND cqucod = ' ' " _
                        + " UNION SELECT '30', '2', stkcod, sellpr2, sellpr2/sfactor, 0, 0, 'THB' FROM STMAS WHERE sellpr2 <> 0 AND squcod = cqucod AND cqucod <> ' ' " _
                        + " UNION SELECT '30', '3', stkcod, sellpr3, 0, 0, 0, 'THB' FROM STMAS WHERE sellpr3 <> 0 AND squcod = qucod AND cqucod = ' ' " _
                        + " UNION SELECT '30', '3', stkcod, sellpr3, sellpr3/sfactor, 0, 0, 'THB' FROM STMAS WHERE sellpr3 <> 0 AND squcod = cqucod AND cqucod <> ' ' " _
                        + " UNION SELECT '30', '4', stkcod, sellpr4, 0, 0, 0, 'THB' FROM STMAS WHERE sellpr4 <> 0 AND squcod = qucod AND cqucod = ' ' " _
                        + " UNION SELECT '30', '4', stkcod, sellpr4, sellpr4/sfactor, 0, 0, 'THB' FROM STMAS WHERE sellpr4 <> 0 AND squcod = cqucod AND cqucod <> ' ' "
                    
    mCmd.CommandText = strSQL
                                                    
    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 7
            rsISS.Fields(idx) = rsERP.Fields(idx)
        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
    
    rsERP.Close
    rsISS.Close
                                                

    Write_Log (": GroupPricing Finished. . .")
    
End Sub

Public Sub SalesAnalysis()

    Write_Log (": SalesAnalysis Start. . .")

        Dim intFromMonth, intToMonth As Integer
        Dim strstartdate, strDate As String

        intCurrYear = Year(Date)
    strstartdate = Format(intCurrYear, "&&&&") + "0101"
    
  'Clear Table before Insert new records
    conISS.Execute "DELETE FROM Tmp_SalesAnalysis WHERE Company = '30'"
    conISS.Execute "DELETE FROM ProductTraget WHERE Company = '30' "
    
    ' Create Year to date on ProductTraget
    strSQL = "INSERT INTO ProductTraget " _
                                    + " SELECT DISTINCT Company, SalesmanCode, ProductCode, RowType, " _
                                    + " year(getdate()), month(getdate()), SUM(SalesAmount), SUM(SalesQuantity), CurrencyCode " _
                                    + " FROM Report_Daily_SalesSummaries " _
                                    + " WHERE Company = '30' " _
                                    + " AND BillDate >= '" + strstartdate + "' " _
                                    + " GROUP BY Company, SalesmanCode, ProductCode, RowType, CurrencyCode"
    conISS.Execute strSQL, intNoOfRow

' Create Sales Transaction for Current Month
strSQL = "INSERT INTO Tmp_SalesAnalysis " _
                                + " SELECT Company, SalesmanCode, CustomerCode, ProductCode, RowType, " _
                                                    + " SUM(SalesAmount), SUM(SalesQuantity), 0, 0, 0, 'THB' " _
                                + " FROM Report_Daily_SalesSummaries " _
                                + " WHERE Company = '30'" _
                                + " AND SUBSTRING(BillDate, 1, 4) = '" + Format(intCurrYear, "0000") + "' " _
                                + " AND SUBSTRING(BillDate, 5, 2) = '" + Format(intCurrMonth, "00") + "' " _
                                + " GROUP BY Company, SalesmanCode, CustomerCode, ProductCode, RowType "
conISS.Execute strSQL, intNoOfRow
                            

'Create Sales Transaction for Last Month
conISS.Execute "INSERT INTO Tmp_SalesAnalysis " _
                                + " SELECT Company, SalesmanCode, CustomerCode, ProductCode, RowType, " _
                                                    + " 0, 0, SUM(SalesAmount), SUM(SalesQuantity), 0, 'THB' " _
                                + " FROM Report_Daily_SalesSummaries " _
                                + " WHERE Company = '30'" _
                                + " AND SUBSTRING(BillDate, 1, 4) = '" + Format(intLastYear, "0000") + "' " _
                                + " AND SUBSTRING(BillDate, 5, 2) = '" + Format(intLastMonth, "00") + "' " _
                                + " GROUP BY Company, SalesmanCode, CustomerCode, ProductCode, RowType "

'Create Sales Transaction for the Same Month Last Year
conISS.Execute "INSERT INTO Tmp_SalesAnalysis " _
                                + " SELECT Company, SalesmanCode, CustomerCode, ProductCode, 'NOR', " _
                                                    + " 0, 0, 0, 0, SUM(LeftAmount), 'THB' " _
                                + " FROM OrderStatusReportItem " _
                                + " WHERE Company = '30'" _
                                + " GROUP BY Company, SalesmanCode, CustomerCode, ProductCode"

  'Clear Table before Insert new records
  conISS.Execute "DELETE FROM Report_SalesAnalysis WHERE Company = '30'"
  
conISS.Execute "INSERT INTO Report_SalesAnalysis" _
                                    + " SELECT Company,RTRIM(SalesmanCode),CustomerCode,ProductCode,RowType, " _
                                                        + " SUM(MonthBuyAmount) ,SUM(MonthBuyQuantity), " _
                                                        + " SUM(LastMonthBuyAmount) ,SUM(LastMonthBuyQuantity) ,SUM(MonthBuyCount), CurrencyCode " _
                                        + " From Tmp_SalesAnalysis " _
                                        + " WHERE Company = '30'" _
                                        + " GROUP BY Company,SalesmanCode,CustomerCode,ProductCode,RowType, CurrencyCode "
    
    Write_Log (": SalesAnalysis Finished. . .")
    
End Sub

Public Sub SalesHistory()
     Dim dteDate As Date
    Dim strSalesmanCode, strCustomerCode, strProductCode As String
    Dim dteDocDate As Date
    Dim strstartdate, strDate As String

    Write_Log (": SalesHistory Start. . .")
    
    conISS.Execute "DELETE FROM Tmp_SalesHistoryMax WHERE Company = '30'"
    
    
    dteDate = DateAdd("m", -12, Date)
    dteDate = DateAdd("yyyy", -543, dteDate)
    strstartdate = Format(dteDate, "YYYYMMDD")
    
    ' Find the Last Sales Record
    conISS.Execute "INSERT INTO Tmp_SalesHistoryMax " _
                                    + " SELECT Company, SalesmanCode, CustomerCode, ProductCode,RowType, " _
                                                            + " MAX(BillDate), 0, 0, CurrencyCode " _
                                    + " FROM Report_Daily_SalesSummaries " _
                                    + " WHERE Company = '30'" _
                                    + " AND BillDate >= '" + strstartdate + "' " _
                                    + " GROUP BY Company, SalesmanCode, CustomerCode, ProductCode,RowType, CurrencyCode "
    
    conISS.Execute "DELETE FROM Tmp_SalesHistory WHERE Company = '30'"
    conISS.Execute "DELETE FROM Tmp_SalesHistoryYearSum WHERE Company = '30'"
    
    'Create the Last Sales Record
    conISS.Execute " INSERT INTO Tmp_SalesHistory " _
                                    + " SELECT r.Company, r.SalesmanCode, r.CustomerCode, r.ProductCode, r.RowType, " _
                                                            + " 0, 0, r.SalesAmount, r.SalesQuantity, r.BillDate, m.CurrencyCode " _
                                    + " FROM Tmp_SalesHistoryMax m, Report_Daily_SalesSummaries r " _
                                    + " WHERE r.Company = '30'" _
                                    + " AND r.Company = m.Company " _
                                    + " AND r.SalesmanCode = m.SalesmanCode " _
                                    + " AND r.CustomerCode = m.CustomerCode " _
                                    + " AND r.ProductCode = m.ProductCode " _
                                    + " AND r.BillDate = m.LastBuyDate " _
                                    + " AND r.CurrencyCode = m.CurrencyCode "
   'Create the Year Summary Record
   conISS.Execute "INSERT INTO Tmp_SalesHistoryYearSum " _
                                    + " SELECT Company, SalesmanCode, CustomerCode, ProductCode, " _
                                                            + " SUM(SalesAmount), SUM(SalesQuantity), CurrencyCode " _
                                    + " FROM Report_Daily_SalesSummaries " _
                                    + " WHERE Company = '30'" _
                                    + " AND BillDate >= '" + strstartdate + "' " _
                                    + " GROUP BY Company, SalesmanCode, CustomerCode, ProductCode, CurrencyCode "

                                            
    conISS.Execute "DELETE FROM Report_SalesHistory WHERE Company = '30'"
     
    conISS.Execute "INSERT INTO Report_SalesHistory " _
                                        + " SELECT s.Company, RTRIM(s.SalesmanCode), s.CustomerCode, s.ProductCode, s.RowType, y.YearSumAmount, " _
                                                            + " y.YearBuyQuantity, s.LastBuyAmount, s.LastBuyQuantity, s.LastBuyDate, s.CurrencyCode " _
                                            + " From Tmp_SalesHistory s, Tmp_SalesHistoryYearSum y " _
                                            + " WHERE s.Company = '30'" _
                                            + " AND s.Company = y.Company " _
                                            + " AND s.SalesmanCode = y.SalesmanCode " _
                                            + " AND s.CustomerCode = y.CustomerCode " _
                                            + " AND s.ProductCode = y.ProductCode " _
                                            + " AND s.CurrencyCode = y.CurrencyCode "
    
    Write_Log (": SalesHistory Finished. . .")
    
End Sub

Public Sub SalesSummaries()

    Write_Log (": SalesSummaries Start. . .")
    
    'Clear Table before Insert new records
    conISS.Execute "DELETE FROM SalesSummaries " _
                                        + " WHERE Company = '30'" _
                                        + " AND SalesYear = " + CStr(intCurrYear) _
                                        + " AND SalesMonth = " + CStr(intCurrMonth)
                                        
    conISS.Execute "DELETE FROM Tmp_SalesSummaries " _
                                        + " WHERE Company = '30'"
                                               
    'Create Sales Summaries Record
    conISS.Execute "INSERT INTO Tmp_SalesSummaries " _
                                    + " SELECT Company, RTRIM(SalesmanCode), RowType, SUBSTRING(BillDate, 1, 4), SUBSTRING(BillDate, 5, 2), SUM(SalesAmount), CurrencyCode " _
                                    + " FROM Report_Daily_SalesSummaries " _
                                    + " WHERE Company = '30'" _
                                    + " AND SUBSTRING(BillDate, 1, 4) = '" + Format(intCurrYear, "0000") + "' " _
                                    + " AND SUBSTRING(BillDate, 5, 2) = '" + Format(intCurrMonth, "00") + "' " _
                                    + " GROUP BY Company, SalesmanCode, RowType, SUBSTRING(BillDate, 1, 4), SUBSTRING(BillDate, 5, 2), CurrencyCode "

    
    ' Create record for Manager
    conISS.Execute "INSERT INTO Tmp_SalesSummaries " _
                                    + " SELECT ss.Company, RTRIM(s.ParentSalesmanCode), RowType, SalesYear, SalesMonth, SUM(MonthSales), CurrencyCode " _
                                    + " FROM SalesSummaries ss, Salesman s " _
                                    + " WHERE ss.Company = '30'" _
                                    + " AND ss.Company = s.Company " _
                                    + " AND ss.SalesmanCode = s.SalesmanCode " _
                                    + " AND SalesYear = " + Str(intCurrYear) _
                                    + " AND SalesMonth = " + Str(intCurrMonth) _
                                    + " GROUP BY ss.Company, s.ParentSalesmanCode, ss.RowType, ss.SalesYear, ss.SalesMonth, CurrencyCode "

    
        conISS.Execute "INSERT INTO SalesSummaries " _
                                    + " SELECT Company, SalesmanCode, RowType, SalesYear, SalesMonth, SUM(MonthSales), CurrencyCode " _
                                    + " FROM Tmp_SalesSummaries " _
                                    + " WHERE Company = '30'" _
                                    + " GROUP BY Company, SalesmanCode, RowType, SalesYear, SalesMonth, CurrencyCode "
    
    Write_Log (": SalesSummaries Finished. . .")
    
End Sub

Public Sub SingleMaster()
    Dim Date1 As Date
    Dim intPrvCode As Integer
    Dim curTotalAmount As Currency
    Dim intOrderQty, intCancelQty, intRemQty As Integer
    

    Write_Log (": SingleMaster Start. . .")
    
    'Transportation Method (Lookup)
    
'    Write_Log (": Transportation Method-Lookup Start. . .")
'    conISS.Execute "DELETE FROM Lookup WHERE Company = '30'"
'
'    Set rsISS = New ADODB.Recordset
'    strSQL = "SELECT * FROM Lookup "
'    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic
'
'
'    Set rsERP = New ADODB.Recordset
'    Set mCmd = New ADODB.Command
'    mCmd.ActiveConnection = conERP
'
'    mCmd.CommandText = "SELECT '30', 'ModeDly', typcod, typdes, 1 " _
'                                                        + " FROM ISTAB " _
'                                                        + " WHERE tabtyp = '41' "
'
'    mCmd.CommandType = adCmdText
'
'    Set rsERP = mCmd.Execute
'    rsERP.MoveFirst
'
'    Do While Not rsERP.EOF
'        rsISS.AddNew
'
'        For idx = 0 To 4
'            rsISS.Fields(idx) = rsERP.Fields(idx)
'        Next idx
'
'        rsISS.Update
'        rsERP.MoveNext
'
'    Loop
'
'    rsERP.Close
'    rsISS.Close
'
'    Write_Log (": Transportation Method-Lookup Finished. . .")
    
    'Customer Group
    
    Write_Log (": CustomerGroupMaster Start. . .")
    conISS.Execute "DELETE FROM CustomerGroupMaster WHERE Company = '30'"
    
    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM CustomerGroupMaster "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
    
    mCmd.CommandText = "SELECT DISTINCT '30', typcod, typdes, 1 " _
                                                        + " FROM " + strERPPath + "\ARMAS, " + strERPPath + "\ISTAB " _
                                                        + " WHERE ISTAB.tabtyp = '45' " _
                                                        + " UNION " _
                                                        + " SELECT DISTINCT '30', typcod, typdes, 1 " _
                                                        + " FROM " + strERP1Path + "\ARMAS, " + strERP1Path + "\ISTAB " _
                                                        + " WHERE ISTAB.tabtyp = '45' " _
                                                        + " AND typcod not in (SELECT typcod " _
                                                        + " FROM " + strERPPath + "\ARMAS, " + strERPPath + "\ISTAB " _
                                                        + " WHERE ISTAB.tabtyp = '45' ) "
                                                    
    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 3
            rsISS.Fields(idx) = rsERP.Fields(idx)
        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
    
    rsERP.Close
    rsISS.Close
    
    Write_Log (": CustomerGroupMaster Finished. . .")

    'Item Group
    Write_Log (": ItemGroupMaster Start. . .")

    conISS.Execute "DELETE FROM ItemGroupMaster WHERE Company = '30'"

    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM ItemGroupMaster "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP

    mCmd.CommandText = "SELECT DISTINCT '30', typcod, typdes, 0, 1 " _
                                                        + " FROM " + strERPPath + "\ISTAB " _
                                                        + " WHERE tabtyp = '22' " _
                                                        + " UNION " _
                                                        + " SELECT DISTINCT '30', typcod, typdes, 0, 1 " _
                                                        + " FROM " + strERP1Path + "\ISTAB " _
                                                        + " WHERE tabtyp = '22' " _
                                                        + " AND typcod not in (SELECT typcod " _
                                                        + " FROM " + strERPPath + "\ISTAB " _
                                                        + " WHERE tabtyp = '22' )"
    mCmd.CommandType = adCmdText

   Set rsERP = mCmd.Execute
   
       If Not rsERP.EOF Then
        rsERP.MoveFirst
    End If
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 4
            rsISS.Fields(idx) = rsERP.Fields(idx)
        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
    
    rsERP.Close
    rsISS.Close

    Write_Log (": ItemGroupMaster Finished. . .")

'    Product Group
    Write_Log (": ProductGroup Start. . .")


    conISS.Execute "DELETE FROM ProductGroup WHERE Company = '30'"

    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM ProductGroup "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP

    mCmd.CommandText = "SELECT DISTINCT '30', shortnam2, 'F', 1 " _
                                                        + " FROM " + strERPPath + "\ISTAB " _
                                                        + " WHERE tabtyp = '22' " _
                                                        + " UNION " _
                                                        + " SELECT DISTINCT '30', shortnam2, 'F', 1 " _
                                                        + " FROM " + strERP1Path + "\ISTAB " _
                                                        + " WHERE tabtyp = '22' " _
                                                        + " AND shortnam2 not in (SELECT shortnam2 " _
                                                        + " FROM " + strERPPath + "\ISTAB " _
                                                        + " WHERE tabtyp = '22') "
    mCmd.CommandType = adCmdText

    Set rsERP = mCmd.Execute
    
        If Not rsERP.EOF Then
        rsERP.MoveFirst
    End If
  

    Do While Not rsERP.EOF
        rsISS.AddNew

        For idx = 0 To 3
            rsISS.Fields(idx) = rsERP.Fields(idx)
        Next idx

        rsISS.Update
        rsERP.MoveNext

    Loop

    rsERP.Close
    rsISS.Close

    Write_Log (": ProductGroup Finished. . .")

'    Product Line
    Write_Log (": ProductLine Start. . .")

    conISS.Execute "DELETE FROM ProductLine WHERE Company = '30'"

'    Set rsISS = New ADODB.Recordset
'    strSQL = "SELECT * FROM ProductLine "
'    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic
'
'    Set rsERP = New ADODB.Recordset
'    Set mCmd = New ADODB.Command
'    mCmd.ActiveConnection = conERP
'
'    mCmd.CommandText = "SELECT DISTINCT '30', shortnam2, 'F', 1 " _
'                                                        + " FROM ISTAB " _
'                                                        + " WHERE tabtyp = '22' " _
'                                                        + " UNION " _
'                                                        + " SELECT DISTINCT '30', shortnam2, 'F', 1 " _
'                                                        + " FROM " + strERP1Path + "\ISTAB " _
'                                                        + " WHERE tabtyp = '22' " _
'                                                        + " AND shortnam2 not in (SELECT shortnam2 " _
'                                                        + " FROM " + strERPPath + "\ISTAB " _
'                                                        + " WHERE tabtyp = '22')"
'    mCmd.CommandType = adCmdText
'
'    Set rsERP = mCmd.Execute
'
'        If Not rsERP.EOF Then
'        rsERP.MoveFirst
'    End If
'
'    Do While Not rsERP.EOF
'        rsISS.AddNew
'
'        For idx = 0 To 3
'            rsISS.Fields(idx) = rsERP.Fields(idx)
'        Next idx
'
'        rsISS.Update
'        rsERP.MoveNext
'
'    Loop
    
    conISS.Execute "INSERT INTO ProductLine VALUES('30', 'PCK', 'F', '1')"
     conISS.Execute "INSERT INTO ProductLine VALUES('30', 'WAT', 'F', '1')"
 
    
    Write_Log (": ProductLine Finished. . .")

    'ProductMaster
    Write_Log (": ProductMaster Start. . .")

    conISS.Execute "DELETE FROM ProductMaster WHERE Company = '30'"

    'PCK
    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM ProductMaster "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic

    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP1

    ' For single selling unit
    mCmd.CommandText = "SELECT '30', 'PCK'+RTRIM(stkcod), stkdes, stkgrp, '', tab2.typdes, tab3.typdes, tab4.typdes, '', 1, 0, 0, 0, sellpr1, 0, 0, 0, ISINFO.vatrat, tab1.shortnam2 AS ProductGroup, 'PCK', '', 1 " _
                                                    + " FROM " + strERPPath + "\STMAS, " + strERPPath + "\ISTAB tab1, " + strERPPath + "\ISTAB tab2, " + strERPPath + "\ISINFO, " + strERPPath + "\ISTAB tab3, " + strERPPath + "\ISTAB tab4 " _
                                                    + " WHERE tab1.tabtyp = '22' " _
                                                    + " AND tab1.typcod = stkgrp " _
                                                    + " AND tab2.tabtyp = '20' " _
                                                    + " AND tab2.typcod = qucod " _
                                                    + " AND tab3.tabtyp = '20' " _
                                                    + " AND tab3.typcod = pqucod " _
                                                    + " AND tab4.tabtyp = '20' " _
                                                    + " AND tab4.typcod = cqucod " _
                                                    + " UNION " _
                                                    + " SELECT '30', 'PCK'+RTRIM(stkcod), stkdes, stkgrp, '', tab2.typdes, tab3.typdes, '', '', 1, 0, 0, 0, sellpr1, 0, 0, 0, ISINFO.vatrat, tab1.shortnam2 AS ProductGroup, 'PCK', '', 1 " _
                                                    + " FROM " + strERPPath + "\STMAS, " + strERPPath + "\ISTAB tab1, " + strERPPath + "\ISTAB tab2, " + strERPPath + "\ISINFO, " + strERPPath + "\ISTAB tab3 " _
                                                    + " WHERE tab1.tabtyp = '22' " _
                                                    + " AND tab1.typcod = stkgrp " _
                                                    + " AND tab2.tabtyp = '20' " _
                                                    + " AND tab2.typcod = qucod " _
                                                    + " AND tab3.tabtyp = '20' " _
                                                    + " AND tab3.typcod = pqucod " _
                                                    + " AND stkcod not in (SELECT stkcod FROM " + strERPPath + "\STMAS, " + strERPPath + "\ISTAB tab4 where tab4.tabtyp = '20' " _
                                                    + " AND tab4.typcod = cqucod )"
'                                                    + " AND stkcod not in (SELECT stkcod FROM " + strERPPath + "\STMAS ) "

        

    mCmd.CommandType = adCmdText

          
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst

    Do While Not rsERP.EOF
        rsISS.AddNew

        For idx = 0 To 21
            rsISS.Fields(idx) = rsERP.Fields(idx)
        Next idx
        
        
        
        If (rsERP.Fields(6)) = (rsERP.Fields(5)) Then
            rsISS.Fields(6) = " "
        End If

        If (rsERP.Fields(7)) = (rsERP.Fields(5)) Then
            rsISS.Fields(7) = " "
        End If

        If (rsERP.Fields(7)) = (rsERP.Fields(6)) Then
            rsISS.Fields(7) = " "
        End If
        
        If (rsISS.Fields(6)) = " " Then
        rsISS.Fields(6) = rsISS.Fields(7)
        
        rsISS.Fields(7) = " "
        End If
        
        
        rsISS.Update
        rsERP.MoveNext

    Loop

    rsERP.Close
    
    'WAT
        Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM ProductMaster "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic

    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP1

    ' For single selling unit
     mCmd.CommandText = "SELECT '30', 'WAT'+RTRIM(stkcod), stkdes, stkgrp, '', tab2.typdes, tab3.typdes, tab4.typdes, '', 1, 0, 0, 0, sellpr1, 0, 0, 0, ISINFO.vatrat, tab1.shortnam2 AS ProductGroup, 'WAT', '', 1 " _
                                                    + " FROM " + strERP1Path + "\STMAS, " + strERP1Path + "\ISTAB tab1, " + strERP1Path + "\ISTAB tab2, " + strERP1Path + "\ISINFO, " + strERP1Path + "\ISTAB tab3, " + strERP1Path + "\ISTAB tab4 " _
                                                    + " WHERE tab1.tabtyp = '22' " _
                                                    + " AND tab1.typcod = stkgrp " _
                                                    + " AND tab2.tabtyp = '20' " _
                                                    + " AND tab2.typcod = qucod " _
                                                    + " AND tab3.tabtyp = '20' " _
                                                    + " AND tab3.typcod = pqucod " _
                                                    + " AND tab4.tabtyp = '20' " _
                                                    + " AND tab4.typcod = cqucod " _
                                                    + " UNION " _
                                                    + " SELECT '30', 'WAT'+RTRIM(stkcod), stkdes, stkgrp, '', tab2.typdes, tab3.typdes, '', '', 1, 0, 0, 0, sellpr1, 0, 0, 0, ISINFO.vatrat, tab1.shortnam2 AS ProductGroup, 'WAT', '', 1 " _
                                                    + " FROM " + strERP1Path + "\STMAS, " + strERP1Path + "\ISTAB tab1, " + strERP1Path + "\ISTAB tab2, " + strERP1Path + "\ISINFO, " + strERP1Path + "\ISTAB tab3 " _
                                                    + " WHERE tab1.tabtyp = '22' " _
                                                    + " AND tab1.typcod = stkgrp " _
                                                    + " AND tab2.tabtyp = '20' " _
                                                    + " AND tab2.typcod = qucod " _
                                                    + " AND tab3.tabtyp = '20' " _
                                                    + " AND tab3.typcod = pqucod " _
                                                    + " AND stkcod not in (SELECT stkcod FROM " + strERP1Path + "\STMAS, " + strERP1Path + "\ISTAB tab4 where tab4.tabtyp = '20' " _
                                                    + " AND tab4.typcod = cqucod )"
'                                                    + " AND stkcod not in (SELECT stkcod FROM " + strERPPath + "\STMAS ) "

        

    mCmd.CommandType = adCmdText

          
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst

    Do While Not rsERP.EOF
        rsISS.AddNew

        For idx = 0 To 21
            rsISS.Fields(idx) = rsERP.Fields(idx)
        Next idx
        
        
        
        If (rsERP.Fields(6)) = (rsERP.Fields(5)) Then
            rsISS.Fields(6) = " "
        End If

        If (rsERP.Fields(7)) = (rsERP.Fields(5)) Then
            rsISS.Fields(7) = " "
        End If

        If (rsERP.Fields(7)) = (rsERP.Fields(6)) Then
            rsISS.Fields(7) = " "
        End If
        
        If (rsISS.Fields(6)) = " " Then
        rsISS.Fields(6) = rsISS.Fields(7)
        
        rsISS.Fields(7) = " "
        End If

        rsISS.Update
        rsERP.MoveNext

    Loop
    
    
    
    rsERP.Close
    rsISS.Close
   
    


    Write_Log (": ProductMaster Finished. . .")
    
    'Province
     Write_Log (": Province Start. . .")
     
    conISS.Execute "DELETE FROM Province WHERE Company = '30'"
    
    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM Province "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic

    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
    
    mCmd.CommandText = "SELECT '30', typcod, typdes, 1 " _
                                                        + " FROM  " + strERPPath + "\ISTAB " _
                                                        + " WHERE tabtyp = '40' " _
                                                       
                                                        
    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    intPrvCode = 0
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 3
                rsISS.Fields(idx) = rsERP.Fields(idx)
        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
    
    rsERP.Close
    rsISS.Close
    
    conISS.Execute "INSERT INTO Province VALUES('30', 'UAS', 'Unassigned', 1) "
    Write_Log (": Province Finished. . .")


    'Salesman
    Write_Log (": Salesman Start. . .")
    
    conISS.Execute "DELETE FROM Salesman WHERE Company = '30'"
    
    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM Salesman "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
    
    mCmd.CommandText = "SELECT '30', 'PCK'+RTRIM(slmcod), 'PCK'+RTRIM(slmtyp), slmnam, 'PCK', 'ALL', 'ALL', 1 " _
                                                        + " FROM OESLM " _
                                                        + " WHERE slmtyp <> '' "
    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    
    
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 7
            rsISS.Fields(idx) = rsERP.Fields(idx)
        Next idx
        
       
        rsISS.Fields(1) = RTrim(rsERP.Fields(1))
        rsISS.Fields(2) = RTrim(rsERP.Fields(2))
        
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
    
    rsERP.Close
    
    
    conISS.Execute "        INSERT INTO Salesman" _
                            + "  SELECT company, salesmancode+'..' , '" & Mgrcode1 & "'," _
                            + "  SalesmanName , SalesOrganization, DistributionChannel, SalesAreaCode, LangId " _
                            + "  From Salesman " _
                            + "  Where Company ='30' " _
                            + "  And SalesmanCode <> '" & Mgrcode1 & "' " _
                            + "  And ParentSalesmanCode <> ' ' " _
                            + "  And SalesmanCode NOT IN (Select Distinct ParentSalesmanCode " _
                            + "                         From Salesman " _
                            + "                         Where company = '30' " _
                            + "                          )"

    conISS.Execute strSQL
    
    
    rsISS.Close
    
    Write_Log (": Salesman Finished. . .")

    'OrderStatusReport
    Write_Log (": OrderStatusReport Start. . .")
    
    
    'Back Date 12 months
    Date1 = DateAdd("m", -12, Date)
    Date1 = DateAdd("yyyy", -543, Date1)
    
    
    conISS.Execute "DELETE FROM OrderStatusReport WHERE Company = '30'"
    'pck
    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM OrderStatusReport "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
    
    mCmd.CommandText = "SELECT DISTINCT '30'AS Company, ARTRN.slmcod AS SalesmanCode, 'PCK'+ARTRN.cuscod AS CustomerCode, ARTRN.youref AS PO, " _
                                                                        + " '' AS PODate, ARTRN.youref AS PORef, 'C' AS DocumentCategory, ARTRN.sonum AS SO, " _
                                                                        + " DTOS(ARTRN.docdat) AS SODate, '' AS DO, '' AS DODate, " _
                                                                        + " ARTRN.docnum AS Bill, DTOS(ARTRN.docdat) AS BillDate, " _
                                                                        + " '' AS DDPDate, 'THB' " _
                                                        + " FROM " + strERPPath + "\ARTRN INNER JOIN " + strERPPath + "\ISRUN  ON SUBSTR(ARTRN.docnum,1,2) = ISRUN.prefix  " _
                                                        + " WHERE ISRUN.doctyp IN ('IB','IV') " _
                                                        + " AND ARTRN.slmcod <> ' ' " _
                                                        + " AND ARTRN.docdat >=  {" + CStr(Date1) + "}" _
                                                        + " UNION " _
                                                        + "SELECT DISTINCT '30'AS Company, ARTRN.slmcod AS SalesmanCode, 'PCK'+ARTRN.cuscod AS CustomerCode, ARTRN.youref AS PO, " _
                                                                        + " '' AS PODate, ARTRN.youref AS PORef, 'C' AS DocumentCategory, ARTRN.sonum AS SO, " _
                                                                        + " DTOS(ARTRN.docdat) AS SODate, '' AS DO, '' AS DODate, " _
                                                                        + " '' AS Bill, '' AS BillDate, " _
                                                                        + " '' AS DDPDate, 'THB' " _
                                                        + " FROM " + strERPPath + "\ARTRN " _
                                                        + " WHERE ARTRN.docnum LIKE 'SR%' " _
                                                        + " AND ARTRN.slmcod <> ' ' " _
                                                        + " AND ARTRN.docdat >=  {" + CStr(Date1) + "}" _

                                                        

    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 14
            If idx = 1 Then
                rsISS.Fields(idx) = "PCK" & Trim(rsERP.Fields(idx))
            Else
                rsISS.Fields(idx) = rsERP.Fields(idx)
            End If
        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
    
    rsERP.Close
    
    'WAT
     Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM OrderStatusReport "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
    
    mCmd.CommandText = "SELECT DISTINCT '30'AS Company, ARTRN.slmcod AS SalesmanCode, 'PCK'+ARTRN.cuscod AS CustomerCode, ARTRN.youref AS PO, " _
                                                                        + " '' AS PODate, ARTRN.youref AS PORef, 'C' AS DocumentCategory, ARTRN.sonum AS SO, " _
                                                                        + " DTOS(ARTRN.docdat) AS SODate, '' AS DO, '' AS DODate, " _
                                                                        + " ARTRN.docnum AS Bill, DTOS(ARTRN.docdat) AS BillDate, " _
                                                                        + " '' AS DDPDate, 'THB' " _
                                                        + " FROM " + strERP1Path + "\ARTRN INNER JOIN " + strERPPath + "\ISRUN  ON SUBSTR(ARTRN.docnum,1,2) = ISRUN.prefix  " _
                                                        + " WHERE ISRUN.doctyp IN ('IB','IV') " _
                                                        + " AND ARTRN.slmcod <> ' ' " _
                                                        + " AND ARTRN.docdat >=  {" + CStr(Date1) + "}" _
                                                        + " UNION " _
                                                        + "SELECT DISTINCT '30'AS Company, ARTRN.slmcod AS SalesmanCode, 'PCK'+ARTRN.cuscod AS CustomerCode, ARTRN.youref AS PO, " _
                                                                        + " '' AS PODate, ARTRN.youref AS PORef, 'C' AS DocumentCategory, ARTRN.sonum AS SO, " _
                                                                        + " DTOS(ARTRN.docdat) AS SODate, '' AS DO, '' AS DODate, " _
                                                                        + " '' AS Bill, '' AS BillDate, " _
                                                                        + " '' AS DDPDate, 'THB' " _
                                                        + " FROM " + strERP1Path + "\ARTRN " _
                                                        + " WHERE ARTRN.docnum LIKE 'SR%' " _
                                                        + " AND ARTRN.slmcod <> ' ' " _
                                                        + " AND ARTRN.docdat >=  {" + CStr(Date1) + "}" _

                                                        

    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 14
            If idx = 1 Then
                rsISS.Fields(idx) = "PCK" & Trim(rsERP.Fields(idx))
            Else
                rsISS.Fields(idx) = rsERP.Fields(idx)
            End If
        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
    
    rsERP.Close
    conISS.Execute "UPDATE OrderStatusReport SET SalesmanCode = RTRIM(SalesmanCode) "
    
    Write_Log (": OrderStatusReport Copy (Slmcode+.) for MgrCode. . .")

    strSQL = "          INSERT INTO OrderStatusReport"
    strSQL = strSQL & " SELECT Company ,RTRIM(SalesmanCode)+'.' , CustomerCode, PO, PODate, PORef, DocumentCategory, SO,"
    strSQL = strSQL & "        SODate, DO, DODate, Bill, BillDate, DDPDate, CurrencyCode"
    strSQL = strSQL & "   From OrderStatusReport"
    strSQL = strSQL & "   Where Company ='30'"
    strSQL = strSQL & "   and SalesmanCode <> '" & Mgrcode & "'"
    conISS.Execute strSQL, intNoOfRow
    
     Write_Log (": OrderStatusReport Copy (Slmcode+..) for MgrCode1. . .")

    strSQL = "          INSERT INTO OrderStatusReport"
    strSQL = strSQL & " SELECT Company ,RTRIM(SalesmanCode)+'..' , CustomerCode, PO, PODate, PORef, DocumentCategory, SO,"
    strSQL = strSQL & "        SODate, DO, DODate, Bill, BillDate, DDPDate, CurrencyCode"
    strSQL = strSQL & "   From OrderStatusReport"
    strSQL = strSQL & "   Where Company ='30'"
    strSQL = strSQL & "   and SalesmanCode <> '" & Mgrcode1 & "'"
    strSQL = strSQL & "   and SalesmanCode not like '%.'"
    conISS.Execute strSQL, intNoOfRow
    
    Write_Log (": OrderStatusReport Finished. . .")
    
    'OrderStatusReportItem
    Write_Log (": OrderStatusReportItem Start. . .")

    conISS.Execute "DELETE FROM OrderStatusReportItem WHERE Company = '30'"
    'pck
    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM OrderStatusReportItem "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP

    mCmd.CommandText = "SELECT '30'AS Company, OESO.slmcod AS SalesmanCode, 'PCK'+OESO.cuscod AS CustomerCode, ARTRN.youref AS PO, " _
                                                                        + " 'C' AS DocumentCategory, ARTRN.sonum AS SO, '' AS DO, seqnum AS SOItemNo, " _
                                                                        + " 'PCK'+RTRIM(stkcod) AS ProductCode, ordqty AS OrderQty, shortnam AS OrderUnit, " _
                                                                        + " trnval AS OrderAmount, cancelqty AS CancelQty, trnval AS ShipAmount, " _
                                                                        + " remqty AS LeftQty, trnval  AS LeftAmount, " _
                                                                        + " canceltyp AS CancelRemark, 'THB' " _
                                                        + " From " + strERPPath + "\OESO, " + strERPPath + "\ISTAB, " _
                                                        + " " + strERPPath + "\OESOIT inner join " + strERPPath + "\ARTRN on OESOIT.sonum = ARTRN.sonum " _
                                                        + " WHERE OESO.sonum = OESOIT.sonum " _
                                                        + " AND ISTAB.tabtyp = '20' " _
                                                        + " AND ISTAB.typcod = tqucod " _
                                                        + " AND OESO.slmcod <> ' ' " _
                                                        + " AND OESO.sonum <> ' ' " _
                                                        + " AND OESO.sodat >=  {" + CStr(Date1) + "}"
                                                       
                                                        
    mCmd.CommandType = adCmdText

    Set rsERP = mCmd.Execute
    rsERP.MoveFirst

    Do While Not rsERP.EOF
        rsISS.AddNew

        For idx = 0 To 17
            If idx = 1 Then
                rsISS.Fields(idx) = "PCK" & Trim(rsERP.Fields(idx))
            Else
                rsISS.Fields(idx) = rsERP.Fields(idx)
            End If
        Next idx

        curTotalAmount = rsISS.Fields(11)
        intOrderQty = rsISS.Fields(9)
        intCancelQty = rsISS.Fields(12)
        intRemQty = rsISS.Fields(14)

        rsISS.Fields(9) = intOrderQty - intCancelQty
        rsISS.Fields(11) = curTotalAmount * ((intOrderQty - intCancelQty) / intOrderQty)
        rsISS.Fields(12) = intOrderQty - intCancelQty - intRemQty
        rsISS.Fields(13) = curTotalAmount * ((intOrderQty - intCancelQty - intRemQty) / intOrderQty)
        rsISS.Fields(15) = curTotalAmount * (intRemQty / intOrderQty)

        rsISS.Update
        rsERP.MoveNext

    Loop

    rsERP.Close
    
     'WAT
    Set rsISS = New ADODB.Recordset
    strSQL = "SELECT * FROM OrderStatusReportItem "
    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic


    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP

    mCmd.CommandText = "SELECT '30'AS Company, OESO.slmcod AS SalesmanCode, 'WAT'+OESO.cuscod AS CustomerCode, ARTRN.youref AS PO, " _
                                                                        + " 'C' AS DocumentCategory, ARTRN.sonum AS SO, '' AS DO, seqnum AS SOItemNo, " _
                                                                        + " 'WAT'+RTRIM(stkcod) AS ProductCode, ordqty AS OrderQty, shortnam AS OrderUnit, " _
                                                                        + " trnval AS OrderAmount, cancelqty AS CancelQty, trnval AS ShipAmount, " _
                                                                        + " remqty AS LeftQty, trnval  AS LeftAmount, " _
                                                                        + " canceltyp AS CancelRemark, 'THB' " _
                                                        + " From " + strERP1Path + "\OESO, " + strERP1Path + "\ISTAB, " _
                                                        + " " + strERP1Path + "\OESOIT inner join " + strERP1Path + "\ARTRN on OESOIT.sonum = ARTRN.sonum " _
                                                        + " WHERE OESO.sonum = OESOIT.sonum " _
                                                        + " AND ISTAB.tabtyp = '20' " _
                                                        + " AND ISTAB.typcod = tqucod " _
                                                        + " AND OESO.slmcod <> ' ' " _
                                                        + " AND OESO.sonum <> ' ' " _
                                                        + " AND OESO.sodat >=  {" + CStr(Date1) + "}"
                                                       
                                                        
    mCmd.CommandType = adCmdText

    Set rsERP = mCmd.Execute
    rsERP.MoveFirst

    Do While Not rsERP.EOF
        rsISS.AddNew

        For idx = 0 To 17
            If idx = 1 Then
                rsISS.Fields(idx) = "PCK" & Trim(rsERP.Fields(idx))
            Else
                rsISS.Fields(idx) = rsERP.Fields(idx)
            End If
        Next idx

        curTotalAmount = rsISS.Fields(11)
        intOrderQty = rsISS.Fields(9)
        intCancelQty = rsISS.Fields(12)
        intRemQty = rsISS.Fields(14)

        rsISS.Fields(9) = intOrderQty - intCancelQty
        rsISS.Fields(11) = curTotalAmount * ((intOrderQty - intCancelQty) / intOrderQty)
        rsISS.Fields(12) = intOrderQty - intCancelQty - intRemQty
        rsISS.Fields(13) = curTotalAmount * ((intOrderQty - intCancelQty - intRemQty) / intOrderQty)
        rsISS.Fields(15) = curTotalAmount * (intRemQty / intOrderQty)

        rsISS.Update
        rsERP.MoveNext

    Loop

    rsERP.Close

    'OrderStatusReportItem (Header Comment)

'    Write_Log (": OrderStatusReportItem-Header Comment Start. . .")
'
'    Set rsISS = New ADODB.Recordset
'    strSQL = "SELECT * FROM OrderStatusReportItem "
'    rsISS.Open strSQL, conISS, adOpenDynamic, adLockOptimistic
'
'
'    Set rsERP = New ADODB.Recordset
'    Set mCmd = New ADODB.Command
'    mCmd.ActiveConnection = conERP
'
'    mCmd.CommandText = "SELECT '30'AS Company, OESO.slmcod AS SalesmanCode, 'PCK'+OESO.cuscod AS CustomerCode, OESO.youref AS PO, " _
'                                                                        + " 'C' AS DocumentCategory, OESO.sonum AS SO, '' AS DO, ARTRNRM.seqnum AS SOItemNo, " _
'                                                                        + " '' AS ProductCode, 0 AS OrderQty, '' AS OrderUnit, " _
'                                                                        + " 0 AS OrderAmount, 0 AS CancelQty, 0 AS ShipAmount, " _
'                                                                        + " 0 AS LeftQty, 0  AS LeftAmount, " _
'                                                                        + " ARTRNRM.remark AS CancelRemark, 'THB' " _
'                                                        + " From OESO, ARTRNRM " _
'                                                        + " WHERE OESO.sonum = ARTRNRM.docnum " _
'                                                        + " AND OESO.slmcod <> ' ' " _
'                                                        + " AND OESO.sonum <> ' ' " _
'                                                        + " AND OESO.sodat >=  {" + CStr(Date1) + "} " _
'                                                        + " AND ARTRNRM.remark <> ' ' "
'    mCmd.CommandType = adCmdText
'
'    Set rsERP = mCmd.Execute
'    rsERP.MoveFirst
'
'    Do While Not rsERP.EOF
'        rsISS.AddNew
'
'        For idx = 0 To 17
'            If idx = 1 Then
'                rsISS.Fields(idx) = "PCK" & Trim(rsERP.Fields(idx))
'            Else
'                rsISS.Fields(idx) = rsERP.Fields(idx)
'            End If
'
'            'Plus 900 because the comment will be display at the bottom
'            If idx = 7 Then
'                If Left(rsISS.Fields(idx), 1) = "@" Then
'                    rsISS.Fields(idx) = Mid(rsISS.Fields(idx), 2, 10)
'                End If
'                rsISS.Fields(idx) = CStr(Int(rsISS.Fields(idx)) + 900)
'            End If
'
'            If idx = 16 Then
'                rsISS.Fields(idx) = Trim(rsISS.Fields(idx))
'            End If
'        Next idx
'
'        rsISS.Update
'        rsERP.MoveNext
'
'    Loop
'
'    rsERP.Close

    Write_Log (": OrderStatusReportItem-Header Comment Finished. . .")
        
    conISS.Execute "UPDATE OrderStatusReportItem SET SalesmanCode = RTRIM(SalesmanCode) "
    
    Write_Log (": OrderStatusReportItem Copy (Slmcode+.) for MgrCode. . .")

    strSQL = "         INSERT INTO OrderStatusReportItem"
    strSQL = strSQL & "   SELECT Company, SalesmanCode+'.' ,CustomerCode,PO,DocumentCategory,"
    strSQL = strSQL & "     SO,DO,SOItemNo,ProductCode,OrderQty,OrderUnit,OrderAmount,"
    strSQL = strSQL & "     ShipQty , ShipAmount, LeftQty, LeftAmount, CancelRemark, CurrencyCode, productsetcode "
    strSQL = strSQL & "   From OrderStatusReportItem"
    strSQL = strSQL & "   Where Company ='30'"
    strSQL = strSQL & "   and SalesmanCode <> '" & Mgrcode & "'"
    conISS.Execute strSQL, intNoOfRow
    
    Write_Log (": OrderStatusReportItem Copy (Slmcode+..) for MgrCode1. . .")

    strSQL = "         INSERT INTO OrderStatusReportItem"
    strSQL = strSQL & "   SELECT Company, SalesmanCode+'..' ,CustomerCode,PO,DocumentCategory,"
    strSQL = strSQL & "     SO,DO,SOItemNo,ProductCode,OrderQty,OrderUnit,OrderAmount,"
    strSQL = strSQL & "     ShipQty , ShipAmount, LeftQty, LeftAmount, CancelRemark, CurrencyCode, productsetcode "
    strSQL = strSQL & "   From OrderStatusReportItem"
    strSQL = strSQL & "   Where Company ='30'"
    strSQL = strSQL & "   and SalesmanCode <> '" & Mgrcode1 & "'"
    strSQL = strSQL & "   and SalesmanCode not like '%.'"
    conISS.Execute strSQL, intNoOfRow
    
    Write_Log (": Plant Start. . .")
    
    conISS.Execute "DELETE FROM Plant WHERE Company = '30'"
                                           
    Set rsISS = New ADODB.Recordset
    rsISS.Open "SELECT * FROM Plant ", conISS, adOpenDynamic, adLockOptimistic

    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
    
    mCmd.CommandText = "SELECT '30'AS Company, typcod AS Plant, typdes AS PlantName, 'NOR', '1' " _
                                                    + " FROM " + strERPPath + "\ISTAB " _
                                                    + " WHERE tabtyp = '21'" _
                                                    + " UNION " _
                                                    + " SELECT '30'AS Company, typcod AS Plant, typdes AS PlantName, 'NOR', '1' " _
                                                    + " FROM " + strERP1Path + "\ISTAB " _
                                                    + " WHERE tabtyp = '21' "
                                                    
    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 4
            rsISS.Fields(idx) = rsERP.Fields(idx)
        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
    
    rsISS.Close
    rsERP.Close
    
    Write_Log (": Plant Finish. . .")
    
    Write_Log (": LookUp Start. . .")
    
    conISS.Execute "DELETE FROM LookUp WHERE Company = '30' and LookupType = 'ModeDly' "
                                           
    Set rsISS = New ADODB.Recordset
    rsISS.Open "SELECT * FROM LookUp ", conISS, adOpenDynamic, adLockOptimistic

    Set rsERP = New ADODB.Recordset
    Set mCmd = New ADODB.Command
    mCmd.ActiveConnection = conERP
    
    mCmd.CommandText = "SELECT '30'AS Company, 'ModeDly', 'PCK'+TRIM(typcod) AS DlyCod, 'PCK-'+TRIM(typdes) AS DlyDes, '1' " _
                                                    + " FROM " + strERPPath + "\ISTAB " _
                                                    + " WHERE tabtyp = '41'" _
                                                    + " UNION " _
                                                    + " SELECT '30'AS Company, 'ModeDly', 'WAT'+TRIM(typcod) AS DlyCod, 'WAT-'+TRIM(typdes) AS DlyDes, '1'" _
                                                    + " FROM " + strERP1Path + "\ISTAB " _
                                                    + " WHERE tabtyp = '41' "
                                                    
    mCmd.CommandType = adCmdText
    
    Set rsERP = mCmd.Execute
    rsERP.MoveFirst
    
    Do While Not rsERP.EOF
        rsISS.AddNew
        
        For idx = 0 To 4
            rsISS.Fields(idx) = rsERP.Fields(idx)
        Next idx
        
        rsISS.Update
        rsERP.MoveNext
            
    Loop
    
    rsISS.Close
    rsERP.Close
    
    Write_Log (": LookUp Finish. . .")
    
    Write_Log (": SingleMaster Finished. . .")
    
End Sub

Public Sub Write_Log(pstrLogMsg As String)

    Open strWorkingDir & "\PCK2ISS_Log.txt" For Append As #fn
    Write #fn, Now & pstrLogMsg
    Close #fn
    
End Sub


Public Function SendMail(sTo As String, sSubject As String, sFrom As String, _
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
    lobj_cdomsg.To = sTo
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


