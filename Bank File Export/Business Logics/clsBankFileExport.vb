Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsBankFileExport
    Inherits clsBase
    Private strQuery As String
    Private oGrid As SAPbouiCOM.Grid
    Private oCombo As SAPbouiCOM.ComboBox
    Private oDt_Import As SAPbouiCOM.DataTable
    Private oDt_ErrorLog As SAPbouiCOM.DataTable
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oEditColumn As SAPbouiCOM.EditTextColumn
    Private oRecordSet As SAPbobsCOM.Recordset
    Private Const log_PROCESS_ORDERS As String = "Log_ProcessOrders.txt"
    Private Const log_INVOICING As String = "Log_Invoicing.txt"

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Pay_ERBF, frm_Pay_ERBF)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            Initialize(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub Initialize(ByVal aForm As SAPbouiCOM.Form)
        Try
            Dim sQuery As String
            aForm.DataSources.UserDataSources.Add("dtYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            aForm.DataSources.UserDataSources.Add("dtMonth", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            aForm.DataSources.UserDataSources.Add("Bank", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            aForm.DataSources.UserDataSources.Add("Cmp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            aForm.DataSources.UserDataSources.Add("Type", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            aForm.DataSources.UserDataSources.Add("File", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            aForm.DataSources.UserDataSources.Add("Currency", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            aForm.DataSources.UserDataSources.Add("Tab", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            Dim ocheck As SAPbouiCOM.CheckBox
            ocheck = aForm.Items.Item("18").Specific
            ocheck.DataBind.SetBound(True, "", "Tab")
            oCombo = aForm.Items.Item("17").Specific
            oCombo.DataBind.SetBound(True, "", "Currency")
            oCombo = aForm.Items.Item("11").Specific
            oCombo.DataBind.SetBound(True, "", "Cmp")
            oCombo = aForm.Items.Item("13").Specific
            oCombo.DataBind.SetBound(True, "", "Type")
            oCombo = aForm.Items.Item("4").Specific
            oCombo.DataBind.SetBound(True, "", "Bank")
            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCombo.ValidValues.Add("", "")
            oTest.DoQuery("Select Account,AcctName from DSC1 order by BankCode")
            For intRow As Integer = 0 To oTest.RecordCount - 1
                Try
                    oCombo.ValidValues.Add(oTest.Fields.Item(0).Value, oTest.Fields.Item(1).Value)
                Catch ex As Exception

                End Try
                oTest.MoveNext()
            Next
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            oCombo = aForm.Items.Item("17").Specific
            oCombo.ValidValues.Add("", "All")
            oCombo.ValidValues.Add("SA", "Local Currency")
            oCombo.ValidValues.Add("US", "USD")
            oCombo.ValidValues.Add("GB", "GBP")
            oCombo.ValidValues.Add("KW", "KWD")
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            oCombo = aForm.Items.Item("11").Specific
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCombo.ValidValues.Add("", "All")
            oTest.DoQuery("SELECT T0.[U_Z_CompCode], T0.[U_Z_CompName] FROM [dbo].[@Z_OADM]  T0 order by U_Z_CompCode")
            For intRow As Integer = 0 To oTest.RecordCount - 1
                Try
                    oCombo.ValidValues.Add(oTest.Fields.Item(0).Value, oTest.Fields.Item(1).Value)
                Catch ex As Exception

                End Try
                oTest.MoveNext()
            Next
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            oCombo = oForm.Items.Item("13").Specific
            oCombo.ValidValues.Add("", "Both")
            oCombo.ValidValues.Add("R", "Regular")
            oCombo.ValidValues.Add("O", "Offcycle")
            oCombo.ValidValues.Add("T", "Offcycle Transaction") '<<========================================================Edited By Houssam=============================<<
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            oCombo = aForm.Items.Item("6").Specific
            oCombo.DataBind.SetBound(True, "", "dtYear")
            oCombo.ValidValues.Add("", "")
            For intRow As Integer = Now.Year - 5 To Now.Year + 5
                oCombo.ValidValues.Add(intRow, intRow)
            Next
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oCombo = aForm.Items.Item("8").Specific
            oCombo.DataBind.SetBound(True, "", "dtMonth")
            oCombo.ValidValues.Add("", "")
            For intRow As Integer = 1 To 12
                oCombo.ValidValues.Add(intRow, MonthName(intRow))
            Next
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

            oCombo = aForm.Items.Item("15").Specific
            oCombo.DataBind.SetBound(True, "", "File")
            oCombo.ValidValues.Add("T", "Text File")
            oCombo.ValidValues.Add("C", "CSV File")
            oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function Processing(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strQuery As String
        Dim oRec As SAPbobsCOM.Recordset
        Dim intYear, intMonth As Integer
        Dim strBank, strCompany, strType As String
        oCombo = aForm.Items.Item("4").Specific
        If oCombo.Selected.Value = "" Then
            oApplication.Utilities.Message("House bank detail missing..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            strBank = oCombo.Selected.Value
        End If
        oCombo = aForm.Items.Item("6").Specific
        If oCombo.Selected.Value = "" Then
            oApplication.Utilities.Message("Select Payroll Year..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            intYear = oCombo.Selected.Value
        End If

        oCombo = aForm.Items.Item("8").Specific
        If oCombo.Selected.Value = "" Then
            oApplication.Utilities.Message("Select Payroll Month..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            intMonth = oCombo.Selected.Value
        End If


        oCombo = aForm.Items.Item("11").Specific
        strCompany = oCombo.Selected.Value

        oCombo = aForm.Items.Item("13").Specific
        strType = oCombo.Selected.Value
        Dim strTypeValue As String = strType

        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strQuery = "Select * from [@Z_PAYROLL1] where U_Z_Posted='N' and U_Z_Year=" & intYear & " and U_Z_Month=" & intMonth

        If strCompany = "" Then
            strCompany = "1=1"
        Else
            strCompany = "T0.[U_Z_CompNo]='" & strCompany & "'"
        End If

        If strType = "" Then
            strType = "1=1"
        ElseIf strType = "R" Then
            strType = "T0.[U_Z_OffCycle]='N'"
        ElseIf strType = "O" Then '<<================================================Added By Houssam============================================<<
            strType = "T0.[U_Z_OffCycle]='Y'"
        End If

        Dim strCurrencyCode, strCountryCode As String
        oCombo = aForm.Items.Item("17").Specific
        strCountryCode = oCombo.Selected.Value
        strCurrencyCode = oCombo.Selected.Description

        If strTypeValue <> "T" Then
            If strCountryCode = "" Then
                strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName], T0.[U_Z_NetSalary],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T2.[IBAN], T1.[bankAcount] FROM [dbo].[@Z_PAYROLL1]  T0 inner Join OHEM T1 on T1.empID=T0.U_Z_empID left Outer  JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode]"
                strQuery = strQuery & " where T0.U_Z_Posted='N' and T0.U_Z_Year=" & intYear & " and T0.U_Z_Month=" & intMonth & " and " & strCompany & " and " & strType
            Else
                strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName], T0.[U_Z_NetSalary],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T2.[IBAN], T1.[bankAcount] FROM [dbo].[@Z_PAYROLL1]  T0 inner Join OHEM T1 on T1.empID=T0.U_Z_empID left Outer  JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode]"
                strQuery = strQuery & " where T0.U_Z_Posted='N' and T0.U_Z_Year=" & intYear & " and T0.U_Z_Month=" & intMonth & " and " & strCompany & " and " & strType & " and T2.[CountryCod]='" & strCountryCode & "'"

            End If
        Else
            If strCompany = "" Then
                strCompany = " 1 = 1 "
            Else
                strCompany = strCompany.Replace("T0", "T3")
            End If
            If strCountryCode = "" Then
                strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName],SUM((CASE WHEN T0.[U_Z_Type] = 'D' THEN -T0.U_Z_Amount ELSE T0.U_Z_Amount END  ))  as U_Z_NetSalary,T2.[BankCode], T2.[BankName], T2.[SwiftNum], T2.[IBAN], T1.[bankAcount],T3.U_Z_CompNo FROM [dbo].[@Z_PAY_TRANS]  T0 inner Join OHEM T1 on T1.empID = T0.U_Z_empID left Outer  JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode] INNER JOIN [dbo].[OHEM] T3 ON T3.empID = T0.U_Z_EMPID"
                strQuery = strQuery & " where T0.U_Z_Posted='N' and T0.U_Z_Year=" & intYear & " and T0.U_Z_Month=" & intMonth & " and " & strCompany
                strQuery = strQuery & " group by T0.[U_Z_empid], T0.[U_Z_EmpName],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T2.[IBAN], T1.[bankAcount], T3.[U_Z_CompNo]"
            Else
                strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName], SUM((CASE WHEN T0.[U_Z_Type] = 'D' THEN -T0.U_Z_Amount ELSE T0.U_Z_Amount END  )) as U_Z_NetSalary, T2.[BankCode], T2.[BankName], T2.[SwiftNum], T2.[IBAN], T1.[bankAcount],T3.U_Z_CompNo FROM [dbo].[@Z_PAY_TRANS]  T0 inner Join OHEM T1 on T1.empID = T0.U_Z_empID left Outer  JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode] INNER JOIN [dbo].[OHEM] T3 ON T3.empID = T0.U_Z_EMPID"
                strQuery = strQuery & " where T0.U_Z_Posted='N' and T0.U_Z_Year=" & intYear & " and T0.U_Z_Month=" & intMonth & " and " & strCompany & " and T2.[CountryCod]='" & strCountryCode & "'"
                strQuery = strQuery & " group by T0.[U_Z_empid], T0.[U_Z_EmpName],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T2.[IBAN], T1.[bankAcount], T3.[U_Z_CompNo]"
            End If
        End If

        oRec.DoQuery(strQuery)
        If oRec.RecordCount > 0 Then
            oApplication.Utilities.Message("Payroll components are not posted for this selected year and month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        Dim dblExchangeRate As Double
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strLocal As String = oApplication.Utilities.getLocalCurrency()
        If strCurrencyCode = "None" Or strCurrencyCode = "All" Then
            strCurrencyCode = ""
            dblExchangeRate = 1
        Else
            If strCurrencyCode = "Local Currency" Or strCurrencyCode = strLocal Then
                'strCurrencyCode = ""
                dblExchangeRate = 1
            Else
                oTest.DoQuery("Select * from ORTT where Currency='" & strCurrencyCode & "' and RateDate='" & Now.Date.ToString("yyyy-MM-dd") & "'")
                If oTest.Fields.Item("Rate").Value <= 0 Then
                    oApplication.Utilities.Message("Exchange rate not defined...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    dblExchangeRate = oTest.Fields.Item("Rate").Value
                End If
            End If
        End If
        Dim oCheckbox As SAPbouiCOM.CheckBox
        oCheckbox = aForm.Items.Item("18").Specific
        If oCheckbox.Checked = True Then
            If GenerateFile_TAB(strBank, intYear, intMonth, strCompany, strType, strTypeValue, strCurrencyCode, dblExchangeRate) = True Then
                Return True
            Else
                Return False
            End If
        Else
            If GenerateFile(strBank, intYear, intMonth, strCompany, strType, strTypeValue, strCurrencyCode, dblExchangeRate) = True Then
                Return True
            Else
                Return False
            End If
        End If



        Return True
    End Function

    Private Function GenerateFile(ByVal aBank As String, ByVal ayear As Integer, ByVal amonth As Integer, ByVal aCompany As String, ByVal aType As String, ByVal strTypeValue As String, aCurrency As String, aExchangRate As Double) As Boolean
        Dim sLogPath, strQuery As String
        Dim TempString As String
        Dim sLogFilePath As String
        sLogPath = oApplication.Utilities.getApplicationPath() & "\Log"
        If Not Directory.Exists(sLogPath) Then
            Directory.CreateDirectory(sLogPath)
        End If
        oCombo = oForm.Items.Item("15").Specific
        If oCombo.Selected.Value = "T" Then
            sLogFilePath = sLogPath & "\AGOC_BankFile " & ayear.ToString("0000") & amonth.ToString("00") & ".txt"
        Else
            sLogFilePath = sLogPath & "\AGOC_BankFile " & ayear.ToString("0000") & amonth.ToString("00") & ".csv"
        End If


        If File.Exists(sLogFilePath) Then
            File.Delete(sLogFilePath)
        End If
        'Header
        Dim Day As Integer = DateTime.DaysInMonth(ayear, amonth)
        TempString = "000AGOCL001" & ayear.ToString("0000") & amonth.ToString("00") & Day.ToString("00")
        WriteToLog(TempString, sLogFilePath)
        Dim oRec, oTemp As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


        Dim strCurrencyCode, strCountryCode As String
        oCombo = oForm.Items.Item("17").Specific
        strCountryCode = oCombo.Selected.Value
        strCurrencyCode = oCombo.Selected.Description

        oRec.DoQuery("Select * from [@Z_PAYROLL1] where U_Z_Year=" & ayear & " and U_Z_MONTH=" & amonth & " and U_Z_POSTED='Y' order by U_Z_empID")
        If strTypeValue <> "T" Then
            If strCountryCode = "" Then
                strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName], T0.[U_Z_NetSalary],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN] 'IBAN', T1.[bankAcount],T1.[ExtEmpNo],T1.[homeCity],T1.[homeZip] FROM [dbo].[@Z_PAYROLL1]  T0 inner Join OHEM T1 on T1.empID=T0.U_Z_empID Left Outer JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode]"
                strQuery = strQuery & " where T0.U_Z_Posted='Y' and T0.U_Z_Year=" & ayear & " and T0.U_Z_Month=" & amonth & " and " & aCompany & " and " & aType & " order by U_Z_empID"
            Else
                strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName], T0.[U_Z_NetSalary],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN] 'IBAN', T1.[bankAcount],T1.[ExtEmpNo],T1.[homeCity],T1.[homeZip] FROM [dbo].[@Z_PAYROLL1]  T0 inner Join OHEM T1 on T1.empID=T0.U_Z_empID Left Outer JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode]"
                strQuery = strQuery & " where T0.U_Z_Posted='Y' and T0.U_Z_Year=" & ayear & " and T0.U_Z_Month=" & amonth & " and " & aCompany & " and " & aType & " and T2.CountryCod='" & strCountryCode & "' order by U_Z_empID"

            End If
        Else
            If strCountryCode = "" Then
                strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName],SUM((CASE WHEN T0.[U_Z_Type] = 'D' THEN -T0.U_Z_Amount ELSE T0.U_Z_Amount END  ))  as U_Z_NetSalary,T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN] 'IBAN',T1.[ExtEmpNo], T1.[bankAcount],T3.U_Z_CompNo,T1.[homeCity],T1.[homeZip] FROM [dbo].[@Z_PAY_TRANS]  T0 inner Join OHEM T1 on T1.empID=T0.U_Z_empID left Outer  JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode] INNER JOIN [dbo].[OHEM] T3 ON T3.empID = T0.U_Z_EMPID"
                strQuery = strQuery & " where T0.U_Z_Posted='Y' and T0.U_Z_Year=" & ayear & " and T0.U_Z_Month=" & amonth & " and " & aCompany
                strQuery = strQuery & " group by T0.[U_Z_empid], T0.[U_Z_EmpName],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN], T1.[bankAcount],T3.[U_Z_CompNo],T1.[homeCity],T1.[homeZip],T1.[ExtEmpNo]  order by U_Z_empID"
            Else
                strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName],SUM((CASE WHEN T0.[U_Z_Type] = 'D' THEN -T0.U_Z_Amount ELSE T0.U_Z_Amount END  ))  as U_Z_NetSalary,T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN] 'IBAN',T1.[ExtEmpNo], T1.[bankAcount],T3.U_Z_CompNo,T1.[homeCity],T1.[homeZip] FROM [dbo].[@Z_PAY_TRANS]  T0 inner Join OHEM T1 on T1.empID=T0.U_Z_empID left Outer  JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode] INNER JOIN [dbo].[OHEM] T3 ON T3.empID = T0.U_Z_EMPID"
                strQuery = strQuery & " where T0.U_Z_Posted='Y' and T0.U_Z_Year=" & ayear & " and T0.U_Z_Month=" & amonth & " and " & aCompany & " and T2.[CountryCod]='" & strCountryCode & "'"
                strQuery = strQuery & " group by T0.[U_Z_empid], T0.[U_Z_EmpName],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN], T1.[bankAcount],T3.[U_Z_CompNo],T1.[homeCity],T1.[homeZip],T1.[ExtEmpNo]  order by T0.[U_Z_empid]"
            End If
        End If


        oRec.DoQuery(strQuery)
        Dim intLineCount As Integer = 0
        Dim dblTotal As Double = 0
        Dim dblNetSalary As String
        Dim strLocalCurrency As String '= oApplication.Utilities.getLocalCurrency()
        If aCurrency = "" Or aCurrency = "Local Currency" Then
            strLocalCurrency = oApplication.Utilities.getLocalCurrency()
        Else
            strLocalCurrency = aCurrency
        End If
        For intRow As Integer = 0 To oRec.RecordCount - 1
            dblNetSalary = oRec.Fields.Item("U_Z_NetSalary").Value
            dblNetSalary = dblNetSalary * aExchangRate
            Dim dbl As Decimal = dblNetSalary
            dblNetSalary = Math.Round(dbl, 2)
            dblTotal = dblTotal + dblNetSalary
            ' dblNetSalary = 17055.11
            Dim strSal() As String
            strSal = dblNetSalary.ToString().Split(".")
            Dim st As String
            Dim dblSal1 As Double
            If strSal.Length > 1 Then
                dblSal1 = CDbl(strSal(0))
                If dblSal1 < 0 Then
                    st = (Format(CDbl(strSal(0)), "00000000000") & "." & Format(CDbl(strSal(1)), "00"))
                Else
                    st = (Format(CDbl(strSal(0)), "000000000000") & "." & Format(CDbl(strSal(1)), "00"))
                End If

            Else
                st = (Format(CDbl(strSal(0)), "000000000000000")) ' & "." & Format(CDbl(strSal(1)), "00"))
            End If
            '   Dim st As String = String.Format("{000000.00}", dblNetSalary) '; // "123.00"

            TempString = "112"
            TempString = TempString & (intRow + 1).ToString("000000") '9 Char
            TempString = TempString '& vbTab
            'Body a
            Dim ststrin As String = oRec.Fields.Item("ExtEmpNo").Value
            TempString = TempString & CInt(ststrin).ToString("0000000000") '10 Char
            TempString = AddFreeSpace(TempString, 6)
            TempString = TempString '& vbTab

            'Empllyee Batch no -B
            ststrin = oRec.Fields.Item("ExtEmpNo").Value

            TempString = TempString & CDbl(ststrin).ToString("0000000000") '10 Char
            TempString = AddFreeSpace(TempString, 2)
            TempString = TempString '& vbTab
            ' MsgBox(TempString.Length)
            'SwiftCode C
            Dim IBN As String = oRec.Fields.Item("SwiftNum").Value.ToString & "XXX" & oRec.Fields.Item("IBAN").Value
            If IBN.Length < 35 Then
                IBN = AddFreeSpace(IBN, 35 - Len(IBN))
            End If

            ' TempString = TempString & oRec.Fields.Item("SwiftNum").Value.ToString & "XXX" & oRec.Fields.Item("IBAN").Value
            TempString = TempString & IBN
            TempString = AddFreeSpace(TempString, 45 - Len(IBN))
            TempString = TempString '& vbTab


            '   MsgBox(TempString.Length)
            'Net Salary D
            Dim strSalary As String = dblNetSalary
            ' strSalary = (dblNetSalary.ToString("0000000000000.00"))

            TempString = TempString & st  '.ToString("000000000000.00") '15 Char
            TempString = TempString '& vbTab
            Dim otest1 As SAPbobsCOM.Recordset
            otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otest1.DoQuery("Select * from DSC1 where Account='" & aBank & "'")

            Dim strKJOBankAccount = aBank ' "3140500009942"
            strKJOBankAccount = AddFreeSpace(strKJOBankAccount, 13 - Len(strKJOBankAccount))
            TempString = TempString & strLocalCurrency & strKJOBankAccount
            TempString = AddFreeSpace(TempString, 22)
            TempString = TempString '& vbTab
            '  MsgBox(TempString.Length)
            'E Employee City of Residency
            Dim strCityofRes As String = oRec.Fields.Item("homeCity").Value '"KHAFJI"
            strCityofRes = AddFreeSpace(strCityofRes, 20 - Len(strCityofRes))
            TempString = TempString & strCityofRes
            ' TempString = AddFreeSpace(TempString, 20)
            TempString = TempString '& vbTab
            '  MsgBox((TempString.Length))
            'F

            Dim strCityofResPotalCode As String = oRec.Fields.Item("homeZip").Value
            '  MsgBox(strCityofResPotalCode.Length)
            strCityofResPotalCode = AddFreeSpace(strCityofResPotalCode, 9 - Len(strCityofResPotalCode)) '5Char
            TempString = TempString & strCityofResPotalCode
            ' TempString = AddFreeSpace(TempString, 4)
            TempString = TempString '& vbTab
            '  MsgBox((TempString.Length))
            'g Employee Name
            Dim strEmpName As String = oRec.Fields.Item("U_Z_EmpName").Value
            TempString = TempString & strEmpName
            TempString = AddFreeSpace(TempString, 35 - Len(strEmpName))
            TempString = TempString '& vbTab
            'h  Payroll
            TempString = TempString '& "PAYROLL"
            TempString = AddFreeSpace(TempString, 4)
            TempString = TempString ' & vbTab
            WriteToLog(TempString, sLogFilePath)

            intLineCount = intLineCount + 1
            '  dblTotal = dblTotal + oRec.Fields.Item("U_Z_NetSalary").Value
            oRec.MoveNext()
        Next
        'Footer
        Dim strSal1() As String
        dblTotal = Math.Round(dblTotal, 2)
        strSal1 = dblTotal.ToString().Split(".")
        Dim st1 As String
        If strSal1.Length > 1 Then
            st1 = (Format(CDbl(strSal1(0)), "000000000000000") & "." & Format(CDbl(strSal1(1)), "00"))
        Else
            st1 = (Format(CDbl(strSal1(0)), "000000000000000000")) '& "." & Format(CDbl(strSal1(1)), "00"))
        End If

        TempString = "999" & st1 & intLineCount.ToString("000000")
        'MsgBox(TempString.Length)
        WriteToLog(TempString, sLogFilePath)
        ShellExecute(sLogFilePath)
        Return True
    End Function

    Private Function GenerateFile_TAB(ByVal aBank As String, ByVal ayear As Integer, ByVal amonth As Integer, ByVal aCompany As String, ByVal aType As String, ByVal strTypeValue As String, aCurrency As String, aExchangRate As Double) As Boolean
        Dim sLogPath, strQuery As String
        Dim TempString As String
        Dim sLogFilePath As String
        sLogPath = oApplication.Utilities.getApplicationPath() & "\Log"
        If Not Directory.Exists(sLogPath) Then
            Directory.CreateDirectory(sLogPath)
        End If
        oCombo = oForm.Items.Item("15").Specific
        If oCombo.Selected.Value = "T" Then
            sLogFilePath = sLogPath & "\AGOC_BankFile " & ayear.ToString("0000") & amonth.ToString("00") & ".txt"
        Else
            sLogFilePath = sLogPath & "\AGOC_BankFile " & ayear.ToString("0000") & amonth.ToString("00") & ".csv"
        End If


        If File.Exists(sLogFilePath) Then
            File.Delete(sLogFilePath)
        End If
        'Header
        Dim Day As Integer = DateTime.DaysInMonth(ayear, amonth)
        TempString = "000AGOCL001" & ayear.ToString("0000") & amonth.ToString("00") & Day.ToString("00")
        WriteToLog(TempString, sLogFilePath)
        Dim oRec, oTemp As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


        Dim strCurrencyCode, strCountryCode As String
        oCombo = oForm.Items.Item("17").Specific
        strCountryCode = oCombo.Selected.Value
        strCurrencyCode = oCombo.Selected.Description

        oRec.DoQuery("Select * from [@Z_PAYROLL1] where U_Z_Year=" & ayear & " and U_Z_MONTH=" & amonth & " and U_Z_POSTED='Y' order by U_Z_empID")
        If strTypeValue <> "T" Then
            If strCountryCode = "" Then
                strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName], T0.[U_Z_NetSalary],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN] 'IBAN', T1.[bankAcount],T1.[ExtEmpNo],T1.[homeCity],T1.[homeZip] FROM [dbo].[@Z_PAYROLL1]  T0 inner Join OHEM T1 on T1.empID=T0.U_Z_empID Left Outer JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode]"
                strQuery = strQuery & " where T0.U_Z_Posted='Y' and T0.U_Z_Year=" & ayear & " and T0.U_Z_Month=" & amonth & " and " & aCompany & " and " & aType & " order by U_Z_empID"
            Else
                strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName], T0.[U_Z_NetSalary],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN] 'IBAN', T1.[bankAcount],T1.[ExtEmpNo],T1.[homeCity],T1.[homeZip] FROM [dbo].[@Z_PAYROLL1]  T0 inner Join OHEM T1 on T1.empID=T0.U_Z_empID Left Outer JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode]"
                strQuery = strQuery & " where T0.U_Z_Posted='Y' and T0.U_Z_Year=" & ayear & " and T0.U_Z_Month=" & amonth & " and " & aCompany & " and " & aType & " and T2.CountryCod='" & strCountryCode & "' order by U_Z_empID"

            End If
        Else
            If strCountryCode = "" Then
                strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName],SUM((CASE WHEN T0.[U_Z_Type] = 'D' THEN -T0.U_Z_Amount ELSE T0.U_Z_Amount END  ))  as U_Z_NetSalary,T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN] 'IBAN',T1.[ExtEmpNo], T1.[bankAcount],T3.U_Z_CompNo,T1.[homeCity],T1.[homeZip] FROM [dbo].[@Z_PAY_TRANS]  T0 inner Join OHEM T1 on T1.empID=T0.U_Z_empID left Outer  JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode] INNER JOIN [dbo].[OHEM] T3 ON T3.empID = T0.U_Z_EMPID"
                strQuery = strQuery & " where T0.U_Z_Posted='Y' and T0.U_Z_Year=" & ayear & " and T0.U_Z_Month=" & amonth & " and " & aCompany
                strQuery = strQuery & " group by T0.[U_Z_empid], T0.[U_Z_EmpName],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN], T1.[bankAcount],T3.[U_Z_CompNo],T1.[homeCity],T1.[homeZip],T1.[ExtEmpNo]  order by U_Z_empID"
            Else
                strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName],SUM((CASE WHEN T0.[U_Z_Type] = 'D' THEN -T0.U_Z_Amount ELSE T0.U_Z_Amount END  ))  as U_Z_NetSalary,T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN] 'IBAN',T1.[ExtEmpNo], T1.[bankAcount],T3.U_Z_CompNo,T1.[homeCity],T1.[homeZip] FROM [dbo].[@Z_PAY_TRANS]  T0 inner Join OHEM T1 on T1.empID=T0.U_Z_empID left Outer  JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode] INNER JOIN [dbo].[OHEM] T3 ON T3.empID = T0.U_Z_EMPID"
                strQuery = strQuery & " where T0.U_Z_Posted='Y' and T0.U_Z_Year=" & ayear & " and T0.U_Z_Month=" & amonth & " and " & aCompany & " and T2.[CountryCod]='" & strCountryCode & "'"
                strQuery = strQuery & " group by T0.[U_Z_empid], T0.[U_Z_EmpName],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN], T1.[bankAcount],T3.[U_Z_CompNo],T1.[homeCity],T1.[homeZip],T1.[ExtEmpNo]  order by T0.[U_Z_empid]"
            End If
        End If

        ' strQuery = "SELECT T0.[U_Z_empid], T0.[U_Z_EmpName], T0.[U_Z_NetSalary],T2.[BankCode], T2.[BankName], T2.[SwiftNum], T1.[U_Z_IBAN] 'IBAN', T1.[bankAcount],T1.[ExtEmpNo],T1.[homeCity],T1.[homeZip] FROM [dbo].[@Z_PAYROLL1]  T0 inner Join OHEM T1 on T1.empID=T0.U_Z_empID Left Outer JOIN ODSC T2 ON T1.[bankCode] = T2.[BankCode]"
        ' strQuery = strQuery & " where T0.U_Z_Posted='Y' and T0.U_Z_Year=" & ayear & " and T0.U_Z_Month=" & amonth & " and " & aCompany & " and " & aType & " order by U_Z_empID"
        oRec.DoQuery(strQuery)
        Dim intLineCount As Integer = 0
        Dim dblTotal As Double = 0
        Dim dblNetSalary As String
        Dim strLocalCurrency As String '= oApplication.Utilities.getLocalCurrency()
        If aCurrency = "" Or aCurrency = "Local Currency" Then
            strLocalCurrency = oApplication.Utilities.getLocalCurrency()
        Else
            strLocalCurrency = aCurrency
        End If
        Dim otest1 As SAPbobsCOM.Recordset
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intRow As Integer = 0 To oRec.RecordCount - 1
            dblNetSalary = oRec.Fields.Item("U_Z_NetSalary").Value
            dblNetSalary = dblNetSalary * aExchangRate
            Dim dbl As Decimal = dblNetSalary
            dblNetSalary = Math.Round(dbl, 2)
            dblTotal = dblTotal + dblNetSalary
            ' dblNetSalary = 17055.11
            Dim strSal() As String
            strSal = dblNetSalary.ToString().Split(".")
            Dim st As String
            Dim dblSal1 As Double
            If strSal.Length > 1 Then
                dblSal1 = CDbl(strSal(0))
                If dblSal1 < 0 Then
                    st = (Format(CDbl(strSal(0)), "00000000000") & "." & Format(CDbl(strSal(1)), "00"))
                Else
                    st = (Format(CDbl(strSal(0)), "000000000000") & "." & Format(CDbl(strSal(1)), "00"))
                End If

            Else
                st = (Format(CDbl(strSal(0)), "000000000000000")) ' & "." & Format(CDbl(strSal(1)), "00"))
            End If
            '   Dim st As String = String.Format("{000000.00}", dblNetSalary) '; // "123.00"


            'Body a
            TempString = "112"
            TempString = TempString & (intRow + 1).ToString("000000") '9 Char
            TempString = TempString
            Dim ststrin As String = oRec.Fields.Item("ExtEmpNo").Value
            TempString = TempString & CInt(ststrin).ToString("0000000000") '10 Char
            TempString = AddFreeSpace(TempString, 6)
            TempString = TempString & vbTab

            'Empllyee Batch no -B
            ststrin = oRec.Fields.Item("ExtEmpNo").Value
            TempString = TempString & CDbl(ststrin).ToString("0000000000") '10 Char
            TempString = AddFreeSpace(TempString, 2)
            TempString = TempString & vbTab

            'SwiftCode C
            Dim IBN As String = oRec.Fields.Item("SwiftNum").Value.ToString & "XXX" & oRec.Fields.Item("IBAN").Value
            If IBN.Length < 35 Then
                IBN = AddFreeSpace(IBN, 35 - Len(IBN))
            End If
            ' TempString = TempString & oRec.Fields.Item("SwiftNum").Value.ToString & "XXX" & oRec.Fields.Item("IBAN").Value
            TempString = TempString & IBN
            TempString = AddFreeSpace(TempString, 45 - Len(IBN))
            TempString = TempString & vbTab
            '  MsgBox(TempString.Length)
            'Net Salary D
            Dim strSalary As String = dblNetSalary
            ' strSalary = (dblNetSalary.ToString("0000000000000.00"))
            TempString = TempString & st  '.ToString("000000000000.00") '15 Char
            TempString = TempString '& vbTab
            otest1.DoQuery("Select * from DSC1 where Account='" & aBank & "'")
            Dim strKJOBankAccount = aBank ' "3140500009942"
            strKJOBankAccount = AddFreeSpace(strKJOBankAccount, 13 - Len(strKJOBankAccount))
            TempString = TempString & strLocalCurrency & strKJOBankAccount
            TempString = AddFreeSpace(TempString, 22)
            TempString = TempString & vbTab

            'E Employee City of Residency
            Dim strCityofRes As String = oRec.Fields.Item("homeCity").Value '"KHAFJI"
            strCityofRes = AddFreeSpace(strCityofRes, 20 - Len(strCityofRes))
            TempString = TempString & strCityofRes
            TempString = TempString & vbTab

            'F

            Dim strCityofResPotalCode As String = oRec.Fields.Item("homeZip").Value

            strCityofResPotalCode = AddFreeSpace(strCityofResPotalCode, 9 - Len(strCityofResPotalCode)) '5Char
            TempString = TempString & strCityofResPotalCode
            TempString = TempString & vbTab

            'g Employee Name
            Dim strEmpName As String = oRec.Fields.Item("U_Z_EmpName").Value
            TempString = TempString & strEmpName
            TempString = AddFreeSpace(TempString, 35 - Len(strEmpName))
            TempString = TempString & vbTab
            'h  Payroll
            TempString = TempString '& "PAYROLL"
            TempString = AddFreeSpace(TempString, 4)
            TempString = TempString '& vbTab
            WriteToLog(TempString, sLogFilePath)

            intLineCount = intLineCount + 1
            oRec.MoveNext()
        Next
        'Footer
        Dim strSal1() As String
        dblTotal = Math.Round(dblTotal, 2)
        strSal1 = dblTotal.ToString().Split(".")
        Dim st1 As String
        If strSal1.Length > 1 Then
            st1 = (Format(CDbl(strSal1(0)), "000000000000000") & "." & Format(CDbl(strSal1(1)), "00"))
        Else
            st1 = (Format(CDbl(strSal1(0)), "000000000000000000")) '& "." & Format(CDbl(strSal1(1)), "00"))
        End If

        TempString = "999" & st1 & intLineCount.ToString("000000")
        ' MsgBox(TempString.Length)
        WriteToLog(TempString, sLogFilePath)
        ShellExecute(sLogFilePath)
        Return True
    End Function
    Private Function ShellExecute(ByVal File As String) As Boolean
        Dim myProcess As New Process
        myProcess.StartInfo.FileName = File
        myProcess.StartInfo.UseShellExecute = True
        myProcess.StartInfo.RedirectStandardOutput = False
        myProcess.Start()
        myProcess.Dispose()
    End Function
    Private Function AddFreeSpace(ByVal astring As String, ByVal aCount As Integer) As String
        For intRow As Integer = 0 To aCount - 1
            astring = astring & " "
        Next
        Return astring
    End Function

    Public Sub WriteToLog(ByVal sText As String, ByVal sFilePath As String)
        Dim sLogFilePath As String
        Dim oStream As StreamWriter
        Try
            sLogFilePath = sFilePath

            If Not File.Exists(sLogFilePath) Then
                Dim fileStream As FileStream
                fileStream = New FileStream(sLogFilePath, FileMode.Create, FileAccess.Write)
                Dim sw As New StreamWriter(fileStream)
                sw.BaseStream.Seek(0, SeekOrigin.End)
                sw.WriteLine(sText)
                sw.Flush()
                sw.Close()
            Else
                Dim fileStream As FileStream
                fileStream = New FileStream(sLogFilePath, FileMode.Append, FileAccess.Write)
                Dim sw As New StreamWriter(fileStream)
                sw.BaseStream.Seek(0, SeekOrigin.End)
                sw.WriteLine(sText)
                sw.Flush()
                sw.Close()
            End If

        Catch ex As Exception
            Throw (ex)
        Finally
            '  sw.Close()
            '  sw = Nothing
        End Try
    End Sub

    Public Sub Trace_Process(ByVal strContent As String, ByVal strFile As String)
        Try
            Dim strPath As String = System.Windows.Forms.Application.StartupPath.ToString() + strFile
            If Not File.Exists(strPath) Then
                Dim fileStream As FileStream
                fileStream = New FileStream(strPath, FileMode.Create, FileAccess.Write)
                Dim sw As New StreamWriter(fileStream)
                sw.BaseStream.Seek(0, SeekOrigin.End)
                sw.WriteLine(strContent)
                sw.Flush()
                sw.Close()
            Else
                Dim fileStream As FileStream
                fileStream = New FileStream(strPath, FileMode.Append, FileAccess.Write)
                Dim sw As New StreamWriter(fileStream)
                sw.BaseStream.Seek(0, SeekOrigin.End)
                sw.WriteLine(strContent)
                sw.Flush()
                sw.Close()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case mnu_Pay_ERBF
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
            End Select
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Pay_ERBF Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "3" Then 'Browse
                                    If Processing(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else

                                    End If
                                End If

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region


End Class
