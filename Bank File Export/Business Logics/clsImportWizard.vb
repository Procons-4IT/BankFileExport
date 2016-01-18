Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsImportWizard
    Inherits clsBase
    Private strQuery As String
    Private oGrid As SAPbouiCOM.Grid

    Private oDt_Import As SAPbouiCOM.DataTable
    Private oDt_ErrorLog As SAPbouiCOM.DataTable

    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oEditColumn As SAPbouiCOM.EditTextColumn
    Private oRecordSet As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_ImpWiz, frm_ImpWiz)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            Initialize(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case mnu_ImpWiz
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
            If pVal.FormTypeEx = frm_ImpWiz Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "9" Then 'Browse
                                    oApplication.Utilities.OpenFileDialogBox(oForm, "17")
                                ElseIf (pVal.ItemUID = "7") Then 'Next
                                    If CType(oForm.Items.Item("17").Specific, SAPbouiCOM.StaticText).Caption <> "" Then
                                        If oApplication.Utilities.ValidateFile(oForm, "17") Then
                                            If oApplication.Utilities.GetExcelData(oForm, "17") Then
                                                loadData(oForm)
                                                oForm.Items.Item("4").Enabled = True
                                                oApplication.Utilities.Message("PERI Data Imported Successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            Else
                                                BubbleEvent = False
                                            End If
                                        End If
                                    Else
                                        oApplication.Utilities.Message("Select File to Import....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                    End If
                                ElseIf (pVal.ItemUID = "3") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 1
                                    oForm.Items.Item("3").Enabled = False
                                    oForm.Freeze(False)
                                ElseIf (pVal.ItemUID = "4") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 2
                                    oForm.Items.Item("3").Enabled = True
                                    oForm.Items.Item("5").Enabled = True
                                    oForm.Freeze(False)
                                ElseIf (pVal.ItemUID = "5") Then
                                    oGrid = oForm.Items.Item("14").Specific
                                    If oGrid.Rows.Count > 0 Then
                                        If oGrid.DataTable.GetValue("Error", 0).ToString() <> "" Then
                                            oApplication.Utilities.Message("Error found in file (Refer Invalid Data tab)....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                        Else
                                            Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Sure you Want to Import Data to PERI Invoices?", 2, "Yes", "No", "")
                                            If _retVal = 1 Then
                                                oGrid = oForm.Items.Item("8").Specific
                                                If oApplication.Utilities.Import_PERI_InvoiceData(oForm) Then
                                                    oApplication.Utilities.Message("PERI Import Created Successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                    oForm.Close()
                                                End If
                                            Else
                                                BubbleEvent = False
                                            End If
                                        End If
                                    ElseIf oGrid.Rows.Count = 0 Then
                                        Dim _retVal As Integer = oApplication.SBO_Application.MessageBox("Sure you Want to Import Data to PERI Invoices?", 2, "Yes", "No", "")
                                        If _retVal = 1 Then
                                            oGrid = oForm.Items.Item("8").Specific
                                            If oApplication.Utilities.Import_PERI_InvoiceData(oForm) Then
                                                oApplication.Utilities.Message("PERI Import Created Successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                                oForm.Close()
                                            End If
                                        Else
                                            BubbleEvent = False
                                        End If
                                    End If
                                ElseIf (pVal.ItemUID = "12") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 2
                                    oForm.Freeze(False)
                                ElseIf (pVal.ItemUID = "13") Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 3
                                    oForm.Freeze(False)
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
            If oForm.TypeEx = frm_ImpWiz Then

            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"
    Private Sub Initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim sQuery As String
            oForm.DataSources.DataTables.Add("Dt_Import")
            oForm.DataSources.DataTables.Add("Dt_ErrorLog")

            oDt_Import = oForm.DataSources.DataTables.Item("Dt_Import")
            sQuery = " Select InvoiceNo,CardCode,CardName,GetDate() As DocDate,GetDate() As DueDate,Project,LineTotal,DocTotal,Line,RevAcct From Z_PEIM "
            sQuery += "  Where 1 = 2 "
            oDt_Import.ExecuteQuery(sQuery)
            oGrid = oForm.Items.Item("8").Specific
            oGrid.DataTable = oDt_Import
            formatGrid(oForm)

            oDt_ErrorLog = oForm.DataSources.DataTables.Item("Dt_ErrorLog")
            oDt_ErrorLog.ExecuteQuery("Select Convert(VarChar(250),'') As 'Error'")
            oGrid = oForm.Items.Item("14").Specific
            oGrid.DataTable = oDt_ErrorLog

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub formatGrid(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            oGrid = oForm.Items.Item("8").Specific
            formatAll(oForm, oGrid)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub loadData(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDt_Import = oForm.DataSources.DataTables.Item("Dt_Import")
            strQuery = " Select T0.InvoiceNo,T0.CardCode,T0.CardName, "
            strQuery += " Convert(DateTime,(SubString(T0.DocDate,1,4)+'-'+SubString(T0.DocDate,5,2)+'-'+SubString(T0.DocDate,7,2))) As DocDate,"
            strQuery += " Convert(DateTime,(SubString(T0.DueDate,1,4)+'-'+SubString(T0.DueDate,5,2)+'-'+SubString(T0.DueDate,7,2))) As DueDate,"
            strQuery += " T0.Project,T0.LineTotal,T0.DocTotal,T0.Line,T0.RevAcct From Z_PEIM T0 "
            strQuery += " JOIN OCRD T1 On T1.CardCode = T0.CardCode "
            strQuery += " JOIN OACT T2 On T0.RevAcct =  T2.AcctCode "
            strQuery += " And T0.Line > 0 "
            oDt_Import.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("8").Specific
            oGrid.DataTable = oDt_Import
            formatAll(oForm, oGrid)

            For index As Integer = 0 To oGrid.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, index + 1)
            Next

            oDt_ErrorLog = oForm.DataSources.DataTables.Item("Dt_ErrorLog")
            strQuery = " Select 'InValid Card Code : ' + T0.CardCode As 'Error' From Z_PEIM T0 "
            strQuery += " LEFT OUTER JOIN OCRD T1 On T1.CardCode = T0.CardCode "
            strQuery += " Where T1.CardCode Is Null And Line > 0 "
            strQuery += " Union All "
            strQuery += " Select 'InValid Revenue Account : ' + T0.RevAcct As 'Error' From Z_PEIM T0 "
            strQuery += " LEFT OUTER JOIN OACT T1 On T1.AcctCode = T0.RevAcct "
            strQuery += " Where T1.AcctCode Is Null And Line > 0 "

            oDt_ErrorLog.ExecuteQuery(strQuery)
            oGrid = oForm.Items.Item("14").Specific
            oGrid.DataTable = oDt_ErrorLog
            For index As Integer = 0 To oGrid.Rows.Count - 1
                oGrid.RowHeaders.SetText(index, index + 1)
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub formatAll(ByVal oForm As SAPbouiCOM.Form, ByVal oGrid As SAPbouiCOM.Grid)
        Try
            oForm.Freeze(True)

            oGrid.Columns.Item("InvoiceNo").TitleObject.Caption = "Invoice No"
            oGrid.Columns.Item("CardCode").TitleObject.Caption = "Customer Code"
            oGrid.Columns.Item("CardName").TitleObject.Caption = "Customer Name"
            oGrid.Columns.Item("CardName").Visible = False
            oGrid.Columns.Item("DocDate").TitleObject.Caption = "Document Date"
            oGrid.Columns.Item("DueDate").TitleObject.Caption = "Due Date"
            oGrid.Columns.Item("Project").TitleObject.Caption = "Project"
            oGrid.Columns.Item("LineTotal").TitleObject.Caption = "Line Total"
            oGrid.Columns.Item("DocTotal").TitleObject.Caption = "Document Total"
            oGrid.Columns.Item("RevAcct").TitleObject.Caption = "Revenue Account"

            oGrid.Columns.Item("CardCode").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditColumn = oGrid.Columns.Item("CardCode")
            oEditColumn.LinkedObjectType = "2"

            oGrid.Columns.Item("LineTotal").RightJustified = True
            oGrid.Columns.Item("DocTotal").RightJustified = True

            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#End Region

End Class
