
Public Class clsJVApprovalTemplate
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private InvForConsumedItems, count As Integer
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines_1 As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines_2 As SAPbouiCOM.DBDataSource
    Public MatrixId As String
    Public intSelectedMatrixrow As Integer = 0
    Private RowtoDelete As Integer
    Private oRecordSet As SAPbobsCOM.Recordset
    Private dtValidFrom, dtValidTo As Date
    Private strQuery As String

#Region "Initialization"

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#End Region

#Region "Load Form"

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_OJAT, frm_OJAT)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm)
            enableControls(oForm, True)
            oMatrix = oForm.Items.Item("9").Specific
            oMatrix.AutoResizeColumns()

            oMatrix = oForm.Items.Item("10").Specific
            oMatrix.AutoResizeColumns()
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

#End Region

#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_OJAT Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            If validation(oForm) = False Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@OJAT")
                                If pVal.ItemUID = "9" Or pVal.ItemUID = "10" Then
                                    If (oDBDataSource.GetValue("Code", 0).ToString() = "") Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Select Code to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    ElseIf (oDBDataSource.GetValue("Name", 0).ToString() = "") Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Select Name to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    End If
                                    If pVal.ItemUID = "9" Then
                                        MatrixId = pVal.ItemUID
                                        intSelectedMatrixrow = pVal.Row
                                    ElseIf pVal.ItemUID = "10" Then
                                        MatrixId = pVal.ItemUID
                                        intSelectedMatrixrow = pVal.Row
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "13"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
                                    Case "14"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@OJAT")
                                oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@JAT1")
                                oDBDataSourceLines_2 = oForm.DataSources.DBDataSources.Item("@JAT2")
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If (pVal.ItemUID = "9" And pVal.ColUID = "V_0") Then
                                            oMatrix = oForm.Items.Item("9").Specific
                                            oMatrix.LoadFromDataSource()
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 1 Then
                                                intAddRows -= 1
                                                oMatrix.AddRow(intAddRows, pVal.Row - 1)
                                            End If
                                            oMatrix.FlushToDataSource()
                                            For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                oDBDataSourceLines_1.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                oDBDataSourceLines_1.SetValue("U_OUser", pVal.Row + index - 1, oDataTable.GetValue("USER_CODE", index))
                                                oDBDataSourceLines_1.SetValue("U_OName", pVal.Row + index - 1, oDataTable.GetValue("U_NAME", index))
                                            Next
                                            oMatrix.LoadFromDataSource()
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf (pVal.ItemUID = "10" And pVal.ColUID = "V_0") Then
                                            oMatrix = oForm.Items.Item("10").Specific
                                            oMatrix.LoadFromDataSource()
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 1 Then
                                                intAddRows -= 1
                                                oMatrix.AddRow(intAddRows, pVal.Row - 1)
                                            End If
                                            oMatrix.FlushToDataSource()
                                            For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                oDBDataSourceLines_2.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                oDBDataSourceLines_2.SetValue("U_AUser", pVal.Row + index - 1, oDataTable.GetValue("USER_CODE", index))
                                                oDBDataSourceLines_2.SetValue("U_AName", pVal.Row + index - 1, oDataTable.GetValue("U_NAME", index))
                                            Next
                                            oMatrix.LoadFromDataSource()
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    End If
                                Catch ex As Exception
                                    Throw
                                End Try
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Menu Event"

    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_OJAT
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddRow(oForm)
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    End If
                Case mnu_ADD
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        initialize(oForm)
                        enableControls(oForm, True)
                    End If
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        enableControls(oForm, True)
                    End If
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#End Region

#Region "Data Events"

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            If oForm.TypeEx = frm_OJAT Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@OJAT")
                                enableControls(oForm, False)
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@OJAT")
            oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@JAT1")
            oDBDataSourceLines_2 = oForm.DataSources.DBDataSources.Item("@JAT2")

            oMatrix = oForm.Items.Item("9").Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()

            oMatrix = oForm.Items.Item("10").Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()

            oForm.Update()
            MatrixId = "9"
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "Methods"
    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)

            Select Case aForm.PaneLevel
                Case "1"
                    oMatrix = aForm.Items.Item("9").Specific
                    oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@JAT1")
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLines_1.Size
                        oDBDataSourceLines_1.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                Case "2"
                    oMatrix = aForm.Items.Item("10").Specific
                    oDBDataSourceLines_2 = oForm.DataSources.DBDataSources.Item("@JAT2")
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLines_2.Size
                        oDBDataSourceLines_2.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
#End Region

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Select Case aForm.PaneLevel
                Case "1"
                    oMatrix = aForm.Items.Item("9").Specific
                    oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@JAT1")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    Else
                        oMatrix.AddRow(1, oMatrix.RowCount + 1)
                        oMatrix.ClearRowData(oMatrix.RowCount)
                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If
                    End If
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLines_1.Size
                        oDBDataSourceLines_1.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo(aForm)
                Case "2"
                    oMatrix = aForm.Items.Item("10").Specific
                    oDBDataSourceLines_2 = oForm.DataSources.DBDataSources.Item("@JAT2")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    Else
                        oMatrix.AddRow(1, oMatrix.RowCount + 1)
                        oMatrix.ClearRowData(oMatrix.RowCount)
                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If
                    End If
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLines_2.Size
                        oDBDataSourceLines_2.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo(aForm)
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@OJAT")
            oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@JAT1")
            oDBDataSourceLines_2 = oForm.DataSources.DBDataSources.Item("@JAT2")

            If Me.MatrixId = "9" Then
                oMatrix = aForm.Items.Item("9").Specific
                Me.RowtoDelete = intSelectedMatrixrow
                oDBDataSourceLines_1.RemoveRecord(Me.RowtoDelete - 1)
                oMatrix.LoadFromDataSource()
                oMatrix.FlushToDataSource()
                For count = 1 To oDBDataSourceLines_1.Size
                    oDBDataSourceLines_1.SetValue("LineId", count - 1, count)
                Next
            ElseIf (Me.MatrixId = "10") Then
                oMatrix = aForm.Items.Item("10").Specific
                Me.RowtoDelete = intSelectedMatrixrow
                oDBDataSourceLines_2.RemoveRecord(Me.RowtoDelete - 1)
                oMatrix.LoadFromDataSource()
                oMatrix.FlushToDataSource()
                For count = 1 To oDBDataSourceLines_2.Size
                    oDBDataSourceLines_2.SetValue("LineId", count - 1, count)
                Next
            End If
            oMatrix.LoadFromDataSource()
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
#End Region

#Region "Validations"
    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            aForm.Freeze(True)
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@OJAT")
            oDBDataSourceLines_1 = oForm.DataSources.DBDataSources.Item("@JAT1")
            oDBDataSourceLines_2 = oForm.DataSources.DBDataSources.Item("@JAT2")

            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oApplication.Utilities.getEdittextvalue(aForm, "4") = "" Then
                oApplication.Utilities.Message("Enter Code...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            ElseIf oApplication.Utilities.getEdittextvalue(aForm, "6") = "" Then
                oApplication.Utilities.Message("Enter Name...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If

            oMatrix = oForm.Items.Item("9").Specific
            oMatrix.LoadFromDataSource()
            If oMatrix.RowCount = 0 Then
                oApplication.Utilities.Message("Originator Row Cannot be Empty...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                Dim blnItemExist As Boolean = True
                ' For index As Integer = 1 To oMatrix.RowCount
                For index As Integer = oMatrix.RowCount To 1 Step -1
                    If CType(oMatrix.Columns.Item("V_0").Cells.Item(index).Specific, SAPbouiCOM.EditText).Value = "" Then
                        ' oApplication.Utilities.Message("Originator Cannot be Empty for Row: " + index.ToString() + "...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.DeleteRow(index)
                        '  Return False
                    End If
                Next
            End If

            oMatrix = oForm.Items.Item("10").Specific
            oMatrix.FlushToDataSource()
            oMatrix.LoadFromDataSource()
            If oMatrix.RowCount = 0 Then
                oApplication.Utilities.Message("Authorizer Row Cannot be Empty...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                Dim blnItemExist As Boolean = True
                'For index As Integer = 1 To oMatrix.RowCount
                '    If CType(oMatrix.Columns.Item("V_0").Cells.Item(index).Specific, SAPbouiCOM.EditText).Value = "" Then
                '        oApplication.Utilities.Message("Authorizer Cannot be Empty for Row: " + index.ToString() + "...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '        Return False
                '    End If
                'Next
                For index As Integer = oMatrix.RowCount To 1 Step -1
                    If CType(oMatrix.Columns.Item("V_0").Cells.Item(index).Specific, SAPbouiCOM.EditText).Value = "" Then
                        ' oApplication.Utilities.Message("Originator Cannot be Empty for Row: " + index.ToString() + "...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oMatrix.DeleteRow(index)
                        '  Return False
                    End If
                Next
            End If

            oMatrix = oForm.Items.Item("9").Specific
            oMatrix.LoadFromDataSource()
            For i As Integer = 1 To oMatrix.RowCount
                If CType(oMatrix.Columns.Item("V_0").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value <> "" Then
                    For j As Integer = 1 To oMatrix.RowCount
                        If i <> j Then
                            If (CType(oMatrix.Columns.Item("V_0").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value = CType(oMatrix.Columns.Item("V_0").Cells.Item(j).Specific, SAPbouiCOM.EditText).Value) Then
                                oApplication.Utilities.Message("Originator Duplicated in Row : " + j.ToString() + "...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                aForm.Freeze(False)
                                Return False
                            End If
                        End If
                    Next
                End If
            Next
            oMatrix = oForm.Items.Item("10").Specific
            oMatrix.LoadFromDataSource()
            Dim intApprover As Integer
            For i As Integer = 1 To oMatrix.RowCount
                If CType(oMatrix.Columns.Item("V_0").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value <> "" Then
                    intApprover += 1
                End If
            Next
            If intApprover > 4 Then
                oApplication.Utilities.Message("Can Have Maximum of 4 Authorizer...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select 1 As 'Return',DocEntry From [@OJAT]"
            strQuery += " Where "
            strQuery += " Code = '" + oDBDataSource.GetValue("Code", 0).Trim() + "' And DocEntry <> '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oApplication.Utilities.Message("Code Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If
            oMatrix = oForm.Items.Item("9").Specific
            oMatrix.LoadFromDataSource()
            For i As Integer = 1 To oMatrix.RowCount
                If CType(oMatrix.Columns.Item("V_0").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value <> "" Then
                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strQuery = "Select 1 As 'Return' From [@JAT1]"
                    strQuery += " Where "
                    strQuery += " Code <> '" + oDBDataSource.GetValue("Code", 0).Trim() + "'"
                    strQuery += " And U_OUser = '" + CType(oMatrix.Columns.Item("V_0").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value + "'"
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        oApplication.Utilities.Message(CType(oMatrix.Columns.Item("V_0").Cells.Item(i).Specific, SAPbouiCOM.EditText).Value + " Already Defined in Other Template...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aForm.Freeze(False)
                        Return False
                    End If
                End If
            Next
            aForm.Freeze(False)
            Return True
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Function
#End Region

#Region "Disable Controls"

    Private Sub enableControls(ByVal oForm As SAPbouiCOM.Form, ByVal blnEnable As Boolean)
        Try
            'oForm.Items.Item("12").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("4").Enabled = blnEnable
            oForm.Items.Item("6").Enabled = blnEnable
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#End Region

End Class
