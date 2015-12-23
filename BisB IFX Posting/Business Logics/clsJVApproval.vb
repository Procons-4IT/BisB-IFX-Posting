Public Class clsJVApproved
    Inherits clsBase

    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Dim oStatic As SAPbouiCOM.StaticText
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtFirst As SAPbouiCOM.DataTable
    Private dtSecond As SAPbouiCOM.DataTable
    Private dtThird As SAPbouiCOM.DataTable
    Private ocombo As SAPbouiCOM.ComboBoxColumn
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim oRecordSet As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()
        Try
            If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_JVApproval) = False Then
                oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            oForm = oApplication.Utilities.LoadForm(xml_JVApproval, frm_JVApproval)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            initialize(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_JVApproval Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 2) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    ElseIf RecordValidation(oForm) = False Then
                                        BubbleEvent = False
                                        oApplication.Utilities.Message("No Records Found...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 1) Then
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                    changeLabel(oForm)
                                ElseIf pVal.ItemUID = "3" And (oForm.PaneLevel = 2) Then
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                    changeLabel(oForm)
                                    loadApprovalData(oForm)
                                ElseIf pVal.ItemUID = "6" And (oForm.PaneLevel = 3 Or oForm.PaneLevel = 4 Or oForm.PaneLevel = 5 Or oForm.PaneLevel = 6) Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to Approve Documents...?", , "Continue", "Cancel") = 2 Then
                                    Else
                                        If JournalVoucherApproval(oForm) = True Then
                                            oApplication.Utilities.Message("Documents Approved Successfully...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            loadApprovalData(oForm)
                                        End If
                                    End If
                                ElseIf pVal.ItemUID = "1000004" Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 3
                                    changeLabel(oForm)
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "1000005" Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 4
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "1000006" Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 5
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "1000007" Then
                                    oForm.Freeze(True)
                                    oForm.PaneLevel = 6
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "4" Then
                                    oForm.Freeze(True)
                                    If oForm.PaneLevel <> 2 Then
                                        oForm.PaneLevel = 2
                                    Else
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                        changeLabel(oForm)
                                    End If
                                    oForm.Freeze(False)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Minimized Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    reDrawForm(oForm)
                                End If
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
                Case mnu_JVApproval
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Data Event"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
           
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Validations"
    Private Function Validation(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strFromDate, strToDate As String
            strFromDate = oApplication.Utilities.getEdittextvalue(oForm, "8")
            strToDate = oApplication.Utilities.getEdittextvalue(oForm, "10")
            If strFromDate = "" Then
                oApplication.Utilities.Message("Enter From Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strToDate = "" Then
                oApplication.Utilities.Message("Enter To Date ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Function RecordValidation(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strqry As String
            Dim strFromDate, strToDate As String
            strFromDate = oForm.Items.Item("8").Specific.value 'oApplication.Utilities.getEdittextvalue(oForm, "8")
            strToDate = oForm.Items.Item("10").Specific.value 'oApplication.Utilities.getEdittextvalue(oForm, "10")
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'strqry = " Select BatchNum "
            'strqry = strqry & " From OBTF T0 JOIN OUSR T1 On T1.UserID = T0.UserSign  "
            'strqry = strqry & " Where ((T1.U_FApprover = '" & oApplication.Company.UserName & "' And ISNULL(T0.U_FAproval,'O') = 'O') "
            'strqry = strqry & " OR (T1.U_SApprover = '" & oApplication.Company.UserName & "' And ISNULL(T0.U_SAproval,'O') = 'O') "
            'strqry = strqry & " OR (T1.U_TApprover = '" & oApplication.Company.UserName & "' And ISNULL(T0.U_TAproval,'O') = 'O') "
            'strqry = strqry & " OR (T1.U_FoApprover = '" & oApplication.Company.UserName & "' And ISNULL(T0.U_FoAproval,'O') = 'O')) "
            'strqry = strqry & " And Convert(VarChar(12),RefDate,112) >= '" & strFromDate & "' And Convert(VarChar(12),RefDate,112) <= '" & strToDate & "' "
            'strqry = strqry & " And T0.BtfStatus = 'O' "

            strqry = " Select BatchNum From OBTF T0 JOIN OUSR T1 On T1.UserID = T0.UserSign   "
            strqry = strqry & " JOIN [@JAT1] T2 On T1.User_Code = T2.U_OUser  "
            strqry = strqry & " JOIN [@JAT2] T3 On T2.Code = T3.Code "
            strqry = strqry & " Where "
            strqry = strqry & " Convert(VarChar(12),RefDate,112) >= '" & strFromDate & "' And Convert(VarChar(12),RefDate,112) <= '" & strToDate & "' And T0.BtfStatus = 'O'"
            strqry = strqry & " And "
            strqry = strqry & " ( "
            strqry = strqry & " (T3.U_AUser = '" & oApplication.Company.UserName & "' "
            strqry = strqry & " And ISNULL(T0.U_FAproval,'O') = 'O' And T3.LineId = '1') "
            strqry = strqry & " OR (T3.U_AUser = '" & oApplication.Company.UserName & "'  "
            strqry = strqry & " And ISNULL(T0.U_FAproval,'O') <> 'O' "
            strqry = strqry & " AND ISNULL(T0.U_SAproval,'O') = 'O' And T3.LineId = '2')  "
            strqry = strqry & " OR (T3.U_AUser = '" & oApplication.Company.UserName & "'  "
            strqry = strqry & " And ISNULL(T0.U_FAproval,'O') <> 'O'  "
            strqry = strqry & " AND ISNULL(T0.U_SAproval,'O') <> 'O' AND ISNULL(T0.U_TAproval,'O') = 'O'  "
            strqry = strqry & " And T3.LineId = '3')  "
            strqry = strqry & " OR (T3.U_AUser = '" & oApplication.Company.UserName & "' "
            strqry = strqry & " And ISNULL(T0.U_FAproval,'O') <> 'O' "
            strqry = strqry & " AND ISNULL(T0.U_SAproval,'O') <> 'O' "
            strqry = strqry & " AND ISNULL(T0.U_TAproval,'O') <> 'O' "
            strqry = strqry & " And ISNULL(T0.U_FoAproval,'O') = 'O' And T3.LineId = '4') "
            strqry = strqry & " ) "

            oRecordSet.DoQuery(strqry)
            If oRecordSet.EoF Then
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Items.Item("1").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.Items.Item("17").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.DataSources.DataTables.Add("dtFirst")
            oForm.DataSources.DataTables.Add("dtSecond")
            oForm.DataSources.DataTables.Add("dtThird")
            oForm.DataSources.DataTables.Add("dtFourth")
            changeLabel(oForm)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub loadApprovalData(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim blnHasData As Boolean = False
            FirstLevel(oForm)
            SecondLevel(oForm)
            ThirdLevel(oForm)
            FourthLevel(oForm)

            oGrid = oForm.Items.Item("11").Specific
            If oGrid.Rows.Count > 0 Then
                If oGrid.DataTable.GetValue("BatchNum", 0).ToString() <> "0" And oGrid.DataTable.GetValue("BatchNum", 0).ToString() <> "" Then
                    blnHasData = True
                End If
            End If

            oGrid = oForm.Items.Item("15").Specific
            If oGrid.Rows.Count > 0 Then
                If oGrid.DataTable.GetValue("BatchNum", 0).ToString() <> "0" And oGrid.DataTable.GetValue("BatchNum", 0).ToString() <> "" Then
                    blnHasData = True
                End If
            End If

            oGrid = oForm.Items.Item("16").Specific
            If oGrid.Rows.Count > 0 Then
                If oGrid.DataTable.GetValue("BatchNum", 0).ToString() <> "0" And oGrid.DataTable.GetValue("BatchNum", 0).ToString() <> "" Then
                    blnHasData = True
                End If
            End If

            oGrid = oForm.Items.Item("19").Specific
            If oGrid.Rows.Count > 0 Then
                If oGrid.DataTable.GetValue("BatchNum", 0).ToString() <> "0" And oGrid.DataTable.GetValue("BatchNum", 0).ToString() <> "" Then
                    blnHasData = True
                End If
            End If

            If blnHasData Then
                oForm.Items.Item("6").Enabled = True
            Else
                oForm.Items.Item("6").Enabled = False
            End If
            oForm.Items.Item("1000004").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub FirstLevel(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim strFromDate, strToDate As String

            strFromDate = aform.Items.Item("8").Specific.value 'oApplication.Utilities.getEdittextvalue(oForm, "8")
            strToDate = aform.Items.Item("10").Specific.value 'oApplication.Utilities.getEdittextvalue(oForm, "10")
            Dim strqry As String
            oGrid = aform.Items.Item("11").Specific
            oGrid.DataTable = aform.DataSources.DataTables.Item("dtFirst")

            strqry = " Select BatchNum,RefDate As DateID,LocTotal,T1.User_Code,T1.U_Name,T0.Memo As Remarks, "
            strqry = strqry & " Case When ISNULL(T0.U_FAproval,'O') = 'O' Then 'O' When ISNULL(T0.U_FAproval,'O') = 'A'  "
            strqry = strqry & " Then 'A' When ISNULL(T0.U_FAproval,'O') = 'R' Then 'R' End As 'FApproval',  "
            strqry = strqry & " T0.U_FAppRmks "
            strqry = strqry & " From OBTF T0 JOIN OUSR T1 On T1.UserID = T0.UserSign "
            strqry = strqry & " JOIN [@JAT1] T2 On T1.User_Code = T2.U_OUser "
            strqry = strqry & " JOIN [@JAT2] T3 On T2.Code = T3.Code "
            strqry = strqry & " Where T3.U_AUser = '" & oApplication.Company.UserName & "' And ISNULL(T0.U_FAproval,'O') = 'O' And T3.LineId = '1' "
            strqry = strqry & " And Convert(VarChar(12),RefDate,112) >= '" & strFromDate & "' And Convert(VarChar(12),RefDate,112) <= '" & strToDate & "' "
            strqry = strqry & " And T0.BtfStatus = 'O' "
            oGrid.DataTable.ExecuteQuery(strqry)

            oGrid.Columns.Item("BatchNum").TitleObject.Caption = "Voucher No"
            oGrid.Columns.Item("BatchNum").Editable = False
            oEditTextColumn = oGrid.Columns.Item("BatchNum")
            oEditTextColumn.LinkedObjectType = 28

            oGrid.Columns.Item("DateID").TitleObject.Caption = "Voucher Date"
            oGrid.Columns.Item("DateID").Editable = False

            oGrid.Columns.Item("LocTotal").TitleObject.Caption = "Voucher Amount"
            oGrid.Columns.Item("LocTotal").Editable = False

            oGrid.Columns.Item("User_Code").TitleObject.Caption = "User Code"
            oGrid.Columns.Item("User_Code").Editable = False

            oGrid.Columns.Item("U_Name").TitleObject.Caption = "Transaction User"
            oGrid.Columns.Item("U_Name").Editable = False

            oGrid.Columns.Item("Remarks").TitleObject.Caption = "Voucher Remarks"
            oGrid.Columns.Item("Remarks").Editable = False

            oGrid.Columns.Item("FApproval").TitleObject.Caption = "First Level Status"
            'oGrid.Columns.Item("FApproval").Editable = False
            oGrid.Columns.Item("FApproval").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            ocombo = oGrid.Columns.Item("FApproval")
            ocombo.ValidValues.Add("O", "Open")
            ocombo.ValidValues.Add("A", "Approved")
            ocombo.ValidValues.Add("R", "Rejected")
            ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
            oGrid.Columns.Item("U_FAppRmks").TitleObject.Caption = "First Level Remarks"

            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oApplication.Utilities.assignMatrixLineno(oGrid, aform)
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub SecondLevel(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim strqry As String
            Dim strFromDate, strToDate As String

            strFromDate = aform.Items.Item("8").Specific.value 'oApplication.Utilities.getEdittextvalue(oForm, "8")
            strToDate = aform.Items.Item("10").Specific.value 'oApplication.Utilities.getEdittextvalue(oForm, "10")

            oGrid = aform.Items.Item("15").Specific
            oGrid.DataTable = aform.DataSources.DataTables.Item("dtSecond")

            strqry = " Select BatchNum,RefDate As DateID,LocTotal,T1.User_Code,T1.U_Name,T0.Memo As Remarks,  "
            strqry = strqry & " Case When ISNULL(T0.U_FAproval,'O') = 'O' Then 'O' When ISNULL(T0.U_FAproval,'O') = 'A' "
            strqry = strqry & " Then 'A' When ISNULL(T0.U_FAproval,'O') = 'R' Then 'R' End As 'FApproval', "
            strqry = strqry & " T0.U_FApprover,T0.U_FApprTime,T0.U_FAppRmks , "
            strqry = strqry & " Case When ISNULL(T0.U_SAproval,'O') = 'O' Then 'O' When ISNULL(T0.U_SAproval,'O') = 'A'  "
            strqry = strqry & " Then 'A' When ISNULL(T0.U_SAproval,'O') = 'R' Then 'R' End As 'SApproval', "
            strqry = strqry & " T0.U_SAppRmks "
            strqry = strqry & " From OBTF T0 JOIN OUSR T1 On T1.UserID = T0.UserSign  "
            strqry = strqry & " JOIN [@JAT1] T2 On T1.User_Code = T2.U_OUser "
            strqry = strqry & " JOIN [@JAT2] T3 On T2.Code = T3.Code  "
            strqry = strqry & " Where T3.U_AUser = '" & oApplication.Company.UserName & "' And ISNULL(T0.U_FAproval,'O') <> 'O' And T3.LineId = '2'"
            strqry = strqry & " And ISNULL(T0.U_SAproval,'O') = 'O' "
            strqry = strqry & " And Convert(VarChar(12),RefDate,112) >= '" & strFromDate & "' And Convert(VarChar(12),RefDate,112) <= '" & strToDate & "' "
            strqry = strqry & " And T0.BtfStatus = 'O' "
            oGrid.DataTable.ExecuteQuery(strqry)


            oGrid.Columns.Item("BatchNum").TitleObject.Caption = "Voucher No"
            oGrid.Columns.Item("BatchNum").Editable = False
            oEditTextColumn = oGrid.Columns.Item("BatchNum")
            oEditTextColumn.LinkedObjectType = 28

            oGrid.Columns.Item("DateID").TitleObject.Caption = "Voucher Date"
            oGrid.Columns.Item("DateID").Editable = False

            oGrid.Columns.Item("LocTotal").TitleObject.Caption = "Voucher Amount"
            oGrid.Columns.Item("LocTotal").Editable = False

            oGrid.Columns.Item("User_Code").TitleObject.Caption = "User Code"
            oGrid.Columns.Item("User_Code").Editable = False

            oGrid.Columns.Item("U_Name").TitleObject.Caption = "Transaction User"
            oGrid.Columns.Item("U_Name").Editable = False

            oGrid.Columns.Item("Remarks").TitleObject.Caption = "Voucher Remarks"
            oGrid.Columns.Item("Remarks").Editable = False

            oGrid.Columns.Item("U_FApprover").TitleObject.Caption = "First Level Approver"
            oGrid.Columns.Item("U_FApprover").Editable = False
            oGrid.Columns.Item("FApproval").TitleObject.Caption = "First Level Status"
            oGrid.Columns.Item("FApproval").Editable = False
            oGrid.Columns.Item("FApproval").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            ocombo = oGrid.Columns.Item("FApproval")
            ocombo.ValidValues.Add("O", "Open")
            ocombo.ValidValues.Add("A", "Approved")
            ocombo.ValidValues.Add("R", "Rejected")
            ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            oGrid.Columns.Item("U_FApprTime").TitleObject.Caption = "First Level Approver Time"
            oGrid.Columns.Item("U_FApprTime").Editable = False

            oGrid.Columns.Item("U_FAppRmks").TitleObject.Caption = "First Level Remarks"
            oGrid.Columns.Item("U_FAppRmks").Editable = False

            oGrid.Columns.Item("SApproval").TitleObject.Caption = "Second Level Status"
            'oGrid.Columns.Item("SApproval").Editable = False
            oGrid.Columns.Item("SApproval").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            ocombo = oGrid.Columns.Item("SApproval")
            ocombo.ValidValues.Add("O", "Open")
            ocombo.ValidValues.Add("A", "Approved")
            ocombo.ValidValues.Add("R", "Rejected")
            ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
            oGrid.Columns.Item("U_SAppRmks").TitleObject.Caption = "Second Level Remarks"

            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oApplication.Utilities.assignMatrixLineno(oGrid, aform)
            aform.Freeze(False)

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Private Sub ThirdLevel(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim strqry As String
            Dim strFromDate, strToDate As String

            strFromDate = aform.Items.Item("8").Specific.value 'oApplication.Utilities.getEdittextvalue(oForm, "8")
            strToDate = aform.Items.Item("10").Specific.value 'oApplication.Utilities.getEdittextvalue(oForm, "10")

            oGrid = aform.Items.Item("16").Specific
            oGrid.DataTable = aform.DataSources.DataTables.Item("dtThird")

            strqry = " Select BatchNum,RefDate As DateID,LocTotal,T1.User_Code,T1.U_Name,T0.Memo As Remarks,  "
            strqry = strqry & " Case When ISNULL(T0.U_FAproval,'O') = 'O' Then 'O' When ISNULL(T0.U_FAproval,'O') = 'A' "
            strqry = strqry & " Then 'A' When ISNULL(T0.U_FAproval,'O') = 'R' Then 'R' End As 'FApproval', "
            strqry = strqry & " T0.U_FApprover,T0.U_FApprTime,T0.U_FAppRmks , "
            strqry = strqry & " Case When ISNULL(T0.U_SAproval,'O') = 'O' Then 'O' When ISNULL(T0.U_SAproval,'O') = 'A'  "
            strqry = strqry & " Then 'A' When ISNULL(T0.U_SAproval,'O') = 'R' Then 'R' End As 'SApproval', "
            strqry = strqry & " T0.U_SApprover,T0.U_SApprTime,T0.U_SAppRmks , "
            strqry = strqry & " Case When ISNULL(T0.U_TAproval,'O') = 'O' Then 'O' When ISNULL(T0.U_TAproval,'O') = 'A'  "
            strqry = strqry & " Then 'A' When ISNULL(T0.U_TAproval,'O') = 'R' Then 'R' End As 'TApproval', "
            strqry = strqry & " T0.U_TAppRmks "
            strqry = strqry & " From OBTF T0 JOIN OUSR T1 On T1.UserID = T0.UserSign  "
            strqry = strqry & " JOIN [@JAT1] T2 On T1.User_Code = T2.U_OUser  "
            strqry = strqry & " JOIN [@JAT2] T3 On T2.Code = T3.Code  "
            strqry = strqry & " Where T3.U_AUser = '" & oApplication.Company.UserName & "' And ISNULL(T0.U_FAproval,'O') <> 'O' "
            strqry = strqry & " And ISNULL(T0.U_SAproval,'O') <> 'O' "
            strqry = strqry & " And ISNULL(T0.U_TAproval,'O') = 'O' And T3.LineId = '3' "
            strqry = strqry & " And Convert(VarChar(12),RefDate,112) >= '" & strFromDate & "' And Convert(VarChar(12),RefDate,112) <= '" & strToDate & "' "
            strqry = strqry & " And T0.BtfStatus = 'O' "
            oGrid.DataTable.ExecuteQuery(strqry)


            oGrid.Columns.Item("BatchNum").TitleObject.Caption = "Voucher No"
            oGrid.Columns.Item("BatchNum").Editable = False
            oEditTextColumn = oGrid.Columns.Item("BatchNum")
            oEditTextColumn.LinkedObjectType = 28

            oGrid.Columns.Item("DateID").TitleObject.Caption = "Voucher Date"
            oGrid.Columns.Item("DateID").Editable = False

            oGrid.Columns.Item("LocTotal").TitleObject.Caption = "Voucher Amount"
            oGrid.Columns.Item("LocTotal").Editable = False

            oGrid.Columns.Item("User_Code").TitleObject.Caption = "User Code"
            oGrid.Columns.Item("User_Code").Editable = False

            oGrid.Columns.Item("U_Name").TitleObject.Caption = "Transaction User"
            oGrid.Columns.Item("U_Name").Editable = False

            oGrid.Columns.Item("Remarks").TitleObject.Caption = "Voucher Remarks"
            oGrid.Columns.Item("Remarks").Editable = False

            oGrid.Columns.Item("FApproval").TitleObject.Caption = "First Level Status"
            oGrid.Columns.Item("FApproval").Editable = False
            oGrid.Columns.Item("FApproval").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            ocombo = oGrid.Columns.Item("FApproval")
            ocombo.ValidValues.Add("O", "Open")
            ocombo.ValidValues.Add("A", "Approved")
            ocombo.ValidValues.Add("R", "Rejected")
            ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            oGrid.Columns.Item("U_FAppRmks").TitleObject.Caption = "First Level Remarks"
            oGrid.Columns.Item("U_FAppRmks").Editable = False

            oGrid.Columns.Item("U_FApprTime").TitleObject.Caption = "First Level Approver Time"
            oGrid.Columns.Item("U_FApprTime").Editable = False

            oGrid.Columns.Item("U_FApprover").TitleObject.Caption = "First Level Approver"
            oGrid.Columns.Item("U_FApprover").Editable = False

            oGrid.Columns.Item("U_SApprover").TitleObject.Caption = "Second Level Approver"
            oGrid.Columns.Item("U_SApprover").Editable = False

            oGrid.Columns.Item("SApproval").TitleObject.Caption = "Second Level Status"
            oGrid.Columns.Item("SApproval").Editable = False
            oGrid.Columns.Item("SApproval").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            ocombo = oGrid.Columns.Item("SApproval")
            ocombo.ValidValues.Add("O", "Open")
            ocombo.ValidValues.Add("A", "Approved")
            ocombo.ValidValues.Add("R", "Rejected")
            ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            oGrid.Columns.Item("U_SApprTime").TitleObject.Caption = "Second Level Approver Time"
            oGrid.Columns.Item("U_SApprTime").Editable = False

            oGrid.Columns.Item("U_SAppRmks").TitleObject.Caption = "Second Level Remarks"
            oGrid.Columns.Item("U_SAppRmks").Editable = False

            oGrid.Columns.Item("TApproval").TitleObject.Caption = "Third Level Status"
            'oGrid.Columns.Item("TApproval").Editable = False
            oGrid.Columns.Item("TApproval").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            ocombo = oGrid.Columns.Item("TApproval")
            ocombo.ValidValues.Add("O", "Open")
            ocombo.ValidValues.Add("A", "Approved")
            ocombo.ValidValues.Add("R", "Rejected")
            ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
            oGrid.Columns.Item("U_TAppRmks").TitleObject.Caption = "Third Level Remarks"


            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oApplication.Utilities.assignMatrixLineno(oGrid, aform)
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub FourthLevel(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Dim strqry As String
            Dim strFromDate, strToDate As String

            strFromDate = aform.Items.Item("8").Specific.value 'oApplication.Utilities.getEdittextvalue(oForm, "8")
            strToDate = aform.Items.Item("10").Specific.value 'oApplication.Utilities.getEdittextvalue(oForm, "10")

            oGrid = aform.Items.Item("19").Specific
            oGrid.DataTable = aform.DataSources.DataTables.Item("dtFourth")

            strqry = " Select BatchNum,RefDate As DateID,LocTotal,T1.User_Code,T1.U_Name,T0.Memo As Remarks,  "
            strqry = strqry & " Case When ISNULL(T0.U_FAproval,'O') = 'O' Then 'O' When ISNULL(T0.U_FAproval,'O') = 'A' "
            strqry = strqry & " Then 'A' When ISNULL(T0.U_FAproval,'O') = 'R' Then 'R' End As 'FApproval', "
            strqry = strqry & " T0.U_FApprover,T0.U_FApprTime,T0.U_FAppRmks , "
            strqry = strqry & " Case When ISNULL(T0.U_SAproval,'O') = 'O' Then 'O' When ISNULL(T0.U_SAproval,'O') = 'A'  "
            strqry = strqry & " Then 'A' When ISNULL(T0.U_SAproval,'O') = 'R' Then 'R' End As 'SApproval', "
            strqry = strqry & " T0.U_SApprover,T0.U_SApprTime,T0.U_SAppRmks , "
            strqry = strqry & " Case When ISNULL(T0.U_TAproval,'O') = 'O' Then 'O' When ISNULL(T0.U_TAproval,'O') = 'A'  "
            strqry = strqry & " Then 'A' When ISNULL(T0.U_TAproval,'O') = 'R' Then 'R' End As 'TApproval', "
            strqry = strqry & " T0.U_TApprover,T0.U_TApprTime,T0.U_TAppRmks ,"
            strqry = strqry & " Case When ISNULL(T0.U_FoAproval,'O') = 'O' Then 'O' When ISNULL(T0.U_FoAproval,'O') = 'A'  "
            strqry = strqry & " Then 'A' When ISNULL(T0.U_FoAproval,'O') = 'R' Then 'R' End As 'FoApproval', "
            strqry = strqry & " T0.U_FoAppRmks "
            strqry = strqry & " From OBTF T0 JOIN OUSR T1 On T1.UserID = T0.UserSign  "
            strqry = strqry & " JOIN [@JAT1] T2 On T1.User_Code = T2.U_OUser "
            strqry = strqry & " JOIN [@JAT2] T3 On T2.Code = T3.Code   "
            strqry = strqry & " Where T3.U_AUser = '" & oApplication.Company.UserName & "' And ISNULL(T0.U_FAproval,'O') <> 'O' "
            strqry = strqry & " And ISNULL(T0.U_SAproval,'O') <> 'O' "
            strqry = strqry & " And ISNULL(T0.U_TAproval,'O') <> 'O' "
            strqry = strqry & " And ISNULL(T0.U_FoAproval,'O') = 'O' And T3.LineId = '4' "
            strqry = strqry & " And Convert(VarChar(12),RefDate,112) >= '" & strFromDate & "' And Convert(VarChar(12),RefDate,112) <= '" & strToDate & "' "
            strqry = strqry & " And T0.BtfStatus = 'O' "
            oGrid.DataTable.ExecuteQuery(strqry)


            oGrid.Columns.Item("BatchNum").TitleObject.Caption = "Voucher No"
            oGrid.Columns.Item("BatchNum").Editable = False
            oEditTextColumn = oGrid.Columns.Item("BatchNum")
            oEditTextColumn.LinkedObjectType = 28

            oGrid.Columns.Item("DateID").TitleObject.Caption = "Voucher Date"
            oGrid.Columns.Item("DateID").Editable = False

            oGrid.Columns.Item("LocTotal").TitleObject.Caption = "Voucher Amount"
            oGrid.Columns.Item("LocTotal").Editable = False

            oGrid.Columns.Item("User_Code").TitleObject.Caption = "User Code"
            oGrid.Columns.Item("User_Code").Editable = False

            oGrid.Columns.Item("U_Name").TitleObject.Caption = "Transaction User"
            oGrid.Columns.Item("U_Name").Editable = False

            oGrid.Columns.Item("Remarks").TitleObject.Caption = "Voucher Remarks"
            oGrid.Columns.Item("Remarks").Editable = False

            oGrid.Columns.Item("FApproval").TitleObject.Caption = "First Level Status"
            oGrid.Columns.Item("FApproval").Editable = False
            oGrid.Columns.Item("FApproval").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            ocombo = oGrid.Columns.Item("FApproval")
            ocombo.ValidValues.Add("O", "Open")
            ocombo.ValidValues.Add("A", "Approved")
            ocombo.ValidValues.Add("R", "Rejected")
            ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            oGrid.Columns.Item("U_FAppRmks").TitleObject.Caption = "First Level Remarks"
            oGrid.Columns.Item("U_FAppRmks").Editable = False

            oGrid.Columns.Item("U_FApprTime").TitleObject.Caption = "First Level Approver Time"
            oGrid.Columns.Item("U_FApprTime").Editable = False

            oGrid.Columns.Item("U_FApprover").TitleObject.Caption = "First Level Approver"
            oGrid.Columns.Item("U_FApprover").Editable = False

            oGrid.Columns.Item("U_SApprover").TitleObject.Caption = "Second Level Approver"
            oGrid.Columns.Item("U_SApprover").Editable = False

            oGrid.Columns.Item("SApproval").TitleObject.Caption = "Second Level Status"
            oGrid.Columns.Item("SApproval").Editable = False
            oGrid.Columns.Item("SApproval").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            ocombo = oGrid.Columns.Item("SApproval")
            ocombo.ValidValues.Add("O", "Open")
            ocombo.ValidValues.Add("A", "Approved")
            ocombo.ValidValues.Add("R", "Rejected")
            ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

            oGrid.Columns.Item("U_SApprTime").TitleObject.Caption = "Second Level Approver Time"
            oGrid.Columns.Item("U_SApprTime").Editable = False

            oGrid.Columns.Item("U_SAppRmks").TitleObject.Caption = "Second Level Remarks"
            oGrid.Columns.Item("U_SAppRmks").Editable = False

            oGrid.Columns.Item("U_TApprover").TitleObject.Caption = "Third Level Approver"
            oGrid.Columns.Item("U_TApprover").Editable = False

            oGrid.Columns.Item("TApproval").TitleObject.Caption = "Third Level Status"
            oGrid.Columns.Item("TApproval").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            ocombo = oGrid.Columns.Item("TApproval")
            ocombo.ValidValues.Add("O", "Open")
            ocombo.ValidValues.Add("A", "Approved")
            ocombo.ValidValues.Add("R", "Rejected")
            ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
            oGrid.Columns.Item("TApproval").Editable = False

            oGrid.Columns.Item("U_TApprTime").TitleObject.Caption = "Third Level Approver Time"
            oGrid.Columns.Item("U_TApprTime").Editable = False

            oGrid.Columns.Item("U_TAppRmks").TitleObject.Caption = "Third Level Remarks"
            oGrid.Columns.Item("U_TAppRmks").Editable = False

            oGrid.Columns.Item("FoApproval").TitleObject.Caption = "Fourth Level Status"
            oGrid.Columns.Item("FoApproval").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            ocombo = oGrid.Columns.Item("FoApproval")
            ocombo.ValidValues.Add("O", "Open")
            ocombo.ValidValues.Add("A", "Approved")
            ocombo.ValidValues.Add("R", "Rejected")
            ocombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
            oGrid.Columns.Item("U_FoAppRmks").TitleObject.Caption = "Fourth Level Remarks"

            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oApplication.Utilities.assignMatrixLineno(oGrid, aform)
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Function JournalVoucherApproval(ByVal aform As SAPbouiCOM.Form) As Boolean
        Try
            Dim sQuery As String
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            'First Level Approval
            oGrid = oForm.Items.Item("11").Specific
            If oGrid.Rows.Count Then
                For index As Integer = 0 To oGrid.Rows.Count - 1
                    If oGrid.DataTable.GetValue("FApproval", index) <> "O" And oGrid.DataTable.GetValue("BatchNum", index).ToString() <> "" Then
                        sQuery = "Update OBTF Set U_FApprover='" & oApplication.Company.UserName & "', U_FAproval = '" & oGrid.DataTable.GetValue("FApproval", index).ToString() & "', U_FApprTime = '" & System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") & "', U_FAppRmks = '" & oGrid.DataTable.GetValue("U_FAppRmks", index) & "' Where BatchNum = '" & oGrid.DataTable.GetValue("BatchNum", index) & "'"
                        oRecordSet.DoQuery(sQuery)

                        'Final Status Update
                        sQuery = "Select T1.Code from [@JAT2] T1 JOIN [@JAT1] T0 ON T0.Code = T1.Code Where T0.U_OUser = '" + oGrid.DataTable.GetValue("User_Code", index) + "'"
                        oRecordSet.DoQuery(sQuery)
                        If oRecordSet.RecordCount = 1 Then
                            sQuery = "Update OBTF Set U_Approval = '" & oGrid.DataTable.GetValue("FApproval", index).ToString() & "' Where BatchNum = '" & oGrid.DataTable.GetValue("BatchNum", index) & "'"
                            oRecordSet.DoQuery(sQuery)
                        End If
                    End If
                Next
            End If

            'Second Level Approval
            oGrid = oForm.Items.Item("15").Specific
            If oGrid.Rows.Count Then
                For index As Integer = 0 To oGrid.Rows.Count - 1
                    If oGrid.DataTable.GetValue("SApproval", index) <> "O" And oGrid.DataTable.GetValue("BatchNum", index).ToString() <> "" Then
                        sQuery = "Update OBTF Set  U_SApprover='" & oApplication.Company.UserName & "', U_SAproval = '" & oGrid.DataTable.GetValue("SApproval", index).ToString() & "', U_SApprTime = '" & System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") & "', U_SAppRmks = '" & oGrid.DataTable.GetValue("U_SAppRmks", index) & "' Where BatchNum = '" & oGrid.DataTable.GetValue("BatchNum", index) & "'"
                        oRecordSet.DoQuery(sQuery)

                        'Final Status Update
                        sQuery = "Select T1.Code from [@JAT2] T1 JOIN [@JAT1] T0 ON T0.Code = T1.Code Where T0.U_OUser = '" + oGrid.DataTable.GetValue("User_Code", index) + "'"
                        oRecordSet.DoQuery(sQuery)
                        If oRecordSet.RecordCount = 2 Then
                            sQuery = "Update OBTF Set U_Approval = '" & oGrid.DataTable.GetValue("SApproval", index).ToString() & "' Where BatchNum = '" & oGrid.DataTable.GetValue("BatchNum", index) & "'"
                            oRecordSet.DoQuery(sQuery)
                        End If
                    End If
                Next
            End If

            'Third Level Approval
            oGrid = oForm.Items.Item("16").Specific
            If oGrid.Rows.Count Then
                For index As Integer = 0 To oGrid.Rows.Count - 1
                    If oGrid.DataTable.GetValue("TApproval", index) <> "O" And oGrid.DataTable.GetValue("BatchNum", index).ToString() <> "" Then
                        sQuery = "Update OBTF Set  U_TApprover='" & oApplication.Company.UserName & "', U_TAproval = '" & oGrid.DataTable.GetValue("TApproval", index).ToString() & "', U_TApprTime = '" & System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") & "', U_TAppRmks = '" & oGrid.DataTable.GetValue("U_TAppRmks", index) & "' Where BatchNum = '" & oGrid.DataTable.GetValue("BatchNum", index) & "'"
                        oRecordSet.DoQuery(sQuery)

                        'Final Status Update
                        sQuery = "Select T1.Code from [@JAT2] T1 JOIN [@JAT1] T0 ON T0.Code = T1.Code Where T0.U_OUser = '" + oGrid.DataTable.GetValue("User_Code", index) + "'"
                        oRecordSet.DoQuery(sQuery)
                        If oRecordSet.RecordCount = 3 Then
                            sQuery = "Update OBTF Set U_Approval = '" & oGrid.DataTable.GetValue("TApproval", index).ToString() & "' Where BatchNum = '" & oGrid.DataTable.GetValue("BatchNum", index) & "'"
                            oRecordSet.DoQuery(sQuery)
                        End If
                    End If
                Next
            End If

            'Fourth Level Approval
            oGrid = oForm.Items.Item("19").Specific
            If oGrid.Rows.Count Then
                For index As Integer = 0 To oGrid.Rows.Count - 1
                    If oGrid.DataTable.GetValue("FoApproval", index) <> "O" And oGrid.DataTable.GetValue("BatchNum", index).ToString() <> "" Then
                        sQuery = "Update OBTF Set U_FoApprover='" & oApplication.Company.UserName & "', U_FoAproval = '" & oGrid.DataTable.GetValue("FoApproval", index).ToString() & "', U_FoApprTime = '" & System.DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss") & "', U_FoAppRmks = '" & oGrid.DataTable.GetValue("U_FoAppRmks", index) & "' Where BatchNum = '" & oGrid.DataTable.GetValue("BatchNum", index) & "'"
                        oRecordSet.DoQuery(sQuery)

                        'Final Status Update
                        sQuery = "Select T1.Code from [@JAT2] T1 JOIN [@JAT1] T0 ON T0.Code = T1.Code Where T0.U_OUser = '" + oGrid.DataTable.GetValue("User_Code", index) + "'"
                        oRecordSet.DoQuery(sQuery)
                        If oRecordSet.RecordCount = 4 Then
                            sQuery = "Update OBTF Set U_Approval = '" & oGrid.DataTable.GetValue("FoApproval", index).ToString() & "' Where BatchNum = '" & oGrid.DataTable.GetValue("BatchNum", index) & "'"
                            oRecordSet.DoQuery(sQuery)
                        End If
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            oForm.Items.Item("5").Width = oForm.Width - 30
            oForm.Items.Item("5").Height = oForm.Height - 100
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub changeLabel(ByVal oForm As SAPbouiCOM.Form)
        Try
            oStatic = oForm.Items.Item("17").Specific
            oStatic.Caption = "Step " & oForm.PaneLevel & " of 4"
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

End Class
