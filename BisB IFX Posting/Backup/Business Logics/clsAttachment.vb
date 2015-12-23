Imports System.IO
Public Class clsAttachment
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm(ByVal aRefCode As String)
        'If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_IFXPosting) = False Then
        '    oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Exit Sub
        'End If
        oForm = oApplication.Utilities.LoadForm(xml_Attachment, frm_Attachment)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        Try
            oForm.Freeze(True)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oApplication.Utilities.setEdittextvalue(oForm, "5", aRefCode)
            DataBind(oForm)
            '  oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub DataBind(ByVal aForm As SAPbouiCOM.Form)
        Dim strCode As String = oApplication.Utilities.getEdittextvalue(aForm, "5")
        oGrid = aForm.Items.Item("1").Specific
        oGrid.DataTable.ExecuteQuery("Select Code,Name,U_Z_RefCode,U_Z_FileName,U_Z_FilePath,U_Z_CreatedBy ,U_Z_CreateDate  from [@Z_JVATT] where U_Z_RefCode='" & strCode & "'")
        oGrid.Columns.Item("Code").Visible = False
        oGrid.Columns.Item("Name").Visible = False
        oGrid.Columns.Item("U_Z_RefCode").Visible = False
        oGrid.Columns.Item("U_Z_FileName").TitleObject.Caption = "File Name"
        oGrid.Columns.Item("U_Z_FileName").Editable = False
        oGrid.Columns.Item("U_Z_FilePath").TitleObject.Caption = "Path"
        oGrid.Columns.Item("U_Z_FilePath").Editable = False
        oGrid.Columns.Item("U_Z_CreatedBy").TitleObject.Caption = "Created By"
        oGrid.Columns.Item("U_Z_CreatedBy").Editable = False
        oGrid.Columns.Item("U_Z_CreateDate").TitleObject.Caption = "Creation Date"
        oGrid.Columns.Item("U_Z_CreateDate").Editable = False
        oGrid.RowHeaders.TitleObject.Caption = "#"
        oGrid.AutoResizeColumns()
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oGrid.RowHeaders.SetText(intRow, intRow + 1)
        Next

    End Sub
    Private Sub DeleteRow(ByVal aForm As SAPbouiCOM.Form)
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oGrid = aForm.Items.Item("1").Specific
        If oApplication.SBO_Application.MessageBox("Do you want to remove the selected attachment ? ", , "Continue", "Cancel") = 2 Then
            Exit Sub
        End If
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.Rows.IsSelected(intRow) Then
                If oGrid.DataTable.GetValue("Code", intRow) <> "" Then
                    oRecordSet.DoQuery("Update [@Z_JVATT] set Name=Name +'_XD' where Code='" & oGrid.DataTable.GetValue("Code", intRow) & "'")
                    oGrid.DataTable.Rows.Remove(intRow)
                    For intRow1 As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                        oGrid.RowHeaders.SetText(intRow1, intRow1 + 1)
                    Next
                    Exit Sub
                Else
                    oGrid.DataTable.Rows.Remove(intRow)
                    For intRow1 As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                        oGrid.RowHeaders.SetText(intRow1, intRow1 + 1)
                    Next
                    Exit Sub
                End If
            End If
        Next
        oApplication.Utilities.Message("No rows selected", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

    End Sub
    Private Sub AddtoUDT(ByVal aForm As SAPbouiCOM.Form)
        Dim strRefCode, strCode, strFileName, strFilePath As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oUserTable As SAPbobsCOM.UserTable
        oUserTable = oApplication.Company.UserTables.Item("Z_JVATT")
        oGrid = aForm.Items.Item("1").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.DataTable.GetValue("U_Z_FileName", intRow) <> "" Then
                If oGrid.DataTable.GetValue("Code", intRow) = "" Then
                    strCode = oApplication.Utilities.getMaxCode("@Z_JVATT", "Code")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_RefCode").Value = oApplication.Utilities.getEdittextvalue(aForm, "5")
                    oUserTable.UserFields.Fields.Item("U_Z_FileName").Value = oGrid.DataTable.GetValue("U_Z_FileName", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_FilePath").Value = oGrid.DataTable.GetValue("U_Z_FilePath", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_CreatedBy").Value = oApplication.Company.UserName
                    oUserTable.UserFields.Fields.Item("U_Z_CreateDate").Value = Now.Date
                    If oUserTable.Add() <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                Else
                    strCode = oGrid.DataTable.GetValue("Code", intRow)
                    oUserTable.GetByKey(strCode)
                    oUserTable.UserFields.Fields.Item("U_Z_RefCode").Value = oApplication.Utilities.getEdittextvalue(aForm, "5")
                    oUserTable.UserFields.Fields.Item("U_Z_FileName").Value = oGrid.DataTable.GetValue("U_Z_FileName", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_FilePath").Value = oGrid.DataTable.GetValue("U_Z_FilePath", intRow)
                    If oUserTable.Update() <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Exit Sub
                    End If
                End If
            End If
        Next
        CopyAttachment(aForm)
        oRecordSet.DoQuery("Delete from [@Z_JVATT] where Name like '%_XD' and U_Z_RefCode ='" & oApplication.Utilities.getEdittextvalue(aForm, "5") & "'")
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
    Private Sub CopyAttachment(ByVal aForm As SAPbouiCOM.Form)
        oGrid = aForm.Items.Item("1").Specific
        For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            '  Dim odr As DataRow = oDT.Rows(i)
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQry = "Select AttachPath From OADP"
            oRec.DoQuery(strQry)
            Dim SPath As String = oGrid.DataTable.GetValue("U_Z_FilePath", i)
            If SPath = "" Then
            Else
                Dim DPath As String = ""
                If Not oRec.EoF Then
                    DPath = oRec.Fields.Item("AttachPath").Value.ToString()
                End If
                If Not Directory.Exists(DPath) Then
                    Directory.CreateDirectory(DPath)
                End If
                Dim file = New FileInfo(SPath)
                Dim Filename As String = Path.GetFileName(SPath)
                Dim SavePath As String = Path.Combine(DPath, Filename)
                If System.IO.File.Exists(SPath) Then

                    If System.IO.File.Exists(SavePath) Then
                    Else
                        If System.IO.File.Exists(SPath) Then
                            file.CopyTo(Path.Combine(DPath, file.Name), True)
                        End If
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub LoadAttachment(ByVal aForm As SAPbouiCOM.Form)
        oGrid = aForm.Items.Item("1").Specific
        For i As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            '  Dim odr As DataRow = oDT.Rows(i)
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQry = "Select AttachPath From OADP"
            oRec.DoQuery(strQry)
            If oGrid.Rows.IsSelected(i) Then
                Dim strFilename As String
                strFilename = oGrid.DataTable.GetValue("U_Z_FileName", i)
                Dim SPath As String = oGrid.DataTable.GetValue("U_Z_FilePath", i)
                Dim DPath As String = ""
                If Not oRec.EoF Then
                    DPath = oRec.Fields.Item("AttachPath").Value.ToString()
                End If
                If Not Directory.Exists(DPath) Then
                    Directory.CreateDirectory(DPath)
                End If
                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                Dim Filename As String = Path.GetFileName(SPath)
                Dim SavePath As String = Path.Combine(DPath, Filename)
                If System.IO.File.Exists(strFilename) Then
                    strFilename = strFilename
                Else
                    strFilename = SavePath
                End If
                x.FileName = strFilename
                If System.IO.File.Exists(strFilename) Then
                    System.Diagnostics.Process.Start(x)
                Else
                    oApplication.Utilities.Message("File does not exits", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If

                x = Nothing
                Exit Sub
            End If
        Next
    End Sub

#Region "FileOpen"
    Private Sub FileOpen()
        Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
        mythr.SetApartmentState(Threading.ApartmentState.STA)
        mythr.Start()
        mythr.Join()
    End Sub

    Private Sub ShowFileDialog()
        Dim oDialogBox As New OpenFileDialog

        Dim oProcesses() As Process
        Try
            oProcesses = Process.GetProcessesByName("SAP Business One")
            If oProcesses.Length <> 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    Dim MyWindow As New WindowWrapper(oProcesses(i).MainWindowHandle)
                    If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                        strMdbFilePath = oDialogBox.SafeFileName
                        strFilepath = oDialogBox.FileName

                    Else
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
#End Region
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Attachment Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "2" Then
                                    Dim oRecordset As SAPbobsCOM.Recordset
                                    oRecordset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordset.DoQuery("Update [@Z_JVATT] set Name=Code where U_Z_RefCode='" & oApplication.Utilities.getEdittextvalue(oForm, "5") & "'")
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "3"
                                        AddtoUDT(oForm)
                                    Case "6" ' Browse
                                        oGrid = oForm.Items.Item("1").Specific
                                        FileOpen()
                                        If strFilepath = "" Then
                                            '  oApplication.Utilities.Message("Please Select a File", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            ' BubbleEvent = False
                                        Else
                                            Try
                                                oForm.Freeze(True)
                                                If oGrid.DataTable.Rows.Count - 1 < 0 Then
                                                    oGrid.DataTable.Rows.Add()
                                                End If
                                                If oGrid.DataTable.GetValue("U_Z_FileName", oGrid.DataTable.Rows.Count - 1) <> "" Then
                                                    oGrid.DataTable.Rows.Add()
                                                End If
                                                oGrid.DataTable.SetValue("U_Z_FileName", oGrid.DataTable.Rows.Count - 1, strMdbFilePath)
                                                oGrid.DataTable.SetValue("U_Z_FilePath", oGrid.DataTable.Rows.Count - 1, strFilepath)
                                                oGrid.DataTable.SetValue("U_Z_CreatedBy", oGrid.DataTable.Rows.Count - 1, oApplication.Company.UserName)
                                                oGrid.DataTable.SetValue("U_Z_CreateDate", oGrid.DataTable.Rows.Count - 1, Now.Date)
                                                For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                                                    oGrid.RowHeaders.SetText(intRow, intRow + 1)
                                                Next
                                                oForm.Freeze(False)
                                            Catch ex As Exception
                                                oForm.Freeze(False)
                                            End Try
                                        End If
                                    Case "7" 'Display Attachment
                                        LoadAttachment(oForm)
                                    Case "8" 'Delete Attachment
                                        DeleteRow(oForm)
                                End Select

                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Events"

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        If oForm.TypeEx = frm_JournalEntry Or oForm.TypeEx = frm_JournalVoucher Then
            If (eventInfo.BeforeAction = True) Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_PRINT_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "Z_Att"
                        oCreationPackage.String = "Attachment"
                        oCreationPackage.Enabled = True
                        oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)

                        'oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        'oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        'oCreationPackage.UniqueID = "Training"
                        'oCreationPackage.String = "Training Details"
                        'oCreationPackage.Enabled = True
                        'oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                        'oMenus = oMenuItem.SubMenus
                        'oMenus.AddEx(oCreationPackage)
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Else
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    ' If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    oApplication.SBO_Application.Menus.RemoveEx("Z_Att")
                    ' End If

                Catch ex As Exception
                    ' MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case "Z_Att"
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If oForm.TypeEx = frm_JournalEntry Or oForm.TypeEx = frm_JournalVoucher Then
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            oDBDataSource = oForm.DataSources.DBDataSources.Item(0)
                            Dim strCode As String = oDBDataSource.GetValue("U_Z_AttRef", 0) ' oApplication.Utilities.getEdittextvalue(oForm, "_10021")
                            If strCode = "" Then
                                Dim oRecSet As SAPbobsCOM.Recordset
                                oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                strCode = oApplication.Utilities.getMaxCode("@Z_OJVATT", "Code")
                                oRecSet.DoQuery("Insert into [@Z_OJVATT] values ('" & strCode & "','" & strCode & "')")
                                oApplication.Utilities.setEdittextvalue(oForm, "_10021", strCode)
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And oForm.TypeEx = frm_JournalEntry Then
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                End If
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And oForm.TypeEx = frm_JournalVoucher Then
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                                End If


                            End If
                            LoadForm(strCode)

                        End If
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
