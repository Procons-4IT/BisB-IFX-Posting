Imports System.IO
Public Class clsListener
    Inherits Object
    Private ThreadClose As New Threading.Thread(AddressOf CloseApp)
    Private WithEvents _SBO_Application As SAPbouiCOM.Application
    Private _Company As SAPbobsCOM.Company
    Private _RemoteCompany As SAPbobsCOM.Company
    Private _Utilities As clsUtilities
    Private _Collection As Hashtable
    Private _LookUpCollection As Hashtable
    Private _FormUID As String
    ' Private _Log As clsLog_Error
    Private oMenuObject As Object
    Private oItemObject As Object
    Private oSystemForms As Object
    Dim objFilters As SAPbouiCOM.EventFilters
    Dim objFilter As SAPbouiCOM.EventFilter

#Region "New"
    Public Sub New()
        MyBase.New()
        Try
            _Company = New SAPbobsCOM.Company
            _Utilities = New clsUtilities
            _Collection = New Hashtable(10, 0.5)
            _LookUpCollection = New Hashtable(10, 0.5)
            oSystemForms = New clsSystemForms
            '   _Log = New clsLog_Error

            SetApplication()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Public Properties"

    Public ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return _SBO_Application
        End Get
    End Property

    Public ReadOnly Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
    End Property

    Public ReadOnly Property RemoteCompany() As SAPbobsCOM.Company
        Get
            Return _RemoteCompany
        End Get
    End Property

    Public ReadOnly Property Utilities() As clsUtilities
        Get
            Return _Utilities
        End Get
    End Property

    Public ReadOnly Property Collection() As Hashtable
        Get
            Return _Collection
        End Get
    End Property

    Public ReadOnly Property LookUpCollection() As Hashtable
        Get
            Return _LookUpCollection
        End Get
    End Property

    'Public ReadOnly Property Log() As clsLog_Error
    '    Get
    '        Return _Log
    '    End Get
    'End Property
#Region "Filter"

    Public Sub SetFilter(ByVal Filters As SAPbouiCOM.EventFilters)
        oApplication.SetFilter(Filters)
    End Sub
    Public Sub SetFilter()
        Try
            ''Form Load
            objFilters = New SAPbouiCOM.EventFilters

            'objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            'objFilter.AddEx(frm_Import)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region

#End Region

#Region "Menu Event"

    Private Sub _SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.FormDataEvent
        Select Case BusinessObjectInfo.FormTypeEx
            'Case frm_Delivery
            '    Dim objInvoice As clsSalesOrder
            '    objInvoice = New clsSalesOrder
            '    objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "426"
                Dim oJe As SAPbobsCOM.Payments
                If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) And BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    oJe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                    If oJe.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        Dim intJe As Integer = oJe.DocEntry
                        Dim intJENUmber As Integer
                        Dim strBnkAccount As String
                        Dim oRec As SAPbobsCOM.Recordset
                        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                            oRec.DoQuery("Select * from OVPM where Trsfrsum>0 and DocEntry=" & intJe)
                            If oRec.RecordCount > 0 Then
                                intJENUmber = oRec.Fields.Item("TransId").Value
                                strBnkAccount = oRec.Fields.Item("PBnkAccnt").Value
                                oRec.DoQuery("Update OJDT set Ref2='" & strBnkAccount & "' where Transid=" & intJENUmber)
                                oRec.DoQuery("Update JDT1 set Ref2='" & strBnkAccount & "' where Transid=" & intJENUmber)
                            End If
                            oApplication.Utilities.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                            If oApplication.Utilities.SendChecktoSystem(oJe.DocEntry) Then
                                ' BubbleEvent = False
                                '  Exit Sub
                            End If
                        Else
                            oApplication.Utilities.FormDataEvent_Update(BusinessObjectInfo, BubbleEvent)
                        End If
                        
                    End If
                End If

            Case "804"
                Dim oJe As SAPbobsCOM.JournalEntries
                If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) And BusinessObjectInfo.BeforeAction = True And BusinessObjectInfo.ActionSuccess = False Then
                    'If oJe.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                    '    Dim oForm As SAPbouiCOM.Form
                    '    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    '    If oApplication.Utilities.ValidateGLAccunt(oApplication.Utilities.getEdittextvalue(oForm, "13"), "GL", oForm) = False Then
                    '        BubbleEvent = False
                    '        Exit Sub
                    '    End If

                    'End If
                    oApplication.Utilities.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                End If

            Case "141", "181", "65301", "392", "1470000009", "1470000015", "1470000012", "1470000013", "1470000037", "229"
                oApplication.Utilities.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_OJAT
                Dim objJVApTemp As clsJVApprovalTemplate
                objJVApTemp = New clsJVApprovalTemplate
                objJVApTemp.FormDataEvent(BusinessObjectInfo, BubbleEvent)
        End Select
        '  End If
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.MenuEvent
        Try

            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case mnu_GFCSetup
                        oMenuObject = New clsGFCSetup
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case "Z_Att"
                        oMenuObject = New clsAttachment
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_IFXPosting
                        oMenuObject = New clsIFXPosting
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_IFXSetup
                        oMenuObject = New clsIFXSetup
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ADD, mnu_FIND, mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD_ROW, mnu_DELETE_ROW
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If
                    Case mnu_JVApproval
                        oMenuObject = New clsJVApproved
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_OJAT
                        oMenuObject = New clsJVApprovalTemplate
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                End Select

            Else
                Select Case pVal.MenuUID
                    'Case mnu_CLOSE, mnu_ADD_ROW, mnu_DELETE_ROW
                    '    If _Collection.ContainsKey(_FormUID) Then
                    '        oMenuObject = _Collection.Item(_FormUID)
                    '        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    '    End If
                End Select

            End If

        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oMenuObject = Nothing
        End Try
    End Sub
#End Region

#Region "Item Event"
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.ItemEvent
        Try
            _FormUID = FormUID

            If pVal.Before_Action = True Then
                Select Case pVal.FormTypeEx
                    Case frm_GFCSetup
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsGFCSetup
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_IFXPosting
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsIFXPosting
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_JournalEntry
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsJournal
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_JournalVoucher
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsJournal
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If


                    Case frm_Attachment
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsAttachment
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_IFXSetup
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsIFXSetup
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case "65052"
                        If pVal.ItemUID = "1" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim oForm As SAPbouiCOM.Form
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If oApplication.Utilities.ValidateBPAccunt(oForm) = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If

                    Case "60100"
                        If pVal.ItemUID = "2" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim oForm As SAPbouiCOM.Form
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                If oApplication.Utilities.ValidateBPAccunt_EmployeeMaster(oForm) = False Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        End If
                    Case "-392", "-393"
                        Dim oForm As SAPbouiCOM.Form
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        If (pVal.ItemUID = "U_Export" Or pVal.ItemUID = "U_Z_IFXReply" Or pVal.ItemUID = "U_FAproval" Or pVal.ItemUID = "U_SAproval" Or pVal.ItemUID = "U_TAproval" Or pVal.ItemUID = "U_FoAproval" Or pVal.ItemUID = "U_FApprTime" Or pVal.ItemUID = "U_SApprTime" Or pVal.ItemUID = "U_TApprTime" Or pVal.ItemUID = "U_FoApprTime" Or pVal.ItemUID = "U_FAppRmks" Or pVal.ItemUID = "U_SAppRmks" Or pVal.ItemUID = "U_TAppRmks" Or pVal.ItemUID = "U_FoAppRmks" Or pVal.ItemUID = "U_FApprover" Or pVal.ItemUID = "U_SApprover" Or pVal.ItemUID = "U_TApprover" Or pVal.ItemUID = "U_FoApprover") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN And pVal.CharPressed <> 9 Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If (pVal.ItemUID = "U_Export" Or pVal.ItemUID = "U_Z_IFXReply" Or pVal.ItemUID = "U_FAproval" Or pVal.ItemUID = "U_SAproval" Or pVal.ItemUID = "U_TAproval" Or pVal.ItemUID = "U_FoAproval" Or pVal.ItemUID = "U_FApprTime" Or pVal.ItemUID = "U_SApprTime" Or pVal.ItemUID = "U_TApprTime" Or pVal.ItemUID = "U_FoApprTime" Or pVal.ItemUID = "U_FAppRmks" Or pVal.ItemUID = "U_SAppRmks" Or pVal.ItemUID = "U_TAppRmks" Or pVal.ItemUID = "U_FoAppRmks" Or pVal.ItemUID = "U_FApprover" Or pVal.ItemUID = "U_SApprover" Or pVal.ItemUID = "U_TApprover" Or pVal.ItemUID = "U_FoApprover") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED) Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        'Case "392", "393"
                        '    Dim oForm As SAPbouiCOM.Form
                        '    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        '    Dim oEditText As SAPbouiCOM.EditText
                        '    If (pVal.ItemUID = "8" Or pVal.ItemUID = "540002023") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN And pVal.CharPressed <> 9 Then
                        '        BubbleEvent = False
                        '        Exit Sub
                        '    End If
                        '    If (pVal.ItemUID = "8" Or pVal.ItemUID = "540002023") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED) Then
                        '        BubbleEvent = False
                        '        Exit Sub
                        '    End If

                        '    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.Before_Action = False Then
                        '        oApplication.Utilities.AddControls(oform, "_10020", "9", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , , , "Attachment Ref")
                        '        oApplication.Utilities.AddControls(oForm, "_10021", "_10020", SAPbouiCOM.BoFormItemTypes.it_EDIT, "RIGHT", , , , "")
                        '        oEditText = oForm.Items.Item("_10021").Specific
                        '        If pVal.FormTypeEx = frm_JournalEntry Then
                        '            oEditText.DataBind.SetBound(True, "OJDT", "U_Z_AttRef")
                        '        Else
                        '            oEditText.DataBind.SetBound(True, "OBTF", "U_Z_AttRef")
                        '        End If
                        '    End If

                        '    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        '        If pVal.ItemUID = "76" And (pVal.ColUID = "11" Or pVal.ColUID = "12" Or pVal.ColUID = "2001" Or pVal.ColUID = "2006" Or pVal.ColUID = "2003") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN And pVal.CharPressed <> 9 Then
                        '            BubbleEvent = False
                        '            Exit Sub
                        '        End If
                        '        If pVal.ItemUID = "76" And (pVal.ColUID = "11" Or pVal.ColUID = "12" Or pVal.ColUID = "2001" Or pVal.ColUID = "2006" Or pVal.ColUID = "2003") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED) Then
                        '            BubbleEvent = False
                        '            Exit Sub
                        '        End If
                        '    End If
                    Case "196"
                        If pVal.ItemUID = "1" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim oForm As SAPbouiCOM.Form
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If oApplication.Utilities.getCheckNumber(oForm) = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    Case "804"
                        If pVal.ItemUID = "1" And pVal.EventType = pVal.EventType.et_ITEM_PRESSED Then
                            If 1 = 1 Then ' oJe.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                Dim oForm As SAPbouiCOM.Form
                                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    If oApplication.Utilities.ValidateGLAccunt(oApplication.Utilities.getEdittextvalue(oForm, "13"), "GL", oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                            End If

                        End If

                    Case "229"
                        If pVal.ItemUID = "4" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim oForm As SAPbouiCOM.Form
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim sLogPath As String = oApplication.Utilities.getApplicationPath() & "\Log\log_JVPostings.txt"
                                If File.Exists(sLogPath) Then
                                    File.Delete(sLogPath)
                                End If
                                oApplication.Utilities.WriteErrorHeader(sLogPath, "Journal Voucher Posting Validation Started")

                                Dim strStatus As String = String.Empty
                                If Not oApplication.Utilities.ValidateJVPostings(oForm, sLogPath) Then
                                    Dim intRet As Integer
                                    intRet = oApplication.SBO_Application.MessageBox("Some of the Journal Vouchers Required Approval . Click OK to view the log file", 1, "OK", "")
                                    If intRet = 1 Then
                                        oApplication.Utilities.OpenFile(sLogPath)
                                    End If
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        ElseIf pVal.ItemUID = "5" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            Dim oForm As SAPbouiCOM.Form
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                Dim strApprover As String = String.Empty
                                If Not oApplication.Utilities.ValidateJVDocumentStatus(oForm, strApprover) Then
                                    BubbleEvent = False
                                    Utilities.Message("Journal Voucher Already Approved by " & strApprover & " ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If
                            End If
                        End If
                    Case frm_JVApproval
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsJVApproved
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case "-393"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    Case "393"
                        Dim oForm As SAPbouiCOM.Form
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        If pVal.ItemUID = "76" And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED) Then
                            If oForm.DataSources.DBDataSources.Item(0).GetValue("U_Approval", 0) <> "O" Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    Case frm_OJAT
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsJVApprovalTemplate
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                End Select
            End If

            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                Select Case pVal.FormTypeEx
                    'Case frm_FuturaSetup
                    '    If Not _Collection.ContainsKey(FormUID) Then
                    '        oItemObject = New clsAcctMapping
                    '        oItemObject.FrmUID = FormUID
                    '        _Collection.Add(FormUID, oItemObject)
                    '    End If
                End Select
            End If
            'If pVal.Before_Action = False Then
            '    If pVal.FormTypeEx = frm_JournalEntry Or pVal.FormTypeEx = frm_JournalVoucher Then
            '        Dim oForm As SAPbouiCOM.Form
            '        oForm = oApplication.SBO_Application.Forms.ActiveForm()
            '        Dim oEditText As SAPbouiCOM.EditText
            '        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.Before_Action = False Then
            '            oApplication.Utilities.AddControls(oForm, "stRef", "9", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , , , "Attachment Ref")
            '            oApplication.Utilities.AddControls(oForm, "edRef", "stRef", SAPbouiCOM.BoFormItemTypes.it_EDIT, "RIGHT", , , , "")
            '            oEditText = oForm.Items.Item("edRef").Specific
            '            If pVal.FormTypeEx = frm_JournalEntry Then
            '                oEditText.DataBind.SetBound(True, "OJDT", "U_Z_AttRef")
            '            Else
            '                oEditText.DataBind.SetBound(True, "OBTF", "U_Z_AttRef")
            '            End If
            '        End If
            '    End If
            'End If

            If _Collection.ContainsKey(FormUID) Then
                oItemObject = _Collection.Item(FormUID)
                If oItemObject.IsLookUpOpen And pVal.BeforeAction = True Then
                    _SBO_Application.Forms.Item(oItemObject.LookUpFormUID).Select()
                    BubbleEvent = False
                    Exit Sub
                End If
                _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.BeforeAction = False Then
                If _LookUpCollection.ContainsKey(FormUID) Then
                    oItemObject = _Collection.Item(_LookUpCollection.Item(FormUID))
                    If Not oItemObject Is Nothing Then
                        oItemObject.IsLookUpOpen = False
                    End If
                    _LookUpCollection.Remove(FormUID)
                End If

                If _Collection.ContainsKey(FormUID) Then
                    _Collection.Item(FormUID) = Nothing
                    _Collection.Remove(FormUID)
                End If

            End If

        Catch ex As Exception
            If ex.Message.Contains("Form - Invalid Form") Then
            Else
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Application Event"
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles _SBO_Application.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    _Utilities.AddRemoveMenus("RemoveMenus.xml")
                    CloseApp()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        End Try
    End Sub
#End Region

#Region "Close Application"
    Private Sub CloseApp()
        Try
            If Not _SBO_Application Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_SBO_Application)
            End If

            If Not _Company Is Nothing Then
                If _Company.Connected Then
                    _Company.Disconnect()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company)
            End If

            _Utilities = Nothing
            _Collection = Nothing
            _LookUpCollection = Nothing

            ThreadClose.Sleep(10)
            System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            Throw ex
        Finally
            oApplication = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Set Application"
    Private Sub SetApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                SboGuiApi = New SAPbouiCOM.SboGuiApi
                SboGuiApi.Connect(sConnectionString)
                _SBO_Application = SboGuiApi.GetApplication()
            Else
                Throw New Exception("Connection string missing.")
            End If

        Catch ex As Exception
            Throw ex
        Finally
            SboGuiApi = Nothing
        End Try
    End Sub
#End Region

#Region "Finalize"
    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
            '            CloseApp()

            oMenuObject = Nothing
            oItemObject = Nothing
            oSystemForms = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Addon Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

    Private Sub _SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.RightClickEvent
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            'If eventInfo.FormUID = "RightClk" Then
            If oForm.TypeEx = frm_JournalEntry Then
                oMenuObject = New clsAttachment
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            End If
            If oForm.TypeEx = frm_JournalVoucher Then
                oMenuObject = New clsAttachment
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub
End Class
