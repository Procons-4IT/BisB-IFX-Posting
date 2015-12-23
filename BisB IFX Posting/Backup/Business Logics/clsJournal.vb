Public Class clsJournal
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
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_JournalEntry Or pVal.FormTypeEx = frm_JournalVoucher Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                Dim oForm As SAPbouiCOM.Form
                                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                Dim oEditText As SAPbouiCOM.EditText
                                If (pVal.ItemUID = "8" Or pVal.ItemUID = "540002023") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN And pVal.CharPressed <> 9 Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "8" Or pVal.ItemUID = "540002023") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED) Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    If pVal.ItemUID = "76" And (pVal.ColUID = "11" Or pVal.ColUID = "12" Or pVal.ColUID = "2001" Or pVal.ColUID = "2006" Or pVal.ColUID = "2003") And pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN And pVal.CharPressed <> 9 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                    If pVal.ItemUID = "76" And (pVal.ColUID = "11" Or pVal.ColUID = "12" Or pVal.ColUID = "2001" Or pVal.ColUID = "2006" Or pVal.ColUID = "2003") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED) Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oApplication.Utilities.AddControls(oForm, "_10020", "9", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , , , "Attachment Ref")
                                oApplication.Utilities.AddControls(oForm, "_10021", "_10020", SAPbouiCOM.BoFormItemTypes.it_EDIT, "RIGHT", , , , "")
                                oEditText = oForm.Items.Item("_10021").Specific
                                If pVal.FormTypeEx = frm_JournalEntry Then
                                    oEditText.DataBind.SetBound(True, "OJDT", "U_Z_AttRef")
                                Else
                                    oEditText.DataBind.SetBound(True, "OBTF", "U_Z_AttRef")
                                End If
                                Dim oItem As SAPbouiCOM.Item
                                oItem = oForm.Items.Item("_10020")
                                oItem.LinkTo = "_10021"
                                oItem = oForm.Items.Item("_10021")
                                oItem.Enabled = False
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
                Case "12"
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
