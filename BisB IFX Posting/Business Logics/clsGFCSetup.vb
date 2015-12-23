Public Class clsGFCSetup
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


    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_GFCSetup) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_GFCSetup, frm_GFCSetup)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        Try
            oForm.Freeze(True)
            Databind(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub Databind(ByVal aform As SAPbouiCOM.Form)
        Dim oTempRec, oTemp1, otemp2, otemp3, otemp4 As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1.DoQuery("Select *  from  [@Z_SQLDETAIL]")
        If oTemp1.RecordCount > 0 Then
            oApplication.Utilities.setEdittextvalue(aform, "4", oTemp1.Fields.Item("U_Z_SERVER").Value)
            oApplication.Utilities.setEdittextvalue(aform, "6", oTemp1.Fields.Item("U_Z_DBNAME").Value)
            oApplication.Utilities.setEdittextvalue(aform, "10", oTemp1.Fields.Item("U_Z_UID").Value)

            Dim fields() As String
            Dim strLicenseText As String = oTemp1.Fields.Item("U_Z_PWD").Value
            Dim strDecryptText As String = oApplication.Utilities.Decrypt(strLicenseText, oApplication.Utilities.key)
            If strDecryptText.Contains("$") Then
                fields = strDecryptText.Split(vbTab)
            Else
                fields = strDecryptText.Split("$")
            End If
            Dim strPwd As String
            If fields.Length > 0 Then
                strPwd = fields(0)
            Else
                strPwd = ""
            End If
            oApplication.Utilities.setEdittextvalue(aform, "8", strPwd) ' oTemp1.Fields.Item("U_Z_PWD").Value)


        End If
    End Sub
    Private Function AddtoPayroll(ByVal aForm As SAPbouiCOM.Form) As String
        Dim oUserTable, oUserTable1, ousertable2, ousertable3 As SAPbobsCOM.UserTable
        Dim strCode, strECode, strEname, strGLacc, strRefCode, strPayrollRefNo, strempID As String
        Dim oTempRec, oTemp1, otemp2, otemp3, otemp4 As SAPbobsCOM.Recordset
        oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oUserTable = oApplication.Company.UserTables.Item("Z_SQLDETAIL")
        oTemp1.DoQuery("Update [@Z_SQLDETAIL] set Name= Name +'_XY'")

        strCode = oApplication.Utilities.getMaxCode("@Z_SQLDETAIL", "Code")
        oUserTable.Code = strCode
        oUserTable.Name = strCode & "N"
        oUserTable.UserFields.Fields.Item("U_Z_SERVER").Value = oApplication.Utilities.getEdittextvalue(aForm, "4")
        oUserTable.UserFields.Fields.Item("U_Z_DBNAME").Value = oApplication.Utilities.getEdittextvalue(aForm, "6")
        oUserTable.UserFields.Fields.Item("U_Z_UID").Value = oApplication.Utilities.getEdittextvalue(aForm, "10")

        Dim strEncryptText As String = oApplication.Utilities.Encrypt(oApplication.Utilities.getEdittextvalue(aForm, "8"), oApplication.Utilities.key)
        oUserTable.UserFields.Fields.Item("U_Z_PWD").Value = strEncryptText ' oApplication.Utilities.getEdittextvalue(aForm, "8")
        If oUserTable.Add <> 0 Then
            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return ""
        Else
            oApplication.Utilities.Message("Operation Completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oTemp1.DoQuery("Delete from  [@Z_SQLDETAIL] where Name Like'%_XY'")
            Databind(aForm)
            Return strCode
        End If
    End Function
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_GFCSetup Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If pVal.ItemUID = "3" Then
                                    If oApplication.Utilities.getEdittextvalue(oForm, "4") = "" Then
                                        oApplication.Utilities.Message("Server is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    End If
                                    If oApplication.Utilities.getEdittextvalue(oForm, "6") = "" Then
                                        oApplication.Utilities.Message("DB Name  is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    End If
                                    If oApplication.Utilities.getEdittextvalue(oForm, "10") = "" Then
                                        oApplication.Utilities.Message("User Code is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    End If
                                    If oApplication.Utilities.getEdittextvalue(oForm, "8") = "" Then
                                        oApplication.Utilities.Message("Password is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    End If
                                    AddtoPayroll(oForm)
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
                Case mnu_GFCSetup
                    LoadForm()
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
