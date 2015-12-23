Imports System.IO
Imports System.Diagnostics.Process
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports system.Net.CredentialCache
Imports BisBIntegration.WebReference.UBSWebservice
Imports BisBIntegration.WebReference

Imports System.Collections.Specialized
Imports System.Security.Cryptography
Imports System.Text
Imports System.Management

Public Class clsUtilities
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private strThousSep As String = ","
    Private strDecSep As String = "."
    Private intQtyDec As Integer = 3
    Private FormNum As Integer
    Dim ConnectionString As String
    Dim cmd As SqlCommand
    Dim ds As DataSet = New DataSet()
    Dim ds1 As DataSet = New DataSet()
    Dim ds2 As DataSet = New DataSet()
    Dim ds3 As DataSet = New DataSet()
    Dim da As SqlDataAdapter
    Dim oRecordSet As SAPbobsCOM.Recordset
    Public key As String = "!@#$%^*()"

    Public Sub New()
        MyBase.New()
        FormNum = 1
    End Sub


    Public Function Encrypt(ByVal strText As String, ByVal strEncrKey _
         As String) As String
        Dim byKey() As Byte = {}
        Dim IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}
        Try
            byKey = System.Text.Encoding.UTF8.GetBytes(Strings.Left(strEncrKey, 8))
            Dim des As New DESCryptoServiceProvider()
            Dim inputByteArray() As Byte = Encoding.UTF8.GetBytes(strText)
            Dim ms As New MemoryStream()
            Dim cs As New CryptoStream(ms, des.CreateEncryptor(byKey, IV), CryptoStreamMode.Write)
            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            Return Convert.ToBase64String(ms.ToArray())
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function Decrypt(ByVal strText As String, ByVal sDecrKey _
               As String) As String
        Dim byKey() As Byte = {}
        Dim IV() As Byte = {&H12, &H34, &H56, &H78, &H90, &HAB, &HCD, &HEF}
        Dim inputByteArray(strText.Length) As Byte
        Try
            byKey = System.Text.Encoding.UTF8.GetBytes(Strings.Left(sDecrKey, 8))
            Dim des As New DESCryptoServiceProvider()
            inputByteArray = Convert.FromBase64String(strText)
            Dim ms As New MemoryStream()
            Dim cs As New CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write)
            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()
            Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8
            Return encoding.GetString(ms.ToArray())
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

    Public Function getLoginPassword(ByVal strLicenseText As String) As String
        Dim fields() As String
        ' Dim strLicenseText As String = oTemp1.Fields.Item("U_Z_PWD").Value
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
        Return strPwd
    End Function
    Public Function createHRMainAuthorization() As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim mUserPermission As SAPbobsCOM.UserPermissionTree
        mUserPermission = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)
        '//Mandatory field, which is the key of the object.
        '//The partner namespace must be included as a prefix followed by _
        mUserPermission.PermissionID = "BiSB"
        '//The Name value that will be displayed in the General Authorization Tree
        mUserPermission.Name = "BisB Integration"
        '//The permission that this object can get
        mUserPermission.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone
        '//In case the level is one, there Is no need to set the FatherID parameter.
        '   mUserPermission.Levels = 1
        RetVal = mUserPermission.Add
        If RetVal = 0 Or -2035 Then
            Return True
        Else
            MsgBox(oApplication.Company.GetLastErrorDescription)
            Return False
        End If


    End Function

    Public Function addChildAuthorization(ByVal aChildID As String, ByVal aChildiDName As String, ByVal aorder As Integer, ByVal aFormType As String, ByVal aParentID As String, ByVal Permission As SAPbobsCOM.BoUPTOptions) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim mUserPermission As SAPbobsCOM.UserPermissionTree
        mUserPermission = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)

        mUserPermission.PermissionID = aChildID
        mUserPermission.Name = aChildiDName
        mUserPermission.Options = Permission ' SAPbobsCOM.BoUPTOptions.bou_FullReadNone

        '//For level 2 and up you must set the object's father unique ID
        'mUserPermission.Level
        mUserPermission.ParentID = aParentID
        mUserPermission.UserPermissionForms.DisplayOrder = aorder
        '//this object manages forms
        ' If aFormType <> "" Then
        mUserPermission.UserPermissionForms.FormType = aFormType
        ' End If

        RetVal = mUserPermission.Add
        If RetVal = 0 Or RetVal = -2035 Then
            Return True
        Else
            MsgBox(oApplication.Company.GetLastErrorDescription)
            Return False
        End If


    End Function

    Public Sub AuthorizationCreation()

        addChildAuthorization("IFX", "IFXSetup", 2, "", "BiSB", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        'Setup
        addChildAuthorization("IFXTrans", "IFX Transactions", 3, "", "IFX", SAPbobsCOM.BoUPTOptions.bou_FullNone)

        addChildAuthorization("IFXSetup", "IFXSetup", 4, frm_IFXSetup, "IFXTrans", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("IFXSetup", "GFCSetup", 4, frm_GFCSetup, "IFXTrans", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("IFXPosting", "IFXPosting", 4, frm_IFXPosting, "IFXTrans", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        'Setup
        addChildAuthorization("IFXApproval", "Journal Voucher Approval", 4, frm_JVApproval, "IFXTrans", SAPbobsCOM.BoUPTOptions.bou_FullNone)

    End Sub

    Public Function validateAuthorization(ByVal aUserId As String, ByVal aFormUID As String) As Boolean
        Dim oAuth As SAPbobsCOM.Recordset
        oAuth = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim struserid As String
        '    Return False
        struserid = oApplication.Company.UserName
        oAuth.DoQuery("select * from UPT1 where FormId='" & aFormUID & "'")
        If (oAuth.RecordCount <= 0) Then
            Return True
        Else
            Dim st As String
            st = oAuth.Fields.Item("PermId").Value
            st = "Select * from USR3 where PermId='" & st & "' and UserLink=" & aUserId
            oAuth.DoQuery(st)
            If oAuth.RecordCount > 0 Then
                If oAuth.Fields.Item("Permission").Value = "N" Then
                    Return False
                End If
                Return True
            Else
                Return True
            End If

        End If

        Return True

    End Function

    Public Function ValidateJVPostings(ByVal oForm As SAPbouiCOM.Form, ByVal strFile As String) As Boolean
        Dim _retVal As Boolean = True
        Dim oMatrix As SAPbouiCOM.Matrix
        ' Dim oUser As SAPbobsCOM.Users
        Dim stString As String
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            oMatrix = oForm.Items.Item("8").Specific
            '  oUser = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)
            '  Dim strJVUserSign As Integer
            '  strJVUserSign = oRecordSet.Fields.Item("UserSign").Value
            For index As Integer = 1 To oMatrix.RowCount
                If oMatrix.IsRowSelected(index) Then
                    stString = "SELECT T1.[BatchNum], T1.[TransId], T1.[BtfStatus], T1.[TransType], T0.[USER_CODE], T0.[U_NAME], T0.[E_Mail], T0.[U_FApprover], T0.[U_SApprover], T0.[U_TApprover] FROM OBTF T1 inner Join OUSR T0 On T1.UserSign = T0.USERID JOIN [@JAT1] T2 On T2.U_OUser = T0.User_Code WHERE (T1.[BatchNum] = '" & oMatrix.Columns.Item("1").Cells.Item(index).Specific.value & "')"
                    oRecordSet.DoQuery(stString)
                    If oRecordSet.RecordCount > 0 Then
                        oRecordSet.DoQuery("Select U_FAproval,U_SAProval,U_TAproval,U_FoAproval,U_Approval From OBTF Where BatchNum = '" & oMatrix.Columns.Item("1").Cells.Item(index).Specific.value & "'")
                        If Not oRecordSet.EoF Then
                            If oRecordSet.Fields.Item("U_FAproval").Value <> "A" And oRecordSet.Fields.Item("U_Approval").Value <> "A" Then
                                WriteErrorlog("First Level Approval not completed / Rejected. Voucher No :" & oMatrix.Columns.Item("1").Cells.Item(index).Specific.value, strFile)
                                _retVal = False
                            ElseIf oRecordSet.Fields.Item("U_SAProval").Value <> "A" And oRecordSet.Fields.Item("U_Approval").Value <> "A" Then
                                WriteErrorlog(" Second Level Approval not completed  / Rejected. Voucher No :" & oMatrix.Columns.Item("1").Cells.Item(index).Specific.value, strFile)
                                _retVal = False
                            ElseIf oRecordSet.Fields.Item("U_TAproval").Value <> "A" And oRecordSet.Fields.Item("U_Approval").Value <> "A" Then
                                WriteErrorlog("Third Level Approval not completed  / Rejected. Voucher No :" & oMatrix.Columns.Item("1").Cells.Item(index).Specific.value, strFile)
                                _retVal = False
                            ElseIf oRecordSet.Fields.Item("U_FoAproval").Value <> "A" And oRecordSet.Fields.Item("U_Approval").Value <> "A" Then
                                WriteErrorlog("Fourth Level Approval not completed  / Rejected. Voucher No :" & oMatrix.Columns.Item("1").Cells.Item(index).Specific.value, strFile)
                                _retVal = False
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

    Public Function ValidateJVDocumentStatus(ByVal oForm As SAPbouiCOM.Form, ByRef strApprover As String) As Boolean
        Dim _retVal As Boolean = True
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oUser As SAPbobsCOM.Users
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            oMatrix = oForm.Items.Item("8").Specific
            oUser = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers)
            If oUser.GetByKey(oApplication.Company.UserSignature) Then
                If oUser.UserFields.Fields.Item("U_FApprover").Value <> "" Then
                    For index As Integer = 1 To oMatrix.RowCount
                        If oMatrix.IsRowSelected(index) Then
                            oRecordSet.DoQuery("Select U_FAproval,U_SAProval,U_TAproval,U_FoAproval From OBTF Where BatchNum = '" & oMatrix.Columns.Item("1").Cells.Item(index).Specific.value & "'")
                            If Not oRecordSet.EoF Then
                                If oRecordSet.Fields.Item("U_FAproval").Value = "A" Then
                                    strApprover = "First Level"
                                    _retVal = False
                                ElseIf oRecordSet.Fields.Item("U_SAProval").Value = "A" Then
                                    strApprover = "Second Level"
                                    _retVal = False
                                ElseIf oRecordSet.Fields.Item("U_TAproval").Value = "A" Then
                                    strApprover = "Third Level"
                                    _retVal = False
                                ElseIf oRecordSet.Fields.Item("U_FoAproval").Value = "A" Then
                                    strApprover = "Fourth Level"
                                    _retVal = False
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

#Region "Open Files"
    Public Sub OpenFile(ByVal strPath As String)
        Try
            If File.Exists(strPath) Then
                Dim process As New System.Diagnostics.Process
                Dim filestart As New System.Diagnostics.ProcessStartInfo(strPath)
                filestart.UseShellExecute = True
                filestart.WindowStyle = ProcessWindowStyle.Normal
                process.StartInfo = filestart
                process.Start()
            End If
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "getStatus"
    Public Function getApprovalStatus(ByVal strStatus As String)
        Dim _retVal As String = String.Empty
        Try
            If strStatus = "O" Then _retVal = "Open"
            If strStatus = "A" Then _retVal = "Approved"
            If strStatus = "R" Then _retVal = "Rejected"
        Catch ex As Exception

        End Try
        Return _retVal
    End Function
#End Region

    Public Sub assignMatrixLineno(ByVal aGrid As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        For intNo As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            aGrid.RowHeaders.SetText(intNo, intNo + 1)
        Next
        aGrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
        aform.Freeze(False)

    End Sub
    'Private Sub FixedAssettest()
    '    Dim oFA As SAPbobsCOM.FixedAssetItemsService

    '    Dim ofas As SAPbobsCOM.AssetDocument
    '    oFA = ofas.AssetDocumentLineCollection

    'End Sub

#Region "FillComboBox"
    Public Sub FillCombobox(ByVal aCombo As SAPbouiCOM.ComboBox, ByVal aQuery As String)
        Dim oRS As SAPbobsCOM.Recordset
        oRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS.DoQuery(aQuery)
        For intRow As Integer = aCombo.ValidValues.Count - 1 To 0 Step -1
            aCombo.ValidValues.Remove(intRow)
        Next
        aCombo.ValidValues.Add("", "")
        For intRow As Integer = 0 To oRS.RecordCount - 1
            Try
                aCombo.ValidValues.Add(oRS.Fields.Item(0).Value, oRS.Fields.Item(1).Value)

            Catch ex As Exception

            End Try
            oRS.MoveNext()
        Next
        aCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
    End Sub
#End Region
#Region "BisB Integration"
    Public Function SqlConnectionString() As String
        Dim sername, dbname, uid, pwd As String
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select * from [@Z_SQLDETAIL]")
        If oRec.RecordCount > 0 Then
            sername = oRec.Fields.Item(2).Value
            dbname = oRec.Fields.Item(3).Value
            uid = oRec.Fields.Item(4).Value
            pwd = oRec.Fields.Item(5).Value
            pwd = getLoginPassword(pwd)
            'ConnectionString = "data source=" & sername & ";Integrated Security=SSPI;database=" & dbname & ";Trusted_Connection=True; User id=" & uid & "; password=" & pwd
            'ConnectionString = "data source=" & sername & ";database=" & dbname & ";Trusted_Connection=True; User id=" & uid & "; password=" & pwd
            ConnectionString = "data source=" & sername & ";database=" & dbname & "; User id=" & uid & "; password=" & pwd
        End If
        Return ConnectionString
    End Function
    Public Function getCheckNumber(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim dblcheck, dblSPReturnValue As String
        Dim userid As String = ""
        Dim con As SqlConnection = New SqlConnection(SqlConnectionString)
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMatrix = aForm.Items.Item("28").Specific
        Dim usercode As String = oApplication.Company.UserSignature.ToString()
        oRec.DoQuery("Select U_GFSID from OUSR where UserId='" & usercode & "'")
        If oRec.RecordCount > 0 Then
            userid = oRec.Fields.Item(0).Value
        End If
        For introw As Integer = 1 To oMatrix.RowCount
            dblcheck = getDocumentQuantity(getMatrixValues(oMatrix, "7", introw))
            If dblcheck > 0 Then
                dblcheck = getDocumentQuantity(getMatrixValues(oMatrix, "5", introw))
                If dblcheck <= 0 Then
                    oApplication.Utilities.Message("Check number is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                con.Open()
                '    Dim st As String = "exec dbo.Stock_Move_Retrieve N'CO', -1, 1, 911,0, " & userid & ", 0"
                Dim st As String = "exec dbo.Stock_Move_Retrieve N'CO', -1, 1, 911, " & userid & ", 0"
                cmd = con.CreateCommand()
                'cmd.CommandText = "Stock_MOVE_Retrieve"
                'cmd.CommandType = CommandType.StoredProcedure
                'cmd.Parameters.AddWithValue("@StockCode", "CO")
                'cmd.Parameters.AddWithValue("@InventoryID ", "-1")
                'cmd.Parameters.AddWithValue("@Quantity", 1)
                'cmd.Parameters.AddWithValue("@UserBranch", 911)
                'cmd.Parameters.AddWithValue("@isVault", 0)
                'cmd.Parameters.AddWithValue("@IsSafatejBankCheques", 0)
                'cmd.Parameters.AddWithValue("@UserNumber ", userid)
                cmd.CommandType = CommandType.Text
                ds.Clear()
                da = New SqlDataAdapter(st, con)
                da.Fill(ds)
                If ds.Tables(0).Rows.Count > 0 Then
                    dblSPReturnValue = ds.Tables(0).Rows(0)(0).ToString()
                    If dblcheck <> dblSPReturnValue Then
                        oApplication.Utilities.Message("Invalid check number, Use this check number :" & dblSPReturnValue, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        con.Close()
                        Return False
                    End If
                Else
                    oApplication.Utilities.Message("No Check numbers are available for the user", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    con.Close()
                    Return False
                End If
            End If
        Next
        con.Close()
        Return True
    End Function
    Public Function SendChecktoSystem(ByVal aDocEntry As Integer) As Boolean
        Try
            Dim oTest As SAPbobsCOM.Recordset
            Dim userid As String = ""
            Dim con As SqlConnection = New SqlConnection(SqlConnectionString)
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim usercode As String = oApplication.Company.UserName
            oRec.DoQuery("Select U_GFSID from OUSR where USER_CODE='" & usercode & "'")
            If oRec.RecordCount > 0 Then
                userid = oRec.Fields.Item(0).Value
            End If
            Dim strChequeNo, strCurrency As String
            Dim dblCheckSum As String
            Dim dtDate As Date
            Dim OutGoingQuery As String = "SELECT *  FROM OVPM T0  INNER JOIN VPM1 T1 ON T0.DocEntry = T1.DocNum where DocEntry=" & aDocEntry
            oTest.DoQuery(OutGoingQuery)
            If oTest.RecordCount > 0 Then
                strChequeNo = oTest.Fields.Item("CheckNum").Value
                strCurrency = oTest.Fields.Item("Currency").Value
                dblCheckSum = oTest.Fields.Item("CheckSum").Value
                dtDate = oTest.Fields.Item("DueDate").Value
                con.Open()
                Dim st1 As String = "exec dbo.get_New_RefNo 911, N'" & dtDate.Year.ToString("00") & "', N'CO', ''"
                cmd = con.CreateCommand()
                cmd.CommandText = "get_New_RefNo"
                cmd.CommandType = CommandType.Text
                da = New SqlDataAdapter(st1, con)
                da.Fill(ds1)
                Dim dblSPReturnValue As String
                If ds1.Tables(0).Rows.Count > 0 Then
                    dblSPReturnValue = ds1.Tables(0).Rows(0)(0).ToString()
                    cmd = con.CreateCommand()
                    cmd.CommandText = "Insert_Stock_Move"
                    cmd.CommandType = CommandType.Text
                    Dim stQuery As String
                    stQuery = "exec dbo.Insert_Stock_Move N'CO', -1, " & strChequeNo & ", 1, N'" & dblSPReturnValue & "', 911, " & userid & ", 0, 5, 6, ''"
                    con.Close()
                    con.Open()
                    da = New SqlDataAdapter(stQuery, con)
                    Try
                        da.Fill(ds2)
                    Catch ex As Exception
                        Return False
                    End Try
                    If 1 = 1 Then ' ds2.Tables(0).Rows.Count > 0 Then
                        cmd = con.CreateCommand()
                        cmd.CommandText = "InsertIntoCashierOrderCheques"
                        cmd.CommandType = CommandType.Text
                        Dim strCommand, strCardName, strRefNo, strComments, strBPAccount, strBranch, strDepartment As String
                        strCardName = oTest.Fields.Item("CardName").Value
                        strComments = oTest.Fields.Item("Comments").Value
                        strBPAccount = oTest.Fields.Item("BPAct").Value
                        oRec.DoQuery("Select * from OACT where AcctCode='" & strBPAccount & "'")
                        strBranch = oRec.Fields.Item("OverCode").Value
                        strDepartment = oRec.Fields.Item("OverCode2").Value
                        If strBranch.Length > 3 Then
                            strBranch = strBranch.Substring(strBranch.Length - 3, 3)
                        Else
                            strBranch = strBranch
                        End If

                        If strDepartment.Length > 3 Then
                            strDepartment = strDepartment.Substring(strDepartment.Length - 3, 3)
                        Else
                            strDepartment = strDepartment
                        End If
                        strBPAccount = "01-01-" & strBranch & "-" & strDepartment & "-" & strBPAccount

                        strCommand = "exec dbo.InsertIntoCashierOrderCheques 1, 1, 911, N'" & strChequeNo & "'," & dblCheckSum & ", N'IssueCOAgainstGL-SAP', N'" & strCurrency & "'," & dblCheckSum & ", N'" & strComments & "', N'" & dtDate.ToString("yyyy/MM/dd") & "', N'" & strCardName & "', '', '', '', '', 0, N'" & strBPAccount & "', N'Issued'"
                        da = New SqlDataAdapter(strCommand, con)
                        da.Fill(ds)
                        If ds.Tables(0).Rows.Count > 0 Then
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            Else
                Return True
            End If
            Return True
            'Call the three stored procedure with parameter

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Public Sub ExportJournalEntries(ByVal aPath As String)
        Dim oRecItem, oRecItemCode, oTemp, oMainRec As SAPbobsCOM.Recordset
        Dim strSQL, strFilename As String
        Dim sValue As String
        Dim sPath, strLogDirectory, strPath, strMessage, strSelectedFolderPath, strExportFilePaty As String
        Dim blnErrorflag As Boolean
        Dim FILedatetiem As String
        strPath = System.Windows.Forms.Application.StartupPath
        strFilename = Now.ToLongDateString
        strPath = aPath
        ' strPath = aFileName
        oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMainRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        strSQL = "Select T1.TransId ,T1.RefDate,Count(*) from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where Isnull(T1.U_Export,'N')='N'  group by T1.TransId,T1.RefDate" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
        oMainRec.DoQuery(strSQL)
        Dim strTransID As String
        Dim dtJEDate As Date
        For intLoop As Integer = 0 To oMainRec.RecordCount - 1
            strTransID = oMainRec.Fields.Item(0).Value
            dtJEDate = oMainRec.Fields.Item(1).Value

            Dim strPhxId As String = ""
            'BP customer Master
            'strSQL = "Select * from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  where Isnull(T1.U_Export,'N')='N'" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"

            strSQL = "Select *,T0.Ref2 'OutGoing' from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where T1.TransId=" & oMainRec.Fields.Item(0).Value & " and  Isnull(T1.U_Export,'N')='N'" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"


            oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItem.DoQuery(strSQL)
            FILedatetiem = Now.ToString("yyyyMMdd_hh_mm_ss")
            Dim transtype As String
            Dim strTRGFileName As String
            Dim dtDateTime As String
            Dim strJNo As String = ""
            dtDateTime = dtJEDate.ToString("yyyy-MM-dd") & " _" & strTransID ' Format(Now.Date, "yyyy-MM-dd") & Now.ToLongTimeString.Replace(":", "")
            FILedatetiem = dtDateTime
            dtDateTime = dtJEDate.ToString("yyyy-MM-dd")
            strTRGFileName = strExportFilePaty & "\JE_" & FILedatetiem & ".trg"
            strFilename = strExportFilePaty & "\JE_" & FILedatetiem & ".xml"
            Dim MyXMLString As String
            Dim myStringWriter As New StringWriter
            Dim writer As New XmlTextWriter(myStringWriter)
            '(strFilename, System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2
            writer.WriteStartElement("IFX")

            writer.WriteStartElement("SignonRq")

            writer.WriteStartElement("RqUID")
            writer.WriteString("6a8b5973-ec37-47e0-855e-4cc020ab7f02")
            writer.WriteEndElement()

            writer.WriteStartElement("ClientDt")
            writer.WriteString(dtJEDate.ToString("yyyy-MM-dd"))
            writer.WriteEndElement()

            writer.WriteStartElement("ClientApp")
            writer.WriteString("SAP")
            writer.WriteEndElement()

            writer.WriteStartElement("OperatorId")
            writer.WriteString("Manager")
            writer.WriteEndElement()

            writer.WriteEndElement()

            'BankSvcRq
            writer.WriteStartElement("BankSvcRq")
            writer.WriteStartElement("FinancialMessageAddRq")
            Dim strFCCurrency, strLocalCurrency, strCurrency, strType, strAccountType As String
            Dim dblAmount As Double
            oTemp.DoQuery("Select MainCurncy from OADM")
            strLocalCurrency = oTemp.Fields.Item(0).Value

            For intRow As Integer = 0 To oRecItem.RecordCount - 1
                strPhxId = oRecItem.Fields.Item("U_PhxId").Value

                If strJNo = "" Then
                    strJNo = oRecItem.Fields.Item("TransId").Value
                Else
                    strJNo = strJNo & "," & oRecItem.Fields.Item("TransId").Value
                End If
                If oRecItem.Fields.Item("Account").Value <> oRecItem.Fields.Item("ShortName").Value Then
                    strAccountType = "GL"
                Else
                    strAccountType = "GL"
                End If
                strFCCurrency = oRecItem.Fields.Item("FCCurrency").Value
                If strFCCurrency <> "" Then
                    strCurrency = strFCCurrency
                    If oRecItem.Fields.Item("FCDebit").Value <> 0 Then
                        dblAmount = oRecItem.Fields.Item("FCDebit").Value
                        strType = "Dr"
                    Else
                        dblAmount = oRecItem.Fields.Item("FCCredit").Value
                        strType = "Cr"
                    End If
                Else
                    strCurrency = strLocalCurrency
                    If oRecItem.Fields.Item("Debit").Value <> 0 Then
                        dblAmount = oRecItem.Fields.Item("Debit").Value
                        strType = "Dr"
                    Else
                        dblAmount = oRecItem.Fields.Item("Credit").Value
                        strType = "Cr"
                    End If
                End If
                writer.WriteStartElement("FinancialEntry")

                Dim straccount, sAPAccount, CostCenter1, CostCenter2 As String

                If oRecItem.Fields.Item("OutGoing").Value = "" Then
                    writer.WriteStartElement("CurCode")
                    writer.WriteString(strCurrency)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Amt")
                    writer.WriteString(dblAmount.ToString)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Description")
                    writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                    writer.WriteEndElement()


                    straccount = "01-01"
                    If oRecItem.Fields.Item("OutGoing").Value = "" Then
                        sAPAccount = oRecItem.Fields.Item("Account").Value
                        CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                        CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                        If CostCenter1.Length > 0 Then
                            If CostCenter1.Length > 3 Then
                                CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                            Else
                                CostCenter1 = CostCenter1
                            End If
                        Else
                            CostCenter1 = ""
                        End If
                        If CostCenter2.Length > 0 Then
                            If CostCenter2.Length > 3 Then
                                CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                            Else
                                CostCenter2 = CostCenter2
                            End If
                        Else
                            CostCenter2 = ""
                        End If
                        straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                    Else
                        straccount = oRecItem.Fields.Item("OutGoing").Value
                    End If


                    writer.WriteStartElement("AcctId")
                    '   writer.WriteString(oRecItem.Fields.Item("Account").Value)
                    writer.WriteString(straccount)
                    writer.WriteEndElement()

                    If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                        writer.WriteStartElement("ApplType")
                        writer.WriteString("SV")
                        writer.WriteEndElement()
                        writer.WriteStartElement("AcctType")
                        writer.WriteString("SAV")
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("AcctType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    End If

                    writer.WriteStartElement("EntryType")
                    writer.WriteString(strType)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EffDt")
                    writer.WriteString(dtDateTime)
                    writer.WriteEndElement()

                    writer.WriteStartElement("ApplType")
                    writer.WriteString(strAccountType)
                    writer.WriteEndElement()
                Else
                    writer.WriteStartElement("CurCode")
                    writer.WriteString(strCurrency)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Amt")
                    writer.WriteString(dblAmount.ToString)
                    writer.WriteEndElement()



                    straccount = "01-01"
                    If oRecItem.Fields.Item("OutGoing").Value = "" Then
                        sAPAccount = oRecItem.Fields.Item("Account").Value
                        CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                        CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                        If CostCenter1.Length > 0 Then
                            If CostCenter1.Length > 3 Then
                                CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                            Else
                                CostCenter1 = CostCenter1
                            End If
                        Else
                            CostCenter1 = ""
                        End If
                        If CostCenter2.Length > 0 Then
                            If CostCenter2.Length > 3 Then
                                CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                            Else
                                CostCenter2 = CostCenter2
                            End If
                        Else
                            CostCenter2 = ""
                        End If
                        straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                    Else
                        straccount = oRecItem.Fields.Item("OutGoing").Value
                    End If


                    writer.WriteStartElement("AcctId")
                    writer.WriteString(straccount)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EntryType")
                    writer.WriteString(strType)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EffDt")
                    writer.WriteString(dtDateTime)
                    writer.WriteEndElement()

                    If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                        If strType = "Dr" Then
                            writer.WriteStartElement("ApplType")
                            writer.WriteString("SV")
                            writer.WriteEndElement()
                            writer.WriteStartElement("AcctType")
                            writer.WriteString("SAV")
                            writer.WriteEndElement()
                        Else
                            writer.WriteStartElement("ApplType")
                            writer.WriteString("CK")
                            writer.WriteEndElement()
                            writer.WriteStartElement("AcctType")
                            writer.WriteString("CUR")
                            writer.WriteEndElement()
                        End If

                    Else
                        writer.WriteStartElement("AcctType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    End If


                    writer.WriteStartElement("Description")
                    writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                    writer.WriteEndElement()

                End If


                writer.WriteEndElement()
                oRecItem.MoveNext()
            Next
            ' writer.WriteString(dtDateTime)


            writer.WriteStartElement("BankInfo")
            writer.WriteStartElement("BranchId")
            writer.WriteString("911")
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("EmployeeIdent")
            writer.WriteStartElement("EmployeeIdentlNum")
            writer.WriteString(strPhxId)
            writer.WriteEndElement()
            writer.WriteStartElement("SuperEmployeeIdentlNum")
            writer.WriteString("0")
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("RqUID")
            writer.WriteString("85fb650d-336e-4c1d-bea3-6d9551b9fb08")
            writer.WriteEndElement()

            writer.WriteEndElement()

            writer.WriteStartElement("RqUID")
            writer.WriteString("85fb650d-336e-4c1d-bea3-6d9551b9fb08")
            writer.WriteEndElement()


            writer.WriteEndElement()
            writer.WriteEndElement()
            writer.Flush()
            MyXMLString = myStringWriter.ToString()


            myStringWriter.Close()
            writer.Close()
            ' SendXMLtoIFX(MyXMLString)

            Dim doc As New XmlDocument
            doc.LoadXml(MyXMLString)
            Dim locX As String = (doc.SelectSingleNode("IFX/BankSvcRq/FinancialMessageAddRq/FinancialEntry/CurCode").InnerText).ToString
            Dim locY As Integer = (doc.SelectSingleNode("IFX/BankSvcRq/FinancialMessageAddRq/FinancialEntry/Amt").InnerText)
            If strJNo <> "" Then
                oRecItem.DoQuery("Update OJDT set U_Export='Y' where TransId in (" & strJNo & ")")
            End If
            strMessage = "Export Jounral Entry  Compleated : " & strFilename
            WriteErrorlog(strMessage, strPath)
            oMainRec.MoveNext()

        Next

        'If File.Exists(strTRGFileName) Then
        '    File.Delete(strTRGFileName)
        'End If
    End Sub

    Public Sub canceldocument(ByVal aDocEntry As Integer)
        Dim oDoc1, oDoc As SAPbobsCOM.Documents

        oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        If oDoc.GetByKey(aDocEntry) Then
            oDoc1 = oDoc.CreateCancellationDocument()
            If oDoc1.Add() <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        End If
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                Dim oForm As SAPbouiCOM.Form
                Dim strDocNumber, strSQL As String
                Dim oTest As SAPbobsCOM.Recordset
                Dim oDoc As SAPbobsCOM.Documents
                '  oForm = oApplication.SBO_Application.Forms.ActiveForm()
                Select Case BusinessObjectInfo.FormTypeEx
                    Case "141" 'AP Invocie
                        oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                        If oDoc.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                            Dim intDocEntry As Integer = oDoc.DocEntry
                            If ExportJournalEntries_Testing(oDoc.TransNum) = False Then
                                'canceldocument(intDocEntry)
                                Dim oDoc1 As SAPbobsCOM.Documents
                                oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                If oDoc.GetByKey(intDocEntry) Then
                                    oDoc1 = oDoc.CreateCancellationDocument()
                                    If oDoc1.Add() <> 0 Then
                                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                End If
                            End If
                        End If
                    Case "181" 'AP CR
                        oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                        If oDoc.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                            Dim intDocEntry As Integer = oDoc.DocEntry
                            If ExportJournalEntries_Testing(oDoc.TransNum) = False Then
                                Dim oDoc1 As SAPbobsCOM.Documents
                                oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                                If oDoc.GetByKey(intDocEntry) Then
                                    oDoc1 = oDoc.CreateCancellationDocument()
                                    If oDoc1.Add() <> 0 Then
                                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                End If
                            End If

                        End If
                    Case "65301" 'AP Downpayment Invoice
                        oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments)
                        If oDoc.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                            '  ExportJournalEntries_Testing(oDoc.TransNum)
                            Dim intDocEntry As Integer = oDoc.DocEntry
                            If ExportJournalEntries_Testing(oDoc.TransNum) = False Then
                                Dim oDoc1 As SAPbobsCOM.Documents
                                oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments)
                                If oDoc.GetByKey(intDocEntry) Then
                                    oDoc1 = oDoc.CreateCancellationDocument()
                                    If oDoc1.Add() <> 0 Then
                                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End If
                                End If
                            End If
                        End If
                    Case "392", "229" 'Journal Entry, Post Journal Voucher
                        Dim oJE As SAPbobsCOM.JournalEntries
                        oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                        If BusinessObjectInfo.FormTypeEx = "229" Then 'Journal Voucher
                            Dim xmlInput As String = BusinessObjectInfo.ObjectKey
                            Dim xmlObj As New Xml.XmlDocument
                            xmlObj.LoadXml(xmlInput)
                            Dim stAPM_Read As String = ""
                            Dim xmlUnits As XmlElement = xmlObj.GetElementsByTagName("BatchNum")(0)
                            stAPM_Read = xmlUnits.InnerText
                            Dim oRec As SAPbobsCOM.Recordset
                            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRec.DoQuery("Select * from OJDT where BatchNum= " & stAPM_Read)
                            If oRec.RecordCount > 0 Then
                                If oJE.GetByKey(oRec.Fields.Item("TransId").Value) Then
                                    If ExportJournalEntries_Testing(oJE.Number) = False Then
                                        Dim intDocEntry As Integer = oJE.Number
                                        Dim oDoc1 As SAPbobsCOM.JournalEntries
                                        Dim oDoc2 As SAPbobsCOM.JournalEntries
                                        oDoc2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                                        oDoc1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                                        If oDoc2.GetByKey(intDocEntry) Then
                                            If oDoc2.Memo.Contains("Reversal") = True Then
                                                Exit Sub
                                            End If
                                            If oDoc2.Cancel <> 0 Then
                                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            End If

                                        End If
                                    End If
                                End If
                            End If
                        Else
                            If oJE.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then 'Journal Entry
                                Dim oRec As SAPbobsCOM.Recordset
                                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRec.DoQuery("Update OJDT set U_Export='N' where TransId=" & oJE.Number)
                                If ExportJournalEntries_Testing(oJE.Number) = False Then
                                    Dim intDocEntry As Integer = oJE.Number
                                    Dim oDoc1 As SAPbobsCOM.JournalEntries
                                    Dim oDoc2 As SAPbobsCOM.JournalEntries
                                    oDoc2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                                    oDoc1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                                    If oDoc2.GetByKey(intDocEntry) Then
                                        If oDoc2.Memo.Contains("Reversal") = True Then
                                            Exit Sub
                                        End If
                                        If oDoc2.Cancel <> 0 Then
                                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If

                                    End If
                                End If
                            End If

                        End If
                    Case "426" 'Out going Payment
                            Dim oJE As SAPbobsCOM.Payments
                            oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                            If oJE.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                If oJE.Cancelled = SAPbobsCOM.BoYesNoEnum.tNO Then
                                    oTest.DoQuery("Select * from OVPM where DocEntry=" & oJE.DocEntry)
                                    If ExportJournalEntries_Testing(oTest.Fields.Item("TransId").Value) = False Then
                                        If oJE.Cancel <> 0 Then
                                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If
                                Else
                                    oTest.DoQuery("Select * from OVPM where DocEntry=" & oJE.DocEntry)
                                    If ExportJournalEntries_Testing(oTest.Fields.Item("TransId").Value) = False Then
                                    End If
                                End If
                            End If
                    Case "1470000009" 'Captilization OACQ 
                            If 1 = 1 Then ' oJE.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oTest.DoQuery("Select * from OACQ where DocNum=" & oAssetTransactionNumber)
                                ExportJournalEntries_Testing(oTest.Fields.Item("TransId").Value)
                            End If
                    Case "1470000015" 'Captilization CR OACD
                            If 1 = 1 Then ' oJE.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oTest.DoQuery("Select * from OACD where DocNum=" & oAssetTransactionNumber)
                                ExportJournalEntries_Testing(oTest.Fields.Item("TransId").Value)
                            End If
                    Case "1470000012" 'Manual Depriciation OMDP
                            If 1 = 1 Then ' oJE.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oTest.DoQuery("Select * from OMDP where DocNum=" & oAssetTransactionNumber)
                                ExportJournalEntries_Testing(oTest.Fields.Item("TransId").Value)
                            End If
                    Case "1470000013" 'Transfer 
                            If 1 = 1 Then ' oJE.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oTest.DoQuery("Select * from OFTR where DocNum=" & oAssetTransactionNumber)
                                ExportJournalEntries_Testing(oTest.Fields.Item("TransId").Value)
                            End If
                    Case "1470000037" 'Transfer 
                            If 1 = 1 Then ' oJE.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                                oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oTest.DoQuery("Select Max(DocEntry) from ODRN ")
                                oAssetTransactionNumber = oTest.Fields.Item(0).Value
                                oTest.DoQuery("Select * from DRN1 where DocEntry=" & oAssetTransactionNumber)
                                For intRow As Integer = 0 To oTest.RecordCount - 1
                                    ExportJournalEntries_Testing(oTest.Fields.Item("TransId").Value)
                                    oTest.MoveNext()
                                Next
                            End If
                End Select
            Else
                Dim oForm As SAPbouiCOM.Form
                Dim strDocNumber, strSQL As String
                Dim oTest As SAPbobsCOM.Recordset
                Dim oDoc As SAPbobsCOM.Documents
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                Select Case BusinessObjectInfo.FormTypeEx
                    Case "1470000009" 'Captilization OACQ 
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                            oAssetTransactionNumber = getEdittextvalue(oForm, "1470000015")
                        End If
                    Case "1470000015" 'Captilization CR OACD
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                            oAssetTransactionNumber = getEdittextvalue(oForm, "1470000015")
                        End If
                    Case "1470000012" 'Manual Depriciation OMDP
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                            oAssetTransactionNumber = getEdittextvalue(oForm, "1470000015")
                        End If
                    Case "1470000013" 'Transfer 
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                            oAssetTransactionNumber = getEdittextvalue(oForm, "1470000015")
                        End If
                End Select
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub FormDataEvent_Update(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                Dim oForm As SAPbouiCOM.Form
                Dim strDocNumber, strSQL As String
                Dim oTest As SAPbobsCOM.Recordset
                Dim oDoc As SAPbobsCOM.Documents
                '  oForm = oApplication.SBO_Application.Forms.ActiveForm()
                Select Case BusinessObjectInfo.FormTypeEx

                    Case "426" 'Out going Payment
                        Dim oJE As SAPbobsCOM.Payments
                        oJE = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                        If oJE.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            If oJE.Cancelled = SAPbobsCOM.BoYesNoEnum.tNO Then
                                oTest.DoQuery("Select * from OVPM where DocEntry=" & oJE.DocEntry)
                                'ExportJournalEntries_Testing(oTest.Fields.Item("TransId").Value)
                            Else
                                Dim strQuery As String
                                strQuery = "Select * from OJDT where BaseRef=" & oJE.DocEntry & " and TransType=46 and Memo Like 'Reverse Entry for Payment No. " & oJE.DocNum & "'"
                                oTest.DoQuery(strQuery)
                                'oTest.DoQuery("Select * from OVPM where DocEntry=" & oJE.DocEntry)
                                ExportJournalEntries_Testing_Reverse(oTest.Fields.Item("TransId").Value)
                            End If

                        End If

                End Select
            Else
                Dim oForm As SAPbouiCOM.Form
                Dim strDocNumber, strSQL As String
                Dim oTest As SAPbobsCOM.Recordset
                Dim oDoc As SAPbobsCOM.Documents
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                Select Case BusinessObjectInfo.FormTypeEx

                End Select
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function ExportJournalEntries_Testing(ByVal aPath As String) As Boolean
        Dim oRecItem, oRecItemCode, oTemp, oMainRec As SAPbobsCOM.Recordset
        Dim strSQL, strFilename, strUserID As String
        Dim sValue As String
        Dim sPath, strLogDirectory, strPath, strMessage, strSelectedFolderPath, strExportFilePaty As String
        Dim blnErrorflag As Boolean
        Dim FILedatetiem As String
        strPath = System.Windows.Forms.Application.StartupPath
        strFilename = Now.ToLongDateString
        strPath = aPath
        ' strPath = aFileName
        oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMainRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'aPath = 1

        'Foreign Currency Transaction Validation
        strSQL = "Select T1.TransId ,T1.RefDate,Count(*) from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where isnull(T0.FCCurrency,'')<>'' and  Isnull(T1.U_Export,'N')='N' and T0.TransId=" & CInt(aPath) & " group by T1.TransId,T1.RefDate" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
        oMainRec.DoQuery(strSQL)
        If oMainRec.RecordCount > 0 Then
            If ExportJournalEntries_Testing_MultiCurrency(aPath) = True Then
                Return True
            Else
                Return False
            End If
        End If

        strSQL = "Select T1.TransId ,T1.RefDate,T1.TransType,Count(*) from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where  Isnull(T1.U_Export,'N')='N' and T0.TransId=" & CInt(aPath) & " group by T1.TransId,T1.RefDate,T1.TransType" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
        oMainRec.DoQuery(strSQL)
        If oMainRec.RecordCount > 0 Then
            If oMainRec.Fields.Item("TransType").Value = "30" Then '
                If ExportJournalEntries_Testing_JournalEntry(aPath) = True Then
                    Return True
                Else
                    Return False
                End If
            End If
        End If

        'Foreign Currency Transaction Validation


        strSQL = "Select T1.TransId ,T1.RefDate,Count(*) from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where Isnull(T1.U_Export,'N')='N' and T0.TransId=" & CInt(aPath) & " group by T1.TransId,T1.RefDate" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
        oMainRec.DoQuery(strSQL)
        Dim strTransID As String
        Dim dtJEDate As Date
        For intLoop As Integer = 0 To oMainRec.RecordCount - 1
            strTransID = oMainRec.Fields.Item(0).Value
            dtJEDate = oMainRec.Fields.Item(1).Value

            Dim strPhxId As String = ""
            'BP customer Master
            'strSQL = "Select * from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  where Isnull(T1.U_Export,'N')='N'" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"

            strSQL = "Select *,T0.Ref2 'OutGoing',isnull(T2.U_PhxId,'911') 'Uid' from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where T1.TransId=" & oMainRec.Fields.Item(0).Value & " and  Isnull(T1.U_Export,'N')='N'" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
            strSQL = "Select *,T0.Ref2 'OutGoing',T1.TransType 'TransType',T0.ShortName 'BPName',isnull(T2.U_PhxId,'911') 'Uid',T0.Ref3Line 'CheckNo' from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where  isnull(T0.VatGroup,'')<>'X1' and Isnull(T1.U_Export,'N')='N' and  T1.TransId=" & oMainRec.Fields.Item(0).Value  'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"


            oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItem.DoQuery(strSQL)
            FILedatetiem = Now.ToString("yyyyMMdd_hh_mm_ss")
            Dim transtype As String
            Dim strTRGFileName As String
            Dim dtDateTime As String
            Dim strJNo As String = ""
            Dim strCheckNo As String = ""
            dtDateTime = dtJEDate.ToString("yyyy-MM-dd") & " _" & strTransID ' Format(Now.Date, "yyyy-MM-dd") & Now.ToLongTimeString.Replace(":", "")
            FILedatetiem = dtDateTime
            dtDateTime = dtJEDate.ToString("yyyy-MM-dd")
            strTRGFileName = strExportFilePaty & "\JE_" & FILedatetiem & ".trg"
            strFilename = strExportFilePaty & "\JE_" & FILedatetiem & ".xml"
            Dim MyXMLString As String
            Dim myStringWriter As New StringWriter
            Dim writer As New XmlTextWriter(myStringWriter)
            '(strFilename, System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2
            writer.WriteStartElement("IFX")

            writer.WriteStartElement("SignonRq")
            writer.WriteStartElement("RqUID")
            writer.WriteString("6a8b5973-ec37-47e0-855e-4cc020ab7f02")
            writer.WriteEndElement()
            writer.WriteStartElement("ClientDt")
            writer.WriteString(dtJEDate.ToString("yyyy-MM-dd"))
            writer.WriteEndElement()
            writer.WriteStartElement("ClientApp")
            writer.WriteString("SAP")
            writer.WriteEndElement()
            writer.WriteStartElement("OperatorId")
            writer.WriteString("Manager")
            writer.WriteEndElement()
            writer.WriteEndElement()

            'BankSvcRq
            writer.WriteStartElement("BankSvcRq")
            writer.WriteStartElement("FinancialMessageAddRq")
            Dim strFCCurrency, strLocalCurrency, strCurrency, strType, strAccountType As String
            Dim dblAmount As Double
            oTemp.DoQuery("Select MainCurncy from OADM")
            strLocalCurrency = oTemp.Fields.Item(0).Value

            For intRow As Integer = 0 To oRecItem.RecordCount - 1

                strPhxId = oRecItem.Fields.Item("U_PhxId").Value
                If strJNo = "" Then
                    strJNo = oRecItem.Fields.Item("TransId").Value
                Else
                    strJNo = strJNo & "," & oRecItem.Fields.Item("TransId").Value
                End If
                If oRecItem.Fields.Item("Account").Value <> oRecItem.Fields.Item("ShortName").Value Then
                    strAccountType = "GL"
                Else
                    strAccountType = "GL"
                End If
                strFCCurrency = oRecItem.Fields.Item("FCCurrency").Value
                If strFCCurrency <> "" Then
                    strCurrency = strFCCurrency
                    If oRecItem.Fields.Item("FCDebit").Value <> 0 Then
                        dblAmount = oRecItem.Fields.Item("FCDebit").Value
                        strType = "Dr"
                    Else
                        dblAmount = oRecItem.Fields.Item("FCCredit").Value
                        strType = "Cr"
                    End If
                Else
                    strCurrency = strLocalCurrency
                    If oRecItem.Fields.Item("Debit").Value <> 0 Then
                        dblAmount = oRecItem.Fields.Item("Debit").Value
                        strType = "Dr"
                    Else
                        dblAmount = oRecItem.Fields.Item("Credit").Value
                        strType = "Cr"
                    End If
                End If
                writer.WriteStartElement("FinancialEntry")

                Dim straccount, sAPAccount, CostCenter1, CostCenter2 As String

                '                If oRecItem.Fields.Item("OutGoing").Value = "" Then
                If oRecItem.Fields.Item("TransType").Value <> "46" Then
                    writer.WriteStartElement("CurCode")
                    writer.WriteString(strCurrency)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Amt")
                    writer.WriteString(dblAmount.ToString)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Description")
                    Dim strRemk As String = GetJournalRemarks()
                    If strRemk <> "" Then
                        strRemk = strRemk & "-" & oRecItem.Fields.Item("LineMemo").Value
                    Else
                        strRemk = oRecItem.Fields.Item("LineMemo").Value
                    End If

                    writer.WriteString(strRemk)
                    ' writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                    writer.WriteEndElement()


                    straccount = "01-01"
                    'If oRecItem.Fields.Item("OutGoing").Value = "" Then
                    If oRecItem.Fields.Item("TransType").Value <> "46" Then
                        sAPAccount = oRecItem.Fields.Item("Account").Value
                        CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                        CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                        If CostCenter1.Length > 0 Then
                            If CostCenter1.Length > 3 Then
                                CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                            Else
                                CostCenter1 = CostCenter1
                            End If
                        Else
                            CostCenter1 = ""
                        End If
                        If CostCenter2.Length > 0 Then
                            If CostCenter2.Length > 3 Then
                                CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                            Else
                                CostCenter2 = CostCenter2
                            End If
                        Else
                            CostCenter2 = ""
                        End If
                        straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                    Else
                        straccount = oRecItem.Fields.Item("OutGoing").Value
                    End If


                    writer.WriteStartElement("AcctId")
                    '   writer.WriteString(oRecItem.Fields.Item("Account").Value)
                    writer.WriteString(straccount)
                    writer.WriteEndElement()

                    '  If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                    If oRecItem.Fields.Item("TransType").Value = "46" Then
                        writer.WriteStartElement("ApplType")
                        writer.WriteString("SV")
                        writer.WriteEndElement()
                        writer.WriteStartElement("AcctType")
                        writer.WriteString("SAV")
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("AcctType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    End If

                    writer.WriteStartElement("EntryType")
                    writer.WriteString(strType)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EffDt")
                    writer.WriteString(dtDateTime)
                    writer.WriteEndElement()

                    writer.WriteStartElement("ApplType")
                    writer.WriteString(strAccountType)
                    writer.WriteEndElement()
                Else
                    writer.WriteStartElement("CurCode")
                    writer.WriteString(strCurrency)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Amt")
                    writer.WriteString(dblAmount.ToString)
                    writer.WriteEndElement()



                    straccount = "01-01"
                    '   If oRecItem.Fields.Item("OutGoing").Value = "" Then
                    If oRecItem.Fields.Item("TransType").Value <> "46" Then
                        sAPAccount = oRecItem.Fields.Item("Account").Value
                        CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                        CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                        If CostCenter1.Length > 0 Then
                            If CostCenter1.Length > 3 Then
                                CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                            Else
                                CostCenter1 = CostCenter1
                            End If
                        Else
                            CostCenter1 = ""
                        End If
                        If CostCenter2.Length > 0 Then
                            If CostCenter2.Length > 3 Then
                                CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                            Else
                                CostCenter2 = CostCenter2
                            End If
                        Else
                            CostCenter2 = ""
                        End If
                        straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                    Else

                        If strType = "Cr" Then
                            straccount = oRecItem.Fields.Item("OutGoing").Value
                            If straccount = "" Then
                                straccount = "01-01"
                                sAPAccount = oRecItem.Fields.Item("Account").Value
                                CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                                CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                                If CostCenter1.Length > 0 Then
                                    If CostCenter1.Length > 3 Then
                                        CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                    Else
                                        CostCenter1 = CostCenter1
                                    End If
                                Else
                                    CostCenter1 = ""
                                End If
                                If CostCenter2.Length > 0 Then
                                    If CostCenter2.Length > 3 Then
                                        CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                    Else
                                        CostCenter2 = CostCenter2
                                    End If
                                Else
                                    CostCenter2 = ""
                                End If
                                straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                            End If
                        Else
                            '  straccount = sAPAccount ' oRecItem.Fields.Item("OutGoing").Value
                            straccount = "01-01"
                            sAPAccount = oRecItem.Fields.Item("Account").Value
                            CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                            CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                            If CostCenter1.Length > 0 Then
                                If CostCenter1.Length > 3 Then
                                    CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                Else
                                    CostCenter1 = CostCenter1
                                End If
                            Else
                                CostCenter1 = ""
                            End If
                            If CostCenter2.Length > 0 Then
                                If CostCenter2.Length > 3 Then
                                    CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                Else
                                    CostCenter2 = CostCenter2
                                End If
                            Else
                                CostCenter2 = ""
                            End If
                            straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount

                        End If

                    End If


                    writer.WriteStartElement("AcctId")
                    writer.WriteString(straccount)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EntryType")
                    writer.WriteString(strType)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EffDt")
                    writer.WriteString(dtDateTime)
                    writer.WriteEndElement()

                    '  If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                    If oRecItem.Fields.Item("TransType").Value = "46" Then
                        If strType = "Dr" Then
                            writer.WriteStartElement("ApplType")
                            writer.WriteString("GL")
                            writer.WriteEndElement()
                            writer.WriteStartElement("AcctType")
                            writer.WriteString("GL")
                            writer.WriteEndElement()
                        Else
                            If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                                writer.WriteStartElement("ApplType")
                                If straccount.StartsWith("1") = True Then
                                    writer.WriteString("CK")
                                ElseIf straccount.StartsWith("2") = True Then
                                    writer.WriteString("SV")
                                Else
                                    writer.WriteString("CK")
                                End If

                                writer.WriteEndElement()
                                writer.WriteStartElement("AcctType")
                                Dim Test As SAPbobsCOM.Recordset
                                Test = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim sa As String
                                sa = "Select * from OCRB where  Account='" & straccount & "'"
                                Test.DoQuery(sa)
                                If Test.RecordCount > 0 Then
                                    If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                        writer.WriteString("MGC")
                                    Else
                                        writer.WriteString(Test.Fields.Item("MandateID").Value)
                                    End If
                                Else

                                    If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                        writer.WriteString("MGC")
                                    Else
                                        writer.WriteString(Test.Fields.Item("CUR").Value)
                                        writer.WriteString("CUR")
                                    End If
                                    'writer.WriteString("CUR")
                                End If



                                writer.WriteEndElement()

                            Else
                                writer.WriteStartElement("ApplType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                                writer.WriteStartElement("AcctType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                            End If

                            strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                            writer.WriteStartElement("ChkNum")
                            writer.WriteString(strCheckNo)
                            writer.WriteEndElement()
                        End If

                    Else
                        writer.WriteStartElement("AcctType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    End If

                    If oRecItem.Fields.Item("TransType").Value = "46" Then
                        'strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                        'If strType = "Dr" Then
                        '    strCheckNo = "MANAGER CHEQUE NO:" & strCheckNo & " Branch:911"
                        'Else
                        '    strCheckNo = "MGC:" & strCheckNo & " Br:911" & " Bnf :" & oRecItem.Fields.Item("LineMemo").Value
                        'End If

                        Dim strs As SAPbobsCOM.Recordset
                        strs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        strs.DoQuery("Select * from OVPM where DocEntry=" & oRecItem.Fields.Item("BaseRef").Value)

                        If strs.Fields.Item("CheckSum").Value > 0 Then
                            strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                            If strType = "Dr" Then
                                strCheckNo = "MANAGER CHEQUE NO:" & strCheckNo & " Branch:911"
                            Else
                                strCheckNo = "MGC:" & strCheckNo & " Br:911" & " Bnf :" & strs.Fields.Item("Comments").Value 'oRecItem.Fields.Item("LineMemo").Value
                            End If
                        Else
                            strCheckNo = ""
                        End If
                        writer.WriteStartElement("Description")
                        If strCheckNo = "" Then
                            writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                        Else
                            writer.WriteString(strCheckNo)
                        End If
                        '  
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("Description")
                        writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                        writer.WriteEndElement()
                    End If

                End If
                writer.WriteEndElement()
                oRecItem.MoveNext()
            Next
            ' writer.WriteString(dtDateTime)
            writer.WriteStartElement("BankInfo")
            writer.WriteStartElement("BranchId")
            writer.WriteString("911")
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("EmployeeIdent")
            writer.WriteStartElement("EmployeeIdentlNum")
            writer.WriteString(strPhxId)
            writer.WriteEndElement()
            writer.WriteStartElement("SuperEmployeeIdentlNum")
            writer.WriteString("0")
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("RqUID")
            writer.WriteString("85fb650d-336e-4c1d-bea3-6d9551b9fb08")
            writer.WriteEndElement()

            writer.WriteEndElement()

            writer.WriteStartElement("RqUID")
            writer.WriteString("85fb650d-336e-4c1d-bea3-6d9551b9fb08")
            writer.WriteEndElement()


            writer.WriteEndElement()
            writer.WriteEndElement()
            'writer.WriteEndElement()

            writer.Flush()
            MyXMLString = myStringWriter.ToString()

            myStringWriter.Close()
            writer.Close()
            ' SendXMLtoIFX(MyXMLString)

            Dim IFXResponse As String = SendXMLtoIFX(MyXMLString)
            Dim doc As New XmlDocument
            doc.LoadXml(IFXResponse)
            Try
                'Dim x As String = doc.SelectSingleNode("IFX/BankScvRs/RqUID)").InnerText.ToString
                Dim locx, locy, locy1 As String
                locx = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusCode").InnerText).ToString
                locy = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/Severity").InnerText).ToString
                locy1 = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusDesc").InnerText).ToString
                If locx = "0" Then
                    oRecItem.DoQuery("Update OJDT set U_Export='Y',U_Z_IFXReply='" & locy1 & "' where TransId in (" & strJNo & ")")
                    strMessage = "Export Jounral Entry  Compleated . Journal Entry No : " & strJNo
                    oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Return True
                Else
                    'strMessage = "Error while transfering Journal Entry No : " & strJNo & " : Error : " & locy1
                    strMessage = "Error while transfering Journal Entry No : " & strJNo & " : Error : " & locy1 & " . The transaction has been reversed. Create a new Transaction"

                    oApplication.SBO_Application.MessageBox(strMessage)
                    oRecItem.DoQuery("Update OJDT set U_Z_IFXReply='" & locy1 & "' where TransId in (" & strJNo & ")")
                    Return False
                End If
                'WriteErrorlog(strMessage, strErrorFileName)
                oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception
            End Try
            oMainRec.MoveNext()
        Next
        Return True
        'If File.Exists(strTRGFileName) Then
        '    File.Delete(strTRGFileName)
        'End If
    End Function




    Public Function ExportJournalEntries_Testing_JournalEntry(ByVal aPath As String) As Boolean
        Dim oRecItem, oRecItemCode, oTemp, oMainRec As SAPbobsCOM.Recordset
        Dim strSQL, strFilename, strUserID As String
        Dim sValue As String
        Dim sPath, strLogDirectory, strPath, strMessage, strSelectedFolderPath, strExportFilePaty As String
        Dim blnErrorflag As Boolean
        Dim FILedatetiem As String
        strPath = System.Windows.Forms.Application.StartupPath
        strFilename = Now.ToLongDateString
        strPath = aPath
        ' strPath = aFileName
        oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMainRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'aPath = 1

        'Foreign Currency Transaction Validation
        strSQL = "Select T1.TransId ,T1.RefDate,Count(*) from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where isnull(T0.FCCurrency,'')<>'' and  Isnull(T1.U_Export,'N')='N' and T0.TransId=" & CInt(aPath) & " group by T1.TransId,T1.RefDate" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
        oMainRec.DoQuery(strSQL)
        If oMainRec.RecordCount > 0 Then
            If ExportJournalEntries_Testing_MultiCurrency(aPath) = True Then
                Return True
            Else
                Return False
            End If
        End If
        'Foreign Currency Transaction Validation


        strSQL = "Select T1.TransId ,T1.RefDate,Count(*) from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where Isnull(T1.U_Export,'N')='N' and T0.TransId=" & CInt(aPath) & " group by T1.TransId,T1.RefDate" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
        oMainRec.DoQuery(strSQL)
        Dim strTransID As String
        Dim dtJEDate As Date
        For intLoop As Integer = 0 To oMainRec.RecordCount - 1
            strTransID = oMainRec.Fields.Item(0).Value
            dtJEDate = oMainRec.Fields.Item(1).Value

            Dim strPhxId As String = ""
            'BP customer Master
            'strSQL = "Select * from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  where Isnull(T1.U_Export,'N')='N'" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"

            strSQL = "Select *,T0.Ref2 'OutGoing',isnull(T2.U_PhxId,'911') 'Uid' from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where T1.TransId=" & oMainRec.Fields.Item(0).Value & " and  Isnull(T1.U_Export,'N')='N'" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
            strSQL = "Select *,T0.Ref2 'OutGoing',T1.TransType 'TransType',T0.ShortName 'BPName',isnull(T2.U_PhxId,'911') 'Uid',T0.Ref3Line 'CheckNo' from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where  isnull(T0.VatGroup,'')<>'X1' and Isnull(T1.U_Export,'N')='N' and  T1.TransId=" & oMainRec.Fields.Item(0).Value  'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"


            oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItem.DoQuery(strSQL)
            FILedatetiem = Now.ToString("yyyyMMdd_hh_mm_ss")
            Dim transtype As String
            Dim strTRGFileName As String
            Dim dtDateTime As String
            Dim strJNo As String = ""
            Dim strCheckNo As String = ""
            dtDateTime = dtJEDate.ToString("yyyy-MM-dd") & " _" & strTransID ' Format(Now.Date, "yyyy-MM-dd") & Now.ToLongTimeString.Replace(":", "")
            FILedatetiem = dtDateTime
            dtDateTime = dtJEDate.ToString("yyyy-MM-dd")
            strTRGFileName = strExportFilePaty & "\JE_" & FILedatetiem & ".trg"
            strFilename = strExportFilePaty & "\JE_" & FILedatetiem & ".xml"
            Dim MyXMLString As String
            Dim myStringWriter As New StringWriter
            Dim writer As New XmlTextWriter(myStringWriter)
            '(strFilename, System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2
            writer.WriteStartElement("IFX")

            writer.WriteStartElement("SignonRq")
            writer.WriteStartElement("RqUID")
            writer.WriteString("6a8b5973-ec37-47e0-855e-4cc020ab7f02")
            writer.WriteEndElement()
            writer.WriteStartElement("ClientDt")
            writer.WriteString(dtJEDate.ToString("yyyy-MM-dd"))
            writer.WriteEndElement()
            writer.WriteStartElement("ClientApp")
            writer.WriteString("SAP")
            writer.WriteEndElement()
            writer.WriteStartElement("OperatorId")
            writer.WriteString("Manager")
            writer.WriteEndElement()
            writer.WriteEndElement()

            'BankSvcRq
            writer.WriteStartElement("BankSvcRq")
            writer.WriteStartElement("FinancialMessageAddRq")
            Dim strFCCurrency, strLocalCurrency, strCurrency, strType, strAccountType As String
            Dim dblAmount As Double
            oTemp.DoQuery("Select MainCurncy from OADM")
            strLocalCurrency = oTemp.Fields.Item(0).Value

            For intRow As Integer = 0 To oRecItem.RecordCount - 1

                strPhxId = oRecItem.Fields.Item("U_PhxId").Value
                If strJNo = "" Then
                    strJNo = oRecItem.Fields.Item("TransId").Value
                Else
                    strJNo = strJNo & "," & oRecItem.Fields.Item("TransId").Value
                End If
                If oRecItem.Fields.Item("Account").Value <> oRecItem.Fields.Item("ShortName").Value Then
                    strAccountType = "GL"
                Else
                    strAccountType = "GL"
                End If
                strFCCurrency = oRecItem.Fields.Item("FCCurrency").Value
                If strFCCurrency <> "" Then
                    strCurrency = strFCCurrency
                    If oRecItem.Fields.Item("FCDebit").Value <> 0 Then
                        dblAmount = oRecItem.Fields.Item("FCDebit").Value
                        strType = "Dr"
                    Else
                        dblAmount = oRecItem.Fields.Item("FCCredit").Value
                        strType = "Cr"
                    End If
                Else
                    strCurrency = strLocalCurrency
                    If oRecItem.Fields.Item("Debit").Value <> 0 Then
                        dblAmount = oRecItem.Fields.Item("Debit").Value
                        strType = "Dr"
                    Else
                        dblAmount = oRecItem.Fields.Item("Credit").Value
                        strType = "Cr"
                    End If
                End If
                writer.WriteStartElement("FinancialEntry")

                Dim straccount, sAPAccount, CostCenter1, CostCenter2 As String

                '                If oRecItem.Fields.Item("OutGoing").Value = "" Then
                If oRecItem.Fields.Item("TransType").Value <> "30" Then
                    writer.WriteStartElement("CurCode")
                    writer.WriteString(strCurrency)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Amt")
                    writer.WriteString(dblAmount.ToString)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Description")
                    Dim strRemk As String = GetJournalRemarks()
                    If strRemk <> "" Then
                        strRemk = strRemk & "-" & oRecItem.Fields.Item("LineMemo").Value
                    Else
                        strRemk = oRecItem.Fields.Item("LineMemo").Value
                    End If

                    writer.WriteString(strRemk)
                    '  writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                    writer.WriteEndElement()


                    straccount = "01-01"
                    'If oRecItem.Fields.Item("OutGoing").Value = "" Then
                    If oRecItem.Fields.Item("TransType").Value <> "30" Then
                        sAPAccount = oRecItem.Fields.Item("Account").Value
                        CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                        CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                        If CostCenter1.Length > 0 Then
                            If CostCenter1.Length > 3 Then
                                CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                            Else
                                CostCenter1 = CostCenter1
                            End If
                        Else
                            CostCenter1 = ""
                        End If
                        If CostCenter2.Length > 0 Then
                            If CostCenter2.Length > 3 Then
                                CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                            Else
                                CostCenter2 = CostCenter2
                            End If
                        Else
                            CostCenter2 = ""
                        End If
                        straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                    Else
                        straccount = oRecItem.Fields.Item("OutGoing").Value
                    End If


                    writer.WriteStartElement("AcctId")
                    '   writer.WriteString(oRecItem.Fields.Item("Account").Value)
                    writer.WriteString(straccount)
                    writer.WriteEndElement()

                    '  If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                    If oRecItem.Fields.Item("TransType").Value = "30" Then
                        writer.WriteStartElement("ApplType")
                        writer.WriteString("SV")
                        writer.WriteEndElement()
                        writer.WriteStartElement("AcctType")
                        writer.WriteString("SAV")
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("AcctType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    End If

                    writer.WriteStartElement("EntryType")
                    writer.WriteString(strType)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EffDt")
                    writer.WriteString(dtDateTime)
                    writer.WriteEndElement()

                    writer.WriteStartElement("ApplType")
                    writer.WriteString(strAccountType)
                    writer.WriteEndElement()
                Else
                    writer.WriteStartElement("CurCode")
                    writer.WriteString(strCurrency)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Amt")
                    writer.WriteString(dblAmount.ToString)
                    writer.WriteEndElement()



                    straccount = "01-01"
                    '   If oRecItem.Fields.Item("OutGoing").Value = "" Then
                    If oRecItem.Fields.Item("TransType").Value <> "30" Then
                        sAPAccount = oRecItem.Fields.Item("Account").Value
                        CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                        CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                        If CostCenter1.Length > 0 Then
                            If CostCenter1.Length > 3 Then
                                CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                            Else
                                CostCenter1 = CostCenter1
                            End If
                        Else
                            CostCenter1 = ""
                        End If
                        If CostCenter2.Length > 0 Then
                            If CostCenter2.Length > 3 Then
                                CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                            Else
                                CostCenter2 = CostCenter2
                            End If
                        Else
                            CostCenter2 = ""
                        End If
                        straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                    Else

                        If strType = "Cr" Then
                            straccount = oRecItem.Fields.Item("OutGoing").Value
                            If straccount = "" Then
                                straccount = "01-01"
                                sAPAccount = oRecItem.Fields.Item("Account").Value
                                CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                                CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                                If CostCenter1.Length > 0 Then
                                    If CostCenter1.Length > 3 Then
                                        CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                    Else
                                        CostCenter1 = CostCenter1
                                    End If
                                Else
                                    CostCenter1 = ""
                                End If
                                If CostCenter2.Length > 0 Then
                                    If CostCenter2.Length > 3 Then
                                        CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                    Else
                                        CostCenter2 = CostCenter2
                                    End If
                                Else
                                    CostCenter2 = ""
                                End If
                                straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                            End If
                        Else
                            '  straccount = sAPAccount ' oRecItem.Fields.Item("OutGoing").Value
                            straccount = "01-01"
                            sAPAccount = oRecItem.Fields.Item("Account").Value
                            CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                            CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                            If CostCenter1.Length > 0 Then
                                If CostCenter1.Length > 3 Then
                                    CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                Else
                                    CostCenter1 = CostCenter1
                                End If
                            Else
                                CostCenter1 = ""
                            End If
                            If CostCenter2.Length > 0 Then
                                If CostCenter2.Length > 3 Then
                                    CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                Else
                                    CostCenter2 = CostCenter2
                                End If
                            Else
                                CostCenter2 = ""
                            End If
                            straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount

                        End If

                    End If


                    writer.WriteStartElement("AcctId")
                    writer.WriteString(straccount)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EntryType")
                    writer.WriteString(strType)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EffDt")
                    writer.WriteString(dtDateTime)
                    writer.WriteEndElement()

                    '  If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                    If oRecItem.Fields.Item("TransType").Value = "30" Then
                        If strType = "Dr" Then
                            writer.WriteStartElement("ApplType")
                            writer.WriteString("GL")
                            writer.WriteEndElement()
                            writer.WriteStartElement("AcctType")
                            writer.WriteString("GL")
                            writer.WriteEndElement()
                        Else
                            If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                                writer.WriteStartElement("ApplType")
                                If straccount.StartsWith("1") = True Then
                                    writer.WriteString("CK")
                                ElseIf straccount.StartsWith("2") = True Then
                                    writer.WriteString("SV")
                                Else
                                    writer.WriteString("CK")
                                End If

                                writer.WriteEndElement()
                                writer.WriteStartElement("AcctType")
                                Dim Test As SAPbobsCOM.Recordset
                                Test = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim sa As String
                                sa = "Select * from OCRB where  Account='" & straccount & "'"
                                Test.DoQuery(sa)
                                If Test.RecordCount > 0 Then
                                    If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                        writer.WriteString("MGC")
                                    Else
                                        writer.WriteString(Test.Fields.Item("MandateID").Value)
                                    End If
                                Else

                                    If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                        writer.WriteString("MGC")
                                    Else
                                        ' writer.WriteString(Test.Fields.Item("CUR").Value)
                                        writer.WriteString("CUR")
                                    End If
                                    'writer.WriteString("CUR")
                                End If



                                writer.WriteEndElement()

                            Else
                                writer.WriteStartElement("ApplType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                                writer.WriteStartElement("AcctType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                            End If

                            strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                            writer.WriteStartElement("ChkNum")
                            writer.WriteString(strCheckNo)
                            writer.WriteEndElement()
                        End If

                    Else
                        writer.WriteStartElement("AcctType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    End If

                    If oRecItem.Fields.Item("TransType").Value = "30" Then
                        'strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                        'If strType = "Dr" Then
                        '    strCheckNo = "MANAGER CHEQUE NO:" & strCheckNo & " Branch:911"
                        'Else
                        '    strCheckNo = "MGC:" & strCheckNo & " Br:911" & " Bnf :" & oRecItem.Fields.Item("LineMemo").Value
                        'End If

                        'Dim strs As SAPbobsCOM.Recordset
                        'strs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'strs.DoQuery("Select * from OVPM where DocEntry=" & oRecItem.Fields.Item("BaseRef").Value)

                        'If strs.Fields.Item("CheckSum").Value > 0 Then
                        '    strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                        '    If strType <> "Dr" Then
                        '        strCheckNo = "MANAGER CHEQUE NO:" & strCheckNo & " Branch:911"
                        '    Else
                        '        strCheckNo = "MGC:" & strCheckNo & " Br:911" & " Bnf :" & strs.Fields.Item("Comments").Value 'oRecItem.Fields.Item("LineMemo").Value
                        '    End If
                        'Else
                        '    strCheckNo = ""
                        'End Ifch
                        strCheckNo = ""
                        writer.WriteStartElement("Description")
                        If strCheckNo = "" Then
                            Dim strRemk As String = GetJournalRemarks()
                            If strRemk <> "" Then
                                strRemk = strRemk & "-" & oRecItem.Fields.Item("LineMemo").Value
                            Else
                                strRemk = oRecItem.Fields.Item("LineMemo").Value
                            End If

                            writer.WriteString(strRemk)
                            ' writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                        Else
                            writer.WriteString(strCheckNo)
                        End If
                        '  
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("Description")
                        writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                        writer.WriteEndElement()
                    End If

                End If
                writer.WriteEndElement()
                oRecItem.MoveNext()
            Next
            ' writer.WriteString(dtDateTime)
            writer.WriteStartElement("BankInfo")
            writer.WriteStartElement("BranchId")
            writer.WriteString("911")
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("EmployeeIdent")
            writer.WriteStartElement("EmployeeIdentlNum")
            writer.WriteString(strPhxId)
            writer.WriteEndElement()
            writer.WriteStartElement("SuperEmployeeIdentlNum")
            writer.WriteString("0")
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("RqUID")
            writer.WriteString("85fb650d-336e-4c1d-bea3-6d9551b9fb08")
            writer.WriteEndElement()

            writer.WriteEndElement()

            writer.WriteStartElement("RqUID")
            writer.WriteString("85fb650d-336e-4c1d-bea3-6d9551b9fb08")
            writer.WriteEndElement()


            writer.WriteEndElement()
            writer.WriteEndElement()
            'writer.WriteEndElement()

            writer.Flush()
            MyXMLString = myStringWriter.ToString()

            myStringWriter.Close()
            writer.Close()
            ' SendXMLtoIFX(MyXMLString)

            Dim IFXResponse As String = SendXMLtoIFX(MyXMLString)
            Dim doc As New XmlDocument
            doc.LoadXml(IFXResponse)
            Try
                'Dim x As String = doc.SelectSingleNode("IFX/BankScvRs/RqUID)").InnerText.ToString
                Dim locx, locy, locy1 As String
                locx = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusCode").InnerText).ToString
                locy = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/Severity").InnerText).ToString
                locy1 = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusDesc").InnerText).ToString
                If locx = "0" Then
                    oRecItem.DoQuery("Update OJDT set U_Export='Y',U_Z_IFXReply='" & locy1 & "' where TransId in (" & strJNo & ")")
                    strMessage = "Export Jounral Entry  Compleated . Journal Entry No : " & strJNo
                    oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Return True
                Else
                    'strMessage = "Error while transfering Journal Entry No : " & strJNo & " : Error : " & locy1
                    strMessage = "Error while transfering Journal Entry No : " & strJNo & " : Error : " & locy1 & " . The transaction has been reversed. Create a new Transaction"

                    oApplication.SBO_Application.MessageBox(strMessage)
                    oRecItem.DoQuery("Update OJDT set U_Z_IFXReply='" & locy1 & "' where TransId in (" & strJNo & ")")
                    Return False
                End If
                'WriteErrorlog(strMessage, strErrorFileName)
                oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception
            End Try
            oMainRec.MoveNext()
        Next
        Return True
        'If File.Exists(strTRGFileName) Then
        '    File.Delete(strTRGFileName)
        'End If
    End Function

    Public Function GetPayrollCode(ByVal aNo As String) As String
        Dim oRe1 As SAPbobsCOM.Recordset
        oRe1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRe1.DoQuery("Select Code from [@Z_PAYROLL1] where U_Z_JVNo='" & aNo & "'")
        Dim strCode As String = "'111111111'"
        For intLoop As Integer = 0 To oRe1.RecordCount - 1
            If strCode = "" Then
                strCode = "'" & oRe1.Fields.Item(0).Value & "'"
            Else
                strCode = strCode & ",'" & oRe1.Fields.Item(0).Value & "'"
            End If
            oRe1.MoveNext()
        Next
        Return strCode
    End Function

    Public Function checkDirectoryexists(ByVal aPath As String) As Boolean
        If IO.Directory.Exists(aPath) Then
            Return True
        Else
            Message("Directoy does not exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
    End Function

    Public Function ExportJournalEntries_Testing_Payroll(ByVal aPath As String, ByVal aCode As String, ByVal aPostingDate As Date) As Boolean
        Dim oRecItem, oRecItemCode, oTemp, oMainRec As SAPbobsCOM.Recordset
        Dim strSQL, strFilename, strUserID As String
        Dim sValue As String
        Dim sPath, strLogDirectory, strPath, strMessage, strSelectedFolderPath, strExportFilePaty As String
        Dim blnErrorflag As Boolean
        Dim FILedatetiem As String
        strPath = System.Windows.Forms.Application.StartupPath
        strFilename = Now.ToLongDateString
        strPath = aPath
        ' strPath = aFileName
        oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMainRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'aPath = 1

        'Foreign Currency Transaction Validation
        'strSQL = "Select T1.TransId ,T1.RefDate,Count(*) from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where isnull(T0.FCCurrency,'')<>'' and  Isnull(T1.U_Export,'N')='N' and T0.TransId=" & CInt(aPath) & " group by T1.TransId,T1.RefDate" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
        'oMainRec.DoQuery(strSQL)
        'If oMainRec.RecordCount > 0 Then
        '    If ExportJournalEntries_Testing_MultiCurrency(aPath) = True Then
        '        Return True
        '    Else
        '        Return False
        '    End If
        'End If
        'Foreign Currency Transaction Validation


        strSQL = "Select T1.TransId ,T1.RefDate,Count(*) from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where Isnull(T1.U_Export,'N')='N' and T0.TransId=" & CInt(aPath) & " group by T1.TransId,T1.RefDate" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
        oMainRec.DoQuery(strSQL)
        Dim strTransID As String
        Dim dtJEDate As Date
        For intLoop As Integer = 0 To oMainRec.RecordCount - 1
            strTransID = oMainRec.Fields.Item(0).Value
            dtJEDate = aPostingDate ' oMainRec.Fields.Item(1).Value
            'dtJEDate = oMainRec.Fields.Item(1).Value

            Dim strPhxId As String = ""
            'BP customer Master
            'strSQL = "Select * from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  where Isnull(T1.U_Export,'N')='N'" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"

            strSQL = "Select *,T0.Ref2 'OutGoing',isnull(T2.U_PhxId,'911') 'Uid' from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where T1.TransId=" & oMainRec.Fields.Item(0).Value & " and  Isnull(T1.U_Export,'N')='N'" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
            strSQL = "Select *,T0.Ref2 'OutGoing',T1.TransType 'TransType',T0.ShortName 'BPName',isnull(T2.U_PhxId,'911') 'Uid',T0.Ref3Line 'CheckNo' from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where  isnull(T0.VatGroup,'')<>'X1' and Isnull(T1.U_Export,'N')='N' and  T1.TransId=" & oMainRec.Fields.Item(0).Value  'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"


            oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItem.DoQuery(strSQL)
            FILedatetiem = Now.ToString("yyyyMMdd_hh_mm_ss")
            Dim transtype As String
            Dim strTRGFileName As String
            Dim dtDateTime As String
            Dim strJNo As String = ""
            Dim strCheckNo As String = ""
            dtDateTime = dtJEDate.ToString("yyyy-MM-dd") & " _" & strTransID ' Format(Now.Date, "yyyy-MM-dd") & Now.ToLongTimeString.Replace(":", "")
            FILedatetiem = dtDateTime
            dtDateTime = dtJEDate.ToString("yyyy-MM-dd")
            strTRGFileName = strExportFilePaty & "\JE_" & FILedatetiem & ".trg"
            strFilename = strExportFilePaty & "\JE_" & FILedatetiem & ".xml"
            Dim MyXMLString As String
            Dim myStringWriter As New StringWriter
            Dim writer As New XmlTextWriter(myStringWriter)
            '(strFilename, System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2
            writer.WriteStartElement("IFX")

            writer.WriteStartElement("SignonRq")
            writer.WriteStartElement("RqUID")
            writer.WriteString("6a8b5973-ec37-47e0-855e-4cc020ab7f02")
            writer.WriteEndElement()
            writer.WriteStartElement("ClientDt")
            writer.WriteString(dtJEDate.ToString("yyyy-MM-dd"))
            writer.WriteEndElement()
            writer.WriteStartElement("ClientApp")
            writer.WriteString("SAP")
            writer.WriteEndElement()
            writer.WriteStartElement("OperatorId")
            writer.WriteString("Manager")
            writer.WriteEndElement()
            writer.WriteEndElement()

            'BankSvcRq
            writer.WriteStartElement("BankSvcRq")
            writer.WriteStartElement("FinancialMessageAddRq")
            Dim strFCCurrency, strLocalCurrency, strCurrency, strType, strAccountType As String
            Dim dblAmount As Double
            oTemp.DoQuery("Select MainCurncy from OADM")
            strLocalCurrency = oTemp.Fields.Item(0).Value

            For intRow As Integer = 0 To oRecItem.RecordCount - 1

                strPhxId = oRecItem.Fields.Item("U_PhxId").Value
                If strJNo = "" Then
                    strJNo = oRecItem.Fields.Item("TransId").Value
                Else
                    strJNo = strJNo & "," & oRecItem.Fields.Item("TransId").Value
                End If
                If oRecItem.Fields.Item("Account").Value <> oRecItem.Fields.Item("ShortName").Value Then
                    strAccountType = "GL"
                Else
                    strAccountType = "GL"
                End If
                strFCCurrency = oRecItem.Fields.Item("FCCurrency").Value
                If strFCCurrency <> "" Then
                    strCurrency = strFCCurrency
                    If oRecItem.Fields.Item("FCDebit").Value <> 0 Then
                        dblAmount = oRecItem.Fields.Item("FCDebit").Value
                        strType = "Dr"
                    Else
                        dblAmount = oRecItem.Fields.Item("FCCredit").Value
                        strType = "Cr"
                    End If
                Else
                    strCurrency = strLocalCurrency
                    If oRecItem.Fields.Item("Debit").Value <> 0 Then
                        dblAmount = oRecItem.Fields.Item("Debit").Value
                        strType = "Dr"
                    Else
                        dblAmount = oRecItem.Fields.Item("Credit").Value
                        strType = "Cr"
                    End If
                End If
                dblAmount = Math.Round(dblAmount, 3)

                writer.WriteStartElement("FinancialEntry")

                Dim straccount, sAPAccount, CostCenter1, CostCenter2 As String

                '                If oRecItem.Fields.Item("OutGoing").Value = "" Then
                If oRecItem.Fields.Item("TransType").Value <> "46" Then
                    writer.WriteStartElement("CurCode")
                    writer.WriteString(strCurrency)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Amt")
                    dblAmount = Math.Round(dblAmount, 3)
                    writer.WriteString(dblAmount.ToString)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Description")
                    Dim strRemk As String = GetJournalRemarks()
                    If strRemk <> "" Then
                        strRemk = strRemk & "-" & oRecItem.Fields.Item("LineMemo").Value
                    Else
                        strRemk = oRecItem.Fields.Item("LineMemo").Value
                    End If

                    writer.WriteString(strRemk)
                    '  writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                    writer.WriteEndElement()


                    straccount = "01-01"
                    'If oRecItem.Fields.Item("OutGoing").Value = "" Then
                    If oRecItem.Fields.Item("TransType").Value <> "46" Then
                        sAPAccount = oRecItem.Fields.Item("Account").Value
                        CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                        CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                        If CostCenter1.Length > 0 Then
                            If CostCenter1.Length > 3 Then
                                CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                            Else
                                CostCenter1 = CostCenter1
                            End If
                        Else
                            CostCenter1 = ""
                        End If
                        If CostCenter2.Length > 0 Then
                            If CostCenter2.Length > 3 Then
                                CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                            Else
                                CostCenter2 = CostCenter2
                            End If
                        Else
                            CostCenter2 = ""
                        End If
                        straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                    Else
                        straccount = oRecItem.Fields.Item("OutGoing").Value
                    End If


                    writer.WriteStartElement("AcctId")
                    '   writer.WriteString(oRecItem.Fields.Item("Account").Value)
                    writer.WriteString(straccount)
                    writer.WriteEndElement()

                    '  If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                    If oRecItem.Fields.Item("TransType").Value = "46" Then
                        writer.WriteStartElement("ApplType")
                        writer.WriteString("SV")
                        writer.WriteEndElement()
                        writer.WriteStartElement("AcctType")
                        writer.WriteString("SAV")
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("AcctType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    End If

                    writer.WriteStartElement("EntryType")
                    writer.WriteString(strType)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EffDt")
                    writer.WriteString(dtDateTime)
                    writer.WriteEndElement()

                    writer.WriteStartElement("ApplType")
                    writer.WriteString(strAccountType)
                    writer.WriteEndElement()
                Else
                    writer.WriteStartElement("CurCode")
                    writer.WriteString(strCurrency)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Amt")
                    dblAmount = Math.Round(dblAmount, 3)
                    writer.WriteString(dblAmount.ToString)
                    writer.WriteEndElement()



                    straccount = "01-01"
                    '   If oRecItem.Fields.Item("OutGoing").Value = "" Then
                    If oRecItem.Fields.Item("TransType").Value <> "46" Then
                        sAPAccount = oRecItem.Fields.Item("Account").Value
                        CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                        CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                        If CostCenter1.Length > 0 Then
                            If CostCenter1.Length > 3 Then
                                CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                            Else
                                CostCenter1 = CostCenter1
                            End If
                        Else
                            CostCenter1 = ""
                        End If
                        If CostCenter2.Length > 0 Then
                            If CostCenter2.Length > 3 Then
                                CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                            Else
                                CostCenter2 = CostCenter2
                            End If
                        Else
                            CostCenter2 = ""
                        End If
                        straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                    Else

                        If strType = "Cr" Then
                            straccount = oRecItem.Fields.Item("OutGoing").Value
                            If straccount = "" Then
                                straccount = "01-01"
                                sAPAccount = oRecItem.Fields.Item("Account").Value
                                CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                                CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                                If CostCenter1.Length > 0 Then
                                    If CostCenter1.Length > 3 Then
                                        CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                    Else
                                        CostCenter1 = CostCenter1
                                    End If
                                Else
                                    CostCenter1 = ""
                                End If
                                If CostCenter2.Length > 0 Then
                                    If CostCenter2.Length > 3 Then
                                        CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                    Else
                                        CostCenter2 = CostCenter2
                                    End If
                                Else
                                    CostCenter2 = ""
                                End If
                                straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                            End If
                        Else
                            '  straccount = sAPAccount ' oRecItem.Fields.Item("OutGoing").Value
                            straccount = "01-01"
                            sAPAccount = oRecItem.Fields.Item("Account").Value
                            CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                            CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                            If CostCenter1.Length > 0 Then
                                If CostCenter1.Length > 3 Then
                                    CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                Else
                                    CostCenter1 = CostCenter1
                                End If
                            Else
                                CostCenter1 = ""
                            End If
                            If CostCenter2.Length > 0 Then
                                If CostCenter2.Length > 3 Then
                                    CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                Else
                                    CostCenter2 = CostCenter2
                                End If
                            Else
                                CostCenter2 = ""
                            End If
                            straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount

                        End If

                    End If


                    writer.WriteStartElement("AcctId")
                    writer.WriteString(straccount)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EntryType")
                    writer.WriteString(strType)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EffDt")
                    writer.WriteString(dtDateTime)
                    writer.WriteEndElement()

                    '  If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                    If oRecItem.Fields.Item("TransType").Value = "46" Then
                        If strType = "Dr" Then
                            writer.WriteStartElement("ApplType")
                            writer.WriteString("GL")
                            writer.WriteEndElement()
                            writer.WriteStartElement("AcctType")
                            writer.WriteString("GL")
                            writer.WriteEndElement()
                        Else
                            If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                                writer.WriteStartElement("ApplType")
                                If straccount.StartsWith("1") = True Then
                                    writer.WriteString("CK")
                                ElseIf straccount.StartsWith("2") = True Then
                                    writer.WriteString("SV")
                                Else
                                    writer.WriteString("CK")
                                End If

                                writer.WriteEndElement()
                                writer.WriteStartElement("AcctType")
                                Dim Test As SAPbobsCOM.Recordset
                                Test = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim sa As String
                                sa = "Select * from OCRB where  Account='" & straccount & "'"
                                Test.DoQuery(sa)
                                If Test.RecordCount > 0 Then
                                    If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                        writer.WriteString("MGC")
                                    Else
                                        writer.WriteString(Test.Fields.Item("MandateID").Value)
                                    End If
                                Else

                                    If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                        writer.WriteString("MGC")
                                    Else
                                        writer.WriteString(Test.Fields.Item("CUR").Value)
                                        writer.WriteString("CUR")
                                    End If
                                    'writer.WriteString("CUR")
                                End If



                                writer.WriteEndElement()

                            Else
                                writer.WriteStartElement("ApplType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                                writer.WriteStartElement("AcctType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                            End If

                            strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                            writer.WriteStartElement("ChkNum")
                            writer.WriteString(strCheckNo)
                            writer.WriteEndElement()
                        End If

                    Else
                        writer.WriteStartElement("AcctType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    End If

                    If oRecItem.Fields.Item("TransType").Value = "46" Then
                        'strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                        'If strType = "Dr" Then
                        '    strCheckNo = "MANAGER CHEQUE NO:" & strCheckNo & " Branch:911"
                        'Else
                        '    strCheckNo = "MGC:" & strCheckNo & " Br:911" & " Bnf :" & oRecItem.Fields.Item("LineMemo").Value
                        'End If

                        Dim strs As SAPbobsCOM.Recordset
                        strs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        strs.DoQuery("Select * from OVPM where DocEntry=" & oRecItem.Fields.Item("BaseRef").Value)

                        If strs.Fields.Item("CheckSum").Value > 0 Then
                            strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                            If strType <> "Dr" Then
                                strCheckNo = "MANAGER CHEQUE NO:" & strCheckNo & " Branch:911"
                            Else
                                strCheckNo = "MGC:" & strCheckNo & " Br:911" & " Bnf :" & strs.Fields.Item("Comments").Value 'oRecItem.Fields.Item("LineMemo").Value
                            End If
                        Else
                            strCheckNo = ""
                        End If
                        writer.WriteStartElement("Description")
                        If strCheckNo = "" Then
                            writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                        Else
                            writer.WriteString(strCheckNo)
                        End If
                        '  
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("Description")
                        writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                        writer.WriteEndElement()
                    End If

                End If
                writer.WriteEndElement()
                oRecItem.MoveNext()
            Next
            ' writer.WriteString(dtDateTime)
            writer.WriteStartElement("BankInfo")
            writer.WriteStartElement("BranchId")
            writer.WriteString("911")
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("EmployeeIdent")
            writer.WriteStartElement("EmployeeIdentlNum")
            writer.WriteString(strPhxId)
            writer.WriteEndElement()
            writer.WriteStartElement("SuperEmployeeIdentlNum")
            writer.WriteString("0")
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("RqUID")
            writer.WriteString("85fb650d-336e-4c1d-bea3-6d9551b9fb08")
            writer.WriteEndElement()

            writer.WriteEndElement()

            writer.WriteStartElement("RqUID")
            writer.WriteString("85fb650d-336e-4c1d-bea3-6d9551b9fb08")
            writer.WriteEndElement()


            writer.WriteEndElement()
            writer.WriteEndElement()
            'writer.WriteEndElement()

            writer.Flush()
            MyXMLString = myStringWriter.ToString()

            myStringWriter.Close()
            writer.Close()
            ' SendXMLtoIFX(MyXMLString)
            Dim strPaht As String = getXMLPath()
            Dim blnFileExists As Boolean = False
            Dim oDoc1 As New XmlDocument
            If strPaht <> "" Then
                If IO.Directory.Exists(strPaht) Then

                    oDoc1.LoadXml(MyXMLString)
                    blnFileExists = True
                    If IO.Directory.Exists(strPaht & "\Success") = False Then
                        IO.Directory.CreateDirectory(strPaht & "\Success")
                    End If
                    If IO.Directory.Exists(strPaht & "\Error") = False Then
                        IO.Directory.CreateDirectory(strPaht & "\Error")
                    End If
                    'oDoc1.Save(strPaht & "\JE_" & aPath & ".xml")
                End If
            End If
            Dim IFXResponse As String = SendXMLtoIFX(MyXMLString)
            Dim doc As New XmlDocument
            doc.LoadXml(IFXResponse)
            Try
                'Dim x As String = doc.SelectSingleNode("IFX/BankScvRs/RqUID)").InnerText.ToString
                Dim locx, locy, locy1 As String
                locx = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusCode").InnerText).ToString
                locy = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/Severity").InnerText).ToString
                locy1 = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusDesc").InnerText).ToString
                If locx = "0" Then
                    oRecItem.DoQuery("Update OJDT set U_Export='Y',U_Z_IFXReply='" & locy1 & "' where TransId in (" & strJNo & ")")
                    oRecItem.DoQuery("Update [@Z_PAYROLL1] set U_Z_IFXPosting='C',U_Z_IFXResponse='" & locy1 & "' where U_Z_JVNo='" & aPath & "' and isnull(U_Z_IFXPosting,'P')='P' and  Code in (" & aCode & ")")
                    strMessage = "Export Jounral Entry  Compleated . Journal Entry No : " & strJNo
                    oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    If blnFileExists Then


                        If File.Exists(strPaht & "\Success\JE_" & aPath & "_Success.xml") Then
                            File.Delete(strPaht & "\Success\JE_" & aPath & "_Success.xml")

                        End If
                        oDoc1.Save(strPaht & "\Success\JE_" & aPath & "_Success.xml")
                    End If
                    Return True
                Else
                    '  strMessage = "Error while transfering Journal Entry No : " & strJNo & " : Error : " & locy1
                    strMessage = "Error while transfering Journal Entry No : " & strJNo & " : Error : " & locy1 & " . The transaction has been reversed. Create a new Transaction"
                    oApplication.SBO_Application.MessageBox(strMessage)
                    oRecItem.DoQuery("Update OJDT set U_Z_IFXReply='" & locy1 & "' where TransId in (" & strJNo & ")")
                    oRecItem.DoQuery("Update [@Z_PAYROLL1] set U_Z_IFXPosting='P',U_Z_IFXResponse='" & locy1 & "' where  U_Z_JVNo='" & aPath & "' and isnull(U_Z_IFXPosting,'P')='P' and  Code in (" & aCode & ")")
                    If blnFileExists = True Then
                        If File.Exists(strPaht & "\Error\JE_" & aPath & "_Error.xml") Then
                            File.Delete(strPaht & "\Error\JE_" & aPath & "_Error.xml")
                        End If
                        oDoc1.Save(strPaht & "\Error\JE_" & aPath & "_Error.xml")
                    End If
                    Return False
                    End If
                    'WriteErrorlog(strMessage, strErrorFileName)
                    oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception
            End Try
            oMainRec.MoveNext()
        Next
        Return True
        'If File.Exists(strTRGFileName) Then
        '    File.Delete(strTRGFileName)
        'End If
    End Function

    Public Function ExportJournalEntries_Testing_Payroll_Empployee(ByVal aPath As String, ByVal aCode As String, ByVal aPostingDate As Date) As Boolean
        Dim oRecItem, oRecItemCode, oTemp, oMainRec As SAPbobsCOM.Recordset
        Dim strSQL, strFilename, strUserID As String
        Dim sValue As String
        Dim sPath, strLogDirectory, strPath, strMessage, strSelectedFolderPath, strExportFilePaty As String
        Dim blnErrorflag As Boolean
        Dim FILedatetiem As String
        strPath = System.Windows.Forms.Application.StartupPath
        strFilename = Now.ToLongDateString
        strPath = aPath
        ' strPath = aFileName
        oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMainRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strTransID As String
        Dim dtJEDate As Date
        For intLoop As Integer = 1 To 1 'oMainRec.RecordCount - 1
            strTransID = aCode
            '  dtJEDate = oMainRec.Fields.Item(1).Value
            Dim strPhxId As String = ""
            dtJEDate = aPostingDate
            strSQL = "Select * from [@Z_PAYROLL1] where Code='" & strTransID & "'"
            oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItem.DoQuery(strSQL)
            Dim transtype As String
            Dim strTRGFileName As String
            Dim dtDateTime As String
            Dim strJNo As String = ""
            Dim strCheckNo As String = ""
            Dim MyXMLString As String
            Dim myStringWriter As New StringWriter
            Dim writer As New XmlTextWriter(myStringWriter)
            Dim oRec2 As SAPbobsCOM.Recordset
            oRec2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oRecItem.Fields.Item("U_Z_IFXPosting").Value = "C" Then
                Dim strEmpId As String = oRecItem.Fields.Item("U_Z_EmpId").Value
                Dim strEmpBankAccount As String
                oRec2.DoQuery("Select isnull(bankAcount,'') from OHEM where empID=" & CInt(strEmpId))
                strEmpBankAccount = oRec2.Fields.Item(0).Value
                FILedatetiem = Now.ToString("yyyyMMdd_hh_mm_ss")
                dtDateTime = dtJEDate.ToString("yyyy-MM-dd") & " _" & strTransID ' Format(Now.Date, "yyyy-MM-dd") & Now.ToLongTimeString.Replace(":", "")
                FILedatetiem = dtDateTime
                dtJEDate = oRecItem.Fields.Item("U_Z_PayDate").Value
                dtDateTime = dtJEDate.ToString("yyyy-MM-dd")
                strTRGFileName = strExportFilePaty & "\JE_" & FILedatetiem & ".trg"
                strFilename = strExportFilePaty & "\JE_" & FILedatetiem & ".xml"
                '(strFilename, System.Text.Encoding.UTF8)
                Dim st As String
                writer.WriteStartDocument(True)
                writer.Formatting = Formatting.Indented
                writer.Indentation = 2
                writer.WriteStartElement("IFX")
                writer.WriteStartElement("SignonRq")
                writer.WriteStartElement("RqUID")
                st = "Optimum-HRMS-" & dtJEDate.ToString("yyyy-MM-dd")
                writer.WriteString(st)
                writer.WriteEndElement() 'ReUID
                writer.WriteStartElement("ClientDt")
                'writer.WriteString(dtJEDate.ToString("yyyy-MM-dd"))
                writer.WriteString(aPostingDate.ToString("yyyy-MM-dd"))
                writer.WriteEndElement()
                writer.WriteStartElement("ClientApp")
                writer.WriteString("OPT")
                writer.WriteEndElement()
                writer.WriteEndElement() 'SignonRq

                writer.WriteStartElement("BankSvcRq")
                writer.WriteStartElement("FinancialMessageAddRq")
                writer.WriteStartElement("RqUID")
                st = "Optimum-HRMS-" & dtJEDate.ToString("yyyy-MM-dd")
                writer.WriteString(st)
                writer.WriteEndElement() 'RqUid
                writer.WriteStartElement("FinancialEntry")
                writer.WriteStartElement("CurCode")
                st = "BHD"
                writer.WriteString(st)
                writer.WriteEndElement() 'BHD

                writer.WriteStartElement("Amt")
                Dim dblAmount = oRecItem.Fields.Item("U_Z_NetSalary").Value
                dblAmount = Math.Round(dblAmount, 3)
                st = dblAmount.ToString ' oRecItem.Fields.Item("U_Z_NetSalary").Value
                writer.WriteString(st)
                writer.WriteEndElement() 'Amt

                writer.WriteStartElement("AcctId")
                st = "2000021985787"
                st = strEmpBankAccount
                writer.WriteString(st)
                writer.WriteEndElement() 'AcctId


                writer.WriteStartElement("EntryType")
                st = "Cr"

                writer.WriteString(st)
                writer.WriteEndElement() 'EntryType

                writer.WriteStartElement("EffDt")
                st = dtJEDate.ToString("yyyy-MM-dd")
                writer.WriteString(st)
                writer.WriteEndElement() 'EffDt
                writer.WriteStartElement("ApplType")
                'st = "Sv"
                If strEmpBankAccount.StartsWith("2") Then
                    st = "Sv"
                ElseIf strEmpBankAccount.StartsWith("1") Then
                    st = "CK"
                End If
                writer.WriteString(st)
                writer.WriteEndElement() 'ApppelTy
                writer.WriteStartElement("Memo")
                st = MonthName(dtJEDate.Month) & " Salary (OPTIMUM) " & Year(dtJEDate).ToString  ' dtJEDate.ToString("yyyy-MM-dd")
                writer.WriteString(st)
                writer.WriteEndElement() 'memo
                writer.WriteEndElement() 'FinancialEntry
                writer.WriteEndElement() 'FinancialMessageAddRq
                writer.WriteStartElement("RqUID")
                st = "Optimum-HRMS-" & dtJEDate.ToString("yyyy-MM-dd")
                writer.WriteString(st)
                writer.WriteEndElement() 'RqUid
                writer.WriteEndElement() 'BankSvcRq
                writer.WriteEndElement() 'IFX

                writer.Flush()
                MyXMLString = myStringWriter.ToString()

                myStringWriter.Close()
                writer.Close()
                ' SendXMLtoIFX(MyXMLString)

                Dim strPaht As String = getXMLPath()
                Dim blnFileExists As Boolean = False
                Dim oDoc1 As New XmlDocument
                If strPaht <> "" Then
                    If IO.Directory.Exists(strPaht) Then

                        oDoc1.LoadXml(MyXMLString)
                        blnFileExists = True
                        If IO.Directory.Exists(strPaht & "\Success") = False Then
                            IO.Directory.CreateDirectory(strPaht & "\Success")
                        End If
                        If IO.Directory.Exists(strPaht & "\Error") = False Then
                            IO.Directory.CreateDirectory(strPaht & "\Error")
                        End If
                        'oDoc1.Save(strPaht & "\JE_" & aPath & ".xml")
                    End If
                End If


                Dim IFXResponse As String = SendXMLtoIFX(MyXMLString)
                Dim doc As New XmlDocument
                doc.LoadXml(IFXResponse)
                Try
                    'Dim x As String = doc.SelectSingleNode("IFX/BankScvRs/RqUID)").InnerText.ToString
                    Dim locx, locy, locy1 As String
                    locx = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusCode").InnerText).ToString
                    locy = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/Severity").InnerText).ToString
                    locy1 = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusDesc").InnerText).ToString
                    If locx = "0" Then
                        ' oRecItem.DoQuery("Update OJDT set U_Export='Y',U_Z_IFXReply='" & locy1 & "' where TransId in (" & strJNo & ")")
                        oRecItem.DoQuery("Update [@Z_PAYROLL1] set U_Z_IFXPAY='C',U_Z_IFXEmpRes='" & locy1 & "' where Code='" & aPath & "' and isnull(U_Z_IFXPosting,'P')='C'")
                        '  strMessage = "Export Jounral Entry  Compleated . Journal Entry No : " & strJNo
                        ' oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        If blnFileExists Then
                            If File.Exists(strPaht & "\Success\JE_" & aPath & "_Success.xml") Then
                                File.Delete(strPaht & "\Success\JE_" & aPath & "_Success.xml")
                            End If
                            oDoc1.Save(strPaht & "\Success\JE_" & aPath & "_Success.xml")
                        End If
                        Return True
                    Else
                        strMessage = "Error while transfering Journal Entry No : " & strJNo & " : Error : " & locy1
                        oApplication.SBO_Application.MessageBox(strMessage)
                        oRecItem.DoQuery("Update [@Z_PAYROLL1] set U_Z_IFXPAY='P',U_Z_IFXEmpRes='" & locy1 & "' where Code='" & aPath & "' and isnull(U_Z_IFXPosting,'P')='C'")
                        If blnFileExists = True Then
                            If File.Exists(strPaht & "\Error\JE_" & aPath & "_Error.xml") Then
                                File.Delete(strPaht & "\Error\JE_" & aPath & "_Error.xml")
                            End If
                            oDoc1.Save(strPaht & "\Error\JE_" & aPath & "_Error.xml")
                        End If
                    End If
                    Return False
                    'WriteErrorlog(strMessage, strErrorFileName)
                    oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Catch ex As Exception
                End Try
            End If
            oMainRec.MoveNext()
        Next
        Return True
        'If File.Exists(strTRGFileName) Then
        '    File.Delete(strTRGFileName)
        'End If
    End Function


    Public Function ExportJournalEntries_Testing_Payroll_Empployee_Debit(ByVal aBranch As String, ByVal aDept As String, ByVal aMonth As String, ByVal aYear As String, ByVal aPostingdate As Date) As Boolean
        Dim oRecItem, oRecItemCode, oTemp, oMainRec As SAPbobsCOM.Recordset
        Dim strSQL, strFilename, strUserID As String
        Dim sValue As String
        Dim sPath, strLogDirectory, strPath, strMessage, strSelectedFolderPath, strExportFilePaty As String
        Dim blnErrorflag As Boolean
        Dim FILedatetiem As String
        strPath = System.Windows.Forms.Application.StartupPath
        strFilename = Now.ToLongDateString
        strPath = "" 'aPath
        ' strPath = aFileName
        oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMainRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strTransID As String
        Dim dtJEDate As Date
        For intLoop As Integer = 1 To 1 'oMainRec.RecordCount - 1
            strTransID = "" 'aCode
            ' dtJEDate = oMainRec.Fields.Item(1).Value
            dtJEDate = aPostingdate
            Dim strPhxId As String = ""
            oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strDebitAccount As String
            oRecItem.DoQuery("Select * from [@Z_PAY_OGLA]")
            strDebitAccount = oRecItem.Fields.Item("U_Z_SALCRE_ACC").Value
            'strSQL = "SELECT sum(T0.[U_Z_NetSalary]) 'U_Z_NetSalary',T0.[U_Z_Branch], T0.[U_Z_Dept],T0.[U_Z_PayDate] FROM [dbo].[@Z_PAYROLL1]  T0 where isnull(T0.U_Z_JVNo,'')<>'' and Isnull(U_Z_IFXPosting,'P')='C' and isnull(U_Z_IFXPAY,'P')='C' and T0.U_Z_Month=" & aMonth & " and T0.U_Z_Year=" & aYear & " and  T0.U_Z_Branch ='" & aBranch & "' and U_Z_Dept='" & aDept & "' group by T0.[U_Z_Branch], T0.[U_Z_Dept],T0.[U_Z_PayDate]"
            strSQL = "SELECT sum(T0.[U_Z_NetSalary]) 'U_Z_NetSalary',T0.[U_Z_Branch], T0.[U_Z_Dept],T0.[U_Z_PayDate] FROM [dbo].[@Z_PAYROLL1]  T0 where T0.U_Z_Month=" & aMonth & " and T0.U_Z_Year=" & aYear & " and  T0.U_Z_Branch ='" & aBranch & "' and U_Z_Dept='" & aDept & "' group by T0.[U_Z_Branch], T0.[U_Z_Dept],T0.[U_Z_PayDate]"
            oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItem.DoQuery(strSQL)
            Dim strTRGFileName As String
            Dim dtDateTime As String
            Dim strJNo As String = ""
            Dim strCheckNo As String = ""
            Dim MyXMLString As String
            Dim myStringWriter As New StringWriter
            Dim writer As New XmlTextWriter(myStringWriter)
            If 1 = 1 Then 'oRecItem.Fields.Item("U_Z_IFXPosting").Value = "C" Then
                '     Dim strEmpId As String = oRecItem.Fields.Item("U_Z_EmpId").Value
                FILedatetiem = Now.ToString("yyyyMMdd_hh_mm_ss")
                dtDateTime = dtJEDate.ToString("yyyy-MM-dd") & " _" & strTransID ' Format(Now.Date, "yyyy-MM-dd") & Now.ToLongTimeString.Replace(":", "")
                FILedatetiem = dtDateTime
                dtJEDate = oRecItem.Fields.Item("U_Z_PayDate").Value
                dtDateTime = dtJEDate.ToString("yyyy-MM-dd")
                strTRGFileName = strExportFilePaty & "\JE_" & FILedatetiem & ".trg"
                strFilename = strExportFilePaty & "\JE_" & FILedatetiem & ".xml"
                '(strFilename, System.Text.Encoding.UTF8)
                Dim st As String
                writer.WriteStartDocument(True)
                writer.Formatting = Formatting.Indented
                writer.Indentation = 2
                writer.WriteStartElement("IFX")
                writer.WriteStartElement("SignonRq")
                writer.WriteStartElement("RqUID")
                st = "Optimum-HRMS-" & dtJEDate.ToString("yyyy-MM-dd")
                writer.WriteString(st)
                writer.WriteEndElement() 'ReUID
                writer.WriteStartElement("ClientDt")
                '  writer.WriteString(dtJEDate.ToString("yyyy-MM-dd"))
                  writer.WriteString(aPostingdate .ToString("yyyy-MM-dd"))
                writer.WriteEndElement()
                writer.WriteStartElement("ClientApp")
                writer.WriteString("OPT")
                writer.WriteEndElement()
                writer.WriteEndElement() 'SignonRq

                writer.WriteStartElement("BankSvcRq")
                writer.WriteStartElement("FinancialMessageAddRq")
                writer.WriteStartElement("RqUID")
                st = "Optimum-HRMS-" & dtJEDate.ToString("yyyy-MM-dd")
                writer.WriteString(st)
                writer.WriteEndElement() 'RqUid
                writer.WriteStartElement("FinancialEntry")
                writer.WriteStartElement("CurCode")
                st = "BHD"
                writer.WriteString(st)
                writer.WriteEndElement() 'BHD

                writer.WriteStartElement("Amt")
                Dim dblAmount As Double = oRecItem.Fields.Item("U_Z_NetSalary").Value
                dblAmount = Math.Round(dblAmount, 3)
                st = dblAmount.ToString ' oRecItem.Fields.Item("U_Z_NetSalary").Value
                writer.WriteString(st)
                writer.WriteEndElement() 'Amt
                Dim straccount As String
                straccount = "01-01"
                writer.WriteStartElement("AcctId")
                aBranch = aBranch.Substring(aBranch.Length - 3, 3)
                aDept = aDept.Substring(aDept.Length - 3, 3)
                st = "01-01-" & aBranch & "-" & aDept & "-" & strDebitAccount
                writer.WriteString(st)
                writer.WriteEndElement() 'AcctId

                writer.WriteStartElement("AcctTye")
                st = "GL"
                writer.WriteString(st)
                writer.WriteEndElement() 'AcctId

                writer.WriteStartElement("EntryType")
                st = "dr"
                writer.WriteString(st)
                writer.WriteEndElement() 'EntryType

                writer.WriteStartElement("EffDt")
                st = dtJEDate.ToString("yyyy-MM-dd")
                writer.WriteString(st)
                writer.WriteEndElement() 'EffDt
                writer.WriteStartElement("ApplType")
                st = "GL"
                writer.WriteString(st)
                writer.WriteEndElement() 'ApppelTy
                writer.WriteStartElement("Memo")
                st = MonthName(dtJEDate.Month) & " Salary (OPTIMUM) " & Year(dtJEDate).ToString  ' dtJEDate.ToString("yyyy-MM-dd")
                writer.WriteString(st)
                writer.WriteEndElement() 'memo
                writer.WriteEndElement() 'FinancialEntry
                writer.WriteEndElement() 'FinancialMessageAddRq
                writer.WriteStartElement("RqUID")
                st = "Optimum-HRMS-" & dtJEDate.ToString("yyyy-MM-dd")
                writer.WriteString(st)
                writer.WriteEndElement() 'RqUid
                writer.WriteEndElement() 'BankSvcRq
                writer.WriteEndElement() 'IFX

                writer.Flush()
                MyXMLString = myStringWriter.ToString()

                myStringWriter.Close()
                writer.Close()
                ' SendXMLtoIFX(MyXMLString)
                Dim strPaht As String = getXMLPath()
                Dim blnFileExists As Boolean = False
                Dim oDoc1 As New XmlDocument
                If strPaht <> "" Then
                    If IO.Directory.Exists(strPaht) Then
                        oDoc1.LoadXml(MyXMLString)
                        blnFileExists = True
                        If IO.Directory.Exists(strPaht & "\Success") = False Then
                            IO.Directory.CreateDirectory(strPaht & "\Success")
                        End If
                        If IO.Directory.Exists(strPaht & "\Error") = False Then
                            IO.Directory.CreateDirectory(strPaht & "\Error")
                        End If
                        'oDoc1.Save(strPaht & "\JE_" & aPath & ".xml")
                    End If
                End If
                Dim IFXResponse As String = SendXMLtoIFX(MyXMLString)
                Dim doc As New XmlDocument
                doc.LoadXml(IFXResponse)
                Try
                    'Dim x As String = doc.SelectSingleNode("IFX/BankScvRs/RqUID)").InnerText.ToString
                    Dim locx, locy, locy1 As String
                    'locx = "0"
                    'locy = "Success"
                    'locy1 = "Success"
                    locx = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusCode").InnerText).ToString
                    locy = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/Severity").InnerText).ToString
                    locy1 = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusDesc").InnerText).ToString
                    If locx = "0" Then
                        ' oRecItem.DoQuery("Update OJDT set U_Export='Y',U_Z_IFXReply='" & locy1 & "' where TransId in (" & strJNo & ")")
                        '   oRecItem.DoQuery("Update [@Z_PAYROLL1] set U_Z_IFXPAY='C',U_Z_IFXEmpRes='" & locy1 & "' where Code='" & aPath & "' and isnull(U_Z_IFXPosting,'P')='C'")
                        '  strMessage = "Export Jounral Entry  Compleated . Journal Entry No : " & strJNo
                        ' oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        If blnFileExists Then


                            If File.Exists(strPaht & "\Success\JE_EmpDebit_Success.xml") Then
                                File.Delete(strPaht & "\Success\JE_EmpDebit_Success.xml")

                            End If
                            oDoc1.Save(strPaht & "\Success\JE_EmpDebit_Success.xml")

                        End If
                        Return True
                    Else
                        strMessage = "Error while transfering Journal Entry No : " & strJNo & " : Error : " & locy1
                        oApplication.SBO_Application.MessageBox(strMessage)
                        ' oRecItem.DoQuery("Update [@Z_PAYROLL1] set U_Z_IFXPAY='P',U_Z_IFXEmpRes='" & locy1 & "' where Code='" & aPath & "' and isnull(U_Z_IFXPosting,'P')='C'")
                        If blnFileExists = True Then
                            If File.Exists(strPaht & "\Error\JE_EmpDebit_Error.xml") Then
                                File.Delete(strPaht & "\Error\JE_EmpDebit_Error.xml")
                            End If
                            oDoc1.Save(strPaht & "\Error\JE_EmpDebit_Error.xml")
                        End If
                    End If
                    Return False
                    'WriteErrorlog(strMessage, strErrorFileName)
                    oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Catch ex As Exception
                End Try
            End If


            oMainRec.MoveNext()
        Next
        Return True
        'If File.Exists(strTRGFileName) Then
        '    File.Delete(strTRGFileName)
        'End If
    End Function

    Public Function GetJournalRemarks() As String
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Select * from [@Z_IFXSetup]")
        Return oTest.Fields.Item("U_Z_JERemarks").Value
    End Function

    Public Function getXMLPath() As String
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Select * from [@Z_IFXSetup]")
        Return oTest.Fields.Item("U_Z_XMLPath").Value
    End Function
    Public Function ExportJournalEntries_Testing_MultiCurrency(ByVal aPath As String) As Boolean
        Dim oRecItem, oRecItemCode, oTemp, oMainRec, OJEUpdate As SAPbobsCOM.Recordset
        Dim strSQL, strFilename, strUserID As String
        Dim sValue As String
        Dim sPath, strLogDirectory, strPath, strMessage, strSelectedFolderPath, strExportFilePaty As String
        Dim blnErrorflag As Boolean
        Dim FILedatetiem As String
        strPath = System.Windows.Forms.Application.StartupPath
        strFilename = Now.ToLongDateString
        strPath = aPath
        ' strPath = aFileName
        oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMainRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        OJEUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'aPath = 1

        strSQL = "Select T1.TransId ,T1.RefDate,Count(*) from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where Isnull(T1.U_Export,'N')='N' and T0.TransId=" & CInt(aPath) & " group by T1.TransId,T1.RefDate" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
        oMainRec.DoQuery(strSQL)
        Dim strTransID As String
        Dim dtJEDate As Date
        For intLoop As Integer = 0 To oMainRec.RecordCount - 1
            strTransID = oMainRec.Fields.Item(0).Value
            dtJEDate = oMainRec.Fields.Item(1).Value

            Dim strPhxId As String = ""
            'BP customer Master
            'strSQL = "Select * from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  where Isnull(T1.U_Export,'N')='N'" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"

            strSQL = "Select *,T0.Ref2 'OutGoing',isnull(T2.U_PhxId,'911') 'Uid' from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where T1.TransId=" & oMainRec.Fields.Item(0).Value & " and  Isnull(T1.U_Export,'N')='N'" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
            strSQL = "Select *,T0.Ref2 'OutGoing',T1.TransType 'TransType',T0.ShortName 'BPName',isnull(T2.U_PhxId,'911') 'Uid',T0.Ref3Line 'CheckNo' from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where  isnull(T0.VatGroup,'')<>'X1' and Isnull(T1.U_Export,'N')='N' and  T1.TransId=" & oMainRec.Fields.Item(0).Value  'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"


            oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItem.DoQuery(strSQL)
            FILedatetiem = Now.ToString("yyyyMMdd_hh_mm_ss")
            Dim transtype As String
            Dim strTRGFileName As String
            Dim dtDateTime As String
            Dim strJNo As String = ""
            Dim strCheckNo As String = ""
            dtDateTime = dtJEDate.ToString("yyyy-MM-dd") & " _" & strTransID ' Format(Now.Date, "yyyy-MM-dd") & Now.ToLongTimeString.Replace(":", "")
            FILedatetiem = dtDateTime
            dtDateTime = dtJEDate.ToString("yyyy-MM-dd")
            strTRGFileName = strExportFilePaty & "\JE_" & FILedatetiem & ".trg"
            strFilename = strExportFilePaty & "\JE_" & FILedatetiem & ".xml"
            Dim MyXMLString As String
            Dim myStringWriter As New StringWriter
            Dim writer As New XmlTextWriter(myStringWriter)
            '(strFilename, System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2
            writer.WriteStartElement("IFX")

            writer.WriteStartElement("SignonRq")
            writer.WriteStartElement("RqUID")
            writer.WriteString("6a8b5973-ec37-47e0-855e-4cc020ab7f02")
            writer.WriteEndElement()
            writer.WriteStartElement("ClientDt")
            writer.WriteString(dtJEDate.ToString("yyyy-MM-dd"))
            writer.WriteEndElement()
            writer.WriteStartElement("ClientApp")
            writer.WriteString("SAP")
            writer.WriteEndElement()
            writer.WriteStartElement("OperatorId")
            writer.WriteString("Manager")
            writer.WriteEndElement()
            writer.WriteEndElement()

            'BankSvcRq
            writer.WriteStartElement("BankSvcRq")
            writer.WriteStartElement("FinancialMessageAddRq")
            Dim strFCCurrency, strLocalCurrency, strCurrency, strType, strAccountType As String
            Dim dblAmount As Double
            oTemp.DoQuery("Select MainCurncy from OADM")
            strLocalCurrency = oTemp.Fields.Item(0).Value
            Dim strTransCurrency As String = ""

            For intRow As Integer = 0 To oRecItem.RecordCount - 1

                strPhxId = oRecItem.Fields.Item("U_PhxId").Value
                If strJNo = "" Then
                    strJNo = oRecItem.Fields.Item("TransId").Value
                Else
                    strJNo = strJNo & "," & oRecItem.Fields.Item("TransId").Value
                End If
                If oRecItem.Fields.Item("Account").Value <> oRecItem.Fields.Item("ShortName").Value Then
                    strAccountType = "GL"
                Else
                    strAccountType = "GL"
                End If
                strFCCurrency = oRecItem.Fields.Item("FCCurrency").Value
                If strFCCurrency <> "" Then
                    strTransCurrency = strFCCurrency
                    strCurrency = strFCCurrency
                    If oRecItem.Fields.Item("FCDebit").Value <> 0 Then
                        dblAmount = oRecItem.Fields.Item("FCDebit").Value
                        strType = "Dr"
                    Else
                        dblAmount = oRecItem.Fields.Item("FCCredit").Value
                        strType = "Cr"
                    End If
                Else
                    strTransCurrency = strLocalCurrency

                    strCurrency = strLocalCurrency
                    If oRecItem.Fields.Item("Debit").Value <> 0 Then
                        dblAmount = oRecItem.Fields.Item("Debit").Value
                        strType = "Dr"
                    Else
                        dblAmount = oRecItem.Fields.Item("Credit").Value
                        strType = "Cr"
                    End If
                End If
                writer.WriteStartElement("FinancialEntry")

                Dim straccount, sAPAccount, CostCenter1, CostCenter2 As String

                '  If oRecItem.Fields.Item("OutGoing").Value = "" Then
                If oRecItem.Fields.Item("TransType").Value <> "46" Then
                    writer.WriteStartElement("CurCode")
                    writer.WriteString(strCurrency)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Amt")
                    writer.WriteString(dblAmount.ToString)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Description")
                    Dim strRemk As String = GetJournalRemarks()
                    If strRemk <> "" Then
                        strRemk = strRemk & "-" & oRecItem.Fields.Item("LineMemo").Value
                    Else
                        strRemk = oRecItem.Fields.Item("LineMemo").Value
                    End If

                    writer.WriteString(strRemk)
                    '  writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                    writer.WriteEndElement()
                    straccount = "01-01"
                    If oRecItem.Fields.Item("TransType").Value <> "46" Then
                        sAPAccount = oRecItem.Fields.Item("Account").Value
                        CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                        CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                        If CostCenter1.Length > 0 Then
                            If CostCenter1.Length > 3 Then
                                CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                            Else
                                CostCenter1 = CostCenter1
                            End If
                        Else
                            CostCenter1 = ""
                        End If
                        If CostCenter2.Length > 0 Then
                            If CostCenter2.Length > 3 Then
                                CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                            Else
                                CostCenter2 = CostCenter2
                            End If
                        Else
                            CostCenter2 = ""
                        End If
                        straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                    Else
                        straccount = oRecItem.Fields.Item("OutGoing").Value
                    End If


                    writer.WriteStartElement("AcctId")
                    writer.WriteString(straccount)
                    writer.WriteEndElement()
                    If oRecItem.Fields.Item("TransType").Value = "46" Then
                        writer.WriteStartElement("ApplType")
                        writer.WriteString("SV")
                        writer.WriteEndElement()
                        writer.WriteStartElement("AcctType")
                        writer.WriteString("SAV")
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("AcctType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    End If

                    writer.WriteStartElement("EntryType")
                    writer.WriteString(strType)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EffDt")
                    writer.WriteString(dtDateTime)
                    writer.WriteEndElement()

                    writer.WriteStartElement("ApplType")
                    writer.WriteString(strAccountType)
                    writer.WriteEndElement()
                Else
                    writer.WriteStartElement("CurCode")
                    writer.WriteString(strCurrency)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Amt")
                    writer.WriteString(dblAmount.ToString)
                    writer.WriteEndElement()



                    straccount = "01-01"
                    If oRecItem.Fields.Item("TransType").Value <> "46" Then
                        sAPAccount = oRecItem.Fields.Item("Account").Value
                        CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                        CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                        If CostCenter1.Length > 0 Then
                            If CostCenter1.Length > 3 Then
                                CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                            Else
                                CostCenter1 = CostCenter1
                            End If
                        Else
                            CostCenter1 = ""
                        End If
                        If CostCenter2.Length > 0 Then
                            If CostCenter2.Length > 3 Then
                                CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                            Else
                                CostCenter2 = CostCenter2
                            End If
                        Else
                            CostCenter2 = ""
                        End If
                        straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                    Else

                        If strType = "Cr" Then
                            straccount = oRecItem.Fields.Item("OutGoing").Value
                            If straccount = "" Then
                                straccount = "01-01"
                                sAPAccount = oRecItem.Fields.Item("Account").Value
                                CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                                CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                                If CostCenter1.Length > 0 Then
                                    If CostCenter1.Length > 3 Then
                                        CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                    Else
                                        CostCenter1 = CostCenter1
                                    End If
                                Else
                                    CostCenter1 = ""
                                End If
                                If CostCenter2.Length > 0 Then
                                    If CostCenter2.Length > 3 Then
                                        CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                    Else
                                        CostCenter2 = CostCenter2
                                    End If
                                Else
                                    CostCenter2 = ""
                                End If
                                straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                            End If
                        Else
                            '  straccount = sAPAccount ' oRecItem.Fields.Item("OutGoing").Value
                            straccount = "01-01"
                            sAPAccount = oRecItem.Fields.Item("Account").Value
                            CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                            CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                            If CostCenter1.Length > 0 Then
                                If CostCenter1.Length > 3 Then
                                    CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                Else
                                    CostCenter1 = CostCenter1
                                End If
                            Else
                                CostCenter1 = ""
                            End If
                            If CostCenter2.Length > 0 Then
                                If CostCenter2.Length > 3 Then
                                    CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                Else
                                    CostCenter2 = CostCenter2
                                End If
                            Else
                                CostCenter2 = ""
                            End If
                            straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount

                        End If

                    End If


                    writer.WriteStartElement("AcctId")
                    writer.WriteString(straccount)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EntryType")
                    writer.WriteString(strType)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EffDt")
                    writer.WriteString(dtDateTime)
                    writer.WriteEndElement()

                    '  If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                    If oRecItem.Fields.Item("TransType").Value = "46" Then
                        If strType = "Dr" Then
                            writer.WriteStartElement("ApplType")
                            writer.WriteString("GL")
                            writer.WriteEndElement()
                            writer.WriteStartElement("AcctType")
                            writer.WriteString("GL")
                            writer.WriteEndElement()
                        Else
                            If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                                writer.WriteStartElement("ApplType")
                                If straccount.StartsWith("1") = True Then
                                    writer.WriteString("CK")
                                ElseIf straccount.StartsWith("2") = True Then
                                    writer.WriteString("SV")
                                Else
                                    writer.WriteString("CK")
                                End If

                                writer.WriteEndElement()
                                writer.WriteStartElement("AcctType")
                                Dim Test As SAPbobsCOM.Recordset
                                Test = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim sa As String
                                sa = "Select * from OCRB where  Account='" & straccount & "'"
                                Test.DoQuery(sa)
                                If Test.RecordCount > 0 Then
                                    If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                        writer.WriteString("MGC")
                                    Else
                                        writer.WriteString(Test.Fields.Item("MandateID").Value)
                                    End If
                                Else

                                    If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                        writer.WriteString("MGC")
                                    Else
                                        writer.WriteString(Test.Fields.Item("CUR").Value)
                                        writer.WriteString("CUR")
                                    End If
                                    'writer.WriteString("CUR")
                                End If



                                writer.WriteEndElement()

                            Else
                                writer.WriteStartElement("ApplType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                                writer.WriteStartElement("AcctType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                            End If

                            strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                            writer.WriteStartElement("ChkNum")
                            writer.WriteString(strCheckNo)
                            writer.WriteEndElement()
                        End If

                    Else
                        writer.WriteStartElement("AcctType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    End If

                    If oRecItem.Fields.Item("TransType").Value = "46" Then
                        Dim strs As SAPbobsCOM.Recordset
                        strs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        strs.DoQuery("Select * from OVPM where DocEntry=" & oRecItem.Fields.Item("BaseRef").Value)

                        If strs.Fields.Item("CheckSum").Value > 0 Then
                            strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                            If strType <> "Dr" Then
                                strCheckNo = "MANAGER CHEQUE NO:" & strCheckNo & " Branch:911"
                            Else
                                strCheckNo = "MGC:" & strCheckNo & " Br:911" & " Bnf :" & strs.Fields.Item("Comments").Value 'oRecItem.Fields.Item("LineMemo").Value
                            End If
                        Else
                            strCheckNo = ""
                        End If
                        writer.WriteStartElement("Description")
                        If strCheckNo = "" Then
                            writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                        Else
                            writer.WriteString(strCheckNo)
                        End If
                        '  
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("Description")
                        writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                        writer.WriteEndElement()
                    End If

                End If
                writer.WriteEndElement()

                'Foreign Currency XML
                strFCCurrency = oRecItem.Fields.Item("FCCurrency").Value
                If strFCCurrency <> "" Then
                    Dim strPostAccount As String = ""
                    strPhxId = oRecItem.Fields.Item("U_PhxId").Value
                    If strJNo = "" Then
                        strJNo = oRecItem.Fields.Item("TransId").Value
                    Else
                        strJNo = strJNo & "," & oRecItem.Fields.Item("TransId").Value
                    End If
                    If oRecItem.Fields.Item("Account").Value <> oRecItem.Fields.Item("ShortName").Value Then
                        strAccountType = "GL"
                    Else
                        strAccountType = "GL"
                    End If

                    If strFCCurrency <> "" Then
                        strTransCurrency = strFCCurrency
                        strCurrency = strFCCurrency
                        If oRecItem.Fields.Item("FCDebit").Value <> 0 Then
                            dblAmount = oRecItem.Fields.Item("FCDebit").Value
                            ' strType = "Dr"
                            strType = "Cr"
                        Else
                            dblAmount = oRecItem.Fields.Item("FCCredit").Value
                            '  strType = "Cr"
                            strType = "Dr"
                        End If
                    Else
                        strTransCurrency = strLocalCurrency
                        strCurrency = strLocalCurrency
                        If oRecItem.Fields.Item("Debit").Value <> 0 Then
                            dblAmount = oRecItem.Fields.Item("Debit").Value
                            '  strType = "Dr"
                            strType = "Cr"
                        Else
                            dblAmount = oRecItem.Fields.Item("Credit").Value
                            ' strType = "Cr"
                            strType = "Dr"
                        End If
                    End If
                    writer.WriteStartElement("FinancialEntry")
                    Dim Ors3 As SAPbobsCOM.Recordset
                    Ors3 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Ors3.DoQuery("Select isnull(U_Z_Acct,'') from OCRN where CurrCode='" & strTransCurrency & "'")
                    strPostAccount = Ors3.Fields.Item(0).Value
                    '  If oRecItem.Fields.Item("OutGoing").Value = "" Then
                    If oRecItem.Fields.Item("TransType").Value <> "46" Then
                        writer.WriteStartElement("CurCode")
                        writer.WriteString(strCurrency)
                        writer.WriteEndElement()

                        writer.WriteStartElement("Amt")
                        writer.WriteString(dblAmount.ToString)
                        writer.WriteEndElement()

                        writer.WriteStartElement("Description")
                        Dim strRemk As String = GetJournalRemarks()
                        If strRemk <> "" Then
                            strRemk = strRemk & "-" & oRecItem.Fields.Item("LineMemo").Value
                        Else
                            strRemk = oRecItem.Fields.Item("LineMemo").Value
                        End If

                        writer.WriteString(strRemk)
                        ' writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                        writer.WriteEndElement()
                        straccount = "01-01"
                        If oRecItem.Fields.Item("TransType").Value <> "46" Then
                            If strPostAccount <> "" Then
                                sAPAccount = strPostAccount
                            Else
                                sAPAccount = oRecItem.Fields.Item("Account").Value
                            End If
                            CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                            CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                            If CostCenter1.Length > 0 Then
                                If CostCenter1.Length > 3 Then
                                    CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                Else
                                    CostCenter1 = CostCenter1
                                End If
                            Else
                                CostCenter1 = ""
                            End If
                            If CostCenter2.Length > 0 Then
                                If CostCenter2.Length > 3 Then
                                    CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                Else
                                    CostCenter2 = CostCenter2

                                End If
                            Else
                                CostCenter2 = ""
                            End If
                            CostCenter2 = "001"
                            straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                        Else
                            straccount = oRecItem.Fields.Item("OutGoing").Value
                        End If


                        writer.WriteStartElement("AcctId")
                        writer.WriteString(straccount)
                        writer.WriteEndElement()
                        If oRecItem.Fields.Item("TransType").Value = "46" Then
                            writer.WriteStartElement("ApplType")
                            writer.WriteString("SV")
                            writer.WriteEndElement()
                            writer.WriteStartElement("AcctType")
                            writer.WriteString("SAV")
                            writer.WriteEndElement()
                        Else
                            writer.WriteStartElement("AcctType")
                            writer.WriteString(strAccountType)
                            writer.WriteEndElement()
                        End If

                        writer.WriteStartElement("EntryType")
                        writer.WriteString(strType)
                        writer.WriteEndElement()

                        writer.WriteStartElement("EffDt")
                        writer.WriteString(dtDateTime)
                        writer.WriteEndElement()

                        writer.WriteStartElement("ApplType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("CurCode")
                        writer.WriteString(strCurrency)
                        writer.WriteEndElement()

                        writer.WriteStartElement("Amt")
                        writer.WriteString(dblAmount.ToString)
                        writer.WriteEndElement()
                        straccount = "01-01"
                        If oRecItem.Fields.Item("TransType").Value <> "46" Then
                            'sAPAccount = oRecItem.Fields.Item("Account").Value
                            If strPostAccount <> "" Then
                                sAPAccount = strPostAccount
                            Else
                                sAPAccount = oRecItem.Fields.Item("Account").Value
                            End If
                            CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                            CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                            If CostCenter1.Length > 0 Then
                                If CostCenter1.Length > 3 Then
                                    CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                Else
                                    CostCenter1 = CostCenter1
                                End If
                            Else
                                CostCenter1 = ""
                            End If
                            If CostCenter2.Length > 0 Then
                                If CostCenter2.Length > 3 Then
                                    CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                Else
                                    CostCenter2 = CostCenter2
                                End If
                            Else
                                CostCenter2 = ""
                            End If
                            CostCenter2 = "001"
                            straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                        Else

                            If strType = "Dr" Then
                                straccount = oRecItem.Fields.Item("OutGoing").Value
                                If straccount = "" Then
                                    straccount = "01-01"
                                    sAPAccount = oRecItem.Fields.Item("Account").Value
                                    'If strPostAccount <> "" Then
                                    '    sAPAccount = strPostAccount
                                    'Else
                                    '    sAPAccount = oRecItem.Fields.Item("Account").Value
                                    'End If
                                    CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                                    CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                                    If CostCenter1.Length > 0 Then
                                        If CostCenter1.Length > 3 Then
                                            CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                        Else
                                            CostCenter1 = CostCenter1
                                        End If
                                    Else
                                        CostCenter1 = ""
                                    End If
                                    If CostCenter2.Length > 0 Then
                                        If CostCenter2.Length > 3 Then
                                            CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                        Else
                                            CostCenter2 = CostCenter2
                                        End If
                                    Else
                                        CostCenter2 = ""
                                    End If
                                    CostCenter2 = "001"
                                    straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                                End If
                            Else
                                '  straccount = sAPAccount ' oRecItem.Fields.Item("OutGoing").Value
                                straccount = "01-01"
                                ' sAPAccount = oRecItem.Fields.Item("Account").Value
                                If strPostAccount <> "" Then
                                    sAPAccount = strPostAccount
                                Else
                                    sAPAccount = oRecItem.Fields.Item("Account").Value
                                End If
                                CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                                CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                                If CostCenter1.Length > 0 Then
                                    If CostCenter1.Length > 3 Then
                                        CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                    Else
                                        CostCenter1 = CostCenter1
                                    End If
                                Else
                                    CostCenter1 = ""
                                End If
                                If CostCenter2.Length > 0 Then
                                    If CostCenter2.Length > 3 Then
                                        CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                    Else
                                        CostCenter2 = CostCenter2
                                    End If
                                Else
                                    CostCenter2 = ""
                                End If
                                CostCenter2 = "001"
                                straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                            End If

                        End If
                   

                        writer.WriteStartElement("AcctId")
                        writer.WriteString(straccount)
                        writer.WriteEndElement()

                        writer.WriteStartElement("EntryType")
                        writer.WriteString(strType)
                        writer.WriteEndElement()

                        writer.WriteStartElement("EffDt")
                        writer.WriteString(dtDateTime)
                        writer.WriteEndElement()

                        '  If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                        If oRecItem.Fields.Item("TransType").Value = "46" Then
                            If strType = "Dr" Then
                                writer.WriteStartElement("ApplType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                                writer.WriteStartElement("AcctType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                            Else
                                If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                                    writer.WriteStartElement("ApplType")
                                    If straccount.StartsWith("1") = True Then
                                        writer.WriteString("CK")
                                    ElseIf straccount.StartsWith("2") = True Then
                                        writer.WriteString("SV")
                                    Else
                                        writer.WriteString("CK")
                                    End If

                                    writer.WriteEndElement()
                                    writer.WriteStartElement("AcctType")
                                    Dim Test As SAPbobsCOM.Recordset
                                    Test = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim sa As String
                                    sa = "Select * from OCRB where  Account='" & straccount & "'"
                                    Test.DoQuery(sa)
                                    If Test.RecordCount > 0 Then
                                        If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                            writer.WriteString("MGC")
                                        Else
                                            writer.WriteString(Test.Fields.Item("MandateID").Value)
                                        End If
                                    Else

                                        If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                            writer.WriteString("MGC")
                                        Else
                                            writer.WriteString(Test.Fields.Item("CUR").Value)
                                            writer.WriteString("CUR")
                                        End If
                                        'writer.WriteString("CUR")
                                    End If



                                    writer.WriteEndElement()

                                Else
                                    writer.WriteStartElement("ApplType")
                                    writer.WriteString("GL")
                                    writer.WriteEndElement()
                                    writer.WriteStartElement("AcctType")
                                    writer.WriteString("GL")
                                    writer.WriteEndElement()
                                End If

                                strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                                writer.WriteStartElement("ChkNum")
                                writer.WriteString(strCheckNo)
                                writer.WriteEndElement()
                            End If

                        Else
                            writer.WriteStartElement("AcctType")
                            writer.WriteString(strAccountType)
                            writer.WriteEndElement()
                        End If

                        If oRecItem.Fields.Item("TransType").Value = "46" Then
                            Dim strs As SAPbobsCOM.Recordset
                            strs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            strs.DoQuery("Select * from OVPM where DocEntry=" & oRecItem.Fields.Item("BaseRef").Value)

                            If strs.Fields.Item("CheckSum").Value > 0 Then
                                strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                                If strType <> "Dr" Then
                                    strCheckNo = "MANAGER CHEQUE NO:" & strCheckNo & " Branch:911"
                                Else
                                    strCheckNo = "MGC:" & strCheckNo & " Br:911" & " Bnf :" & strs.Fields.Item("Comments").Value 'oRecItem.Fields.Item("LineMemo").Value
                                End If
                            Else
                                strCheckNo = ""
                            End If
                            writer.WriteStartElement("Description")
                            If strCheckNo = "" Then
                                writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                            Else
                                writer.WriteString(strCheckNo)
                            End If
                            '  
                            writer.WriteEndElement()
                        Else
                            writer.WriteStartElement("Description")
                            writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                            writer.WriteEndElement()
                        End If
                    End If
                    writer.WriteEndElement()

                    'update JE Lines UDF
                    If strType = "Dr" Then
                        OJEUpdate.DoQuery("Update JDT1 set U_Z_FCDRAc='" & sAPAccount & "' , U_Z_FCDRAm='" & dblAmount & "',U_Z_FCDRCUR='" & strFCCurrency & "' where TransId=" & oRecItem.Fields.Item("TransId").Value & " and Line_ID=" & oRecItem.Fields.Item("Line_ID").Value)
                    Else
                        OJEUpdate.DoQuery("Update JDT1 set U_Z_FCCRAc='" & sAPAccount & "' , U_Z_FCCRAmt='" & dblAmount & "',U_Z_FCCRCUR='" & strFCCurrency & "' where TransId=" & oRecItem.Fields.Item("TransId").Value & " and Line_ID=" & oRecItem.Fields.Item("Line_ID").Value)
                    End If

                    'Post Local Currency Account
                    strPhxId = oRecItem.Fields.Item("U_PhxId").Value
                    If strJNo = "" Then
                        strJNo = oRecItem.Fields.Item("TransId").Value
                    Else
                        strJNo = strJNo & "," & oRecItem.Fields.Item("TransId").Value
                    End If
                    If oRecItem.Fields.Item("Account").Value <> oRecItem.Fields.Item("ShortName").Value Then
                        strAccountType = "GL"
                    Else
                        strAccountType = "GL"
                    End If
                    strFCCurrency = oRecItem.Fields.Item("FCCurrency").Value
                    If strFCCurrency <> "" Then
                        strTransCurrency = strFCCurrency
                        strCurrency = strFCCurrency
                        If oRecItem.Fields.Item("FCDebit").Value <> 0 Then
                            dblAmount = oRecItem.Fields.Item("SysDeb").Value
                            strType = "Dr"
                            'strType = "Cr"
                        Else
                            dblAmount = oRecItem.Fields.Item("SysCred").Value
                            strType = "Cr"
                            'strType = "Dr"
                        End If
                    Else
                        strTransCurrency = strLocalCurrency
                        strCurrency = strLocalCurrency
                        If oRecItem.Fields.Item("Debit").Value <> 0 Then
                            dblAmount = oRecItem.Fields.Item("Debit").Value
                            strType = "Dr"
                            'strType = "Cr"
                        Else
                            dblAmount = oRecItem.Fields.Item("Credit").Value
                            strType = "Cr"
                            'strType = "Dr"
                        End If
                    End If
                    writer.WriteStartElement("FinancialEntry")
                    Ors3 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Ors3.DoQuery("Select MainCurncy from OADM")
                    strCurrency = Ors3.Fields.Item(0).Value

                    Ors3.DoQuery("Select isnull(U_Z_Acct1,'') from OCRN where CurrCode='" & strTransCurrency & "'")
                    strPostAccount = Ors3.Fields.Item(0).Value

                    '  If oRecItem.Fields.Item("OutGoing").Value = "" Then
                    If oRecItem.Fields.Item("TransType").Value <> "46" Then
                        writer.WriteStartElement("CurCode")
                        writer.WriteString(strCurrency)
                        writer.WriteEndElement()

                        writer.WriteStartElement("Amt")
                        writer.WriteString(dblAmount.ToString)
                        writer.WriteEndElement()

                        writer.WriteStartElement("Description")
                        Dim strRemk As String = GetJournalRemarks()
                        If strRemk <> "" Then
                            strRemk = strRemk & "-" & oRecItem.Fields.Item("LineMemo").Value
                        Else
                            strRemk = oRecItem.Fields.Item("LineMemo").Value
                        End If

                        writer.WriteString(strRemk)
                        ' writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                        writer.WriteEndElement()
                        straccount = "01-01"
                        If oRecItem.Fields.Item("TransType").Value <> "46" Then
                            If strPostAccount <> "" Then
                                sAPAccount = strPostAccount
                            Else
                                sAPAccount = oRecItem.Fields.Item("Account").Value
                            End If

                            CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                            CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                            If CostCenter1.Length > 0 Then
                                If CostCenter1.Length > 3 Then
                                    CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                Else
                                    CostCenter1 = CostCenter1
                                End If
                            Else
                                CostCenter1 = ""
                            End If
                            If CostCenter2.Length > 0 Then
                                If CostCenter2.Length > 3 Then
                                    CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                Else
                                    CostCenter2 = CostCenter2
                                End If
                            Else
                                CostCenter2 = ""
                            End If
                            CostCenter2 = "001"
                            straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                        Else
                            straccount = oRecItem.Fields.Item("OutGoing").Value
                        End If


                        writer.WriteStartElement("AcctId")
                        writer.WriteString(straccount)
                        writer.WriteEndElement()
                        If oRecItem.Fields.Item("TransType").Value = "46" Then
                            writer.WriteStartElement("ApplType")
                            writer.WriteString("SV")
                            writer.WriteEndElement()
                            writer.WriteStartElement("AcctType")
                            writer.WriteString("SAV")
                            writer.WriteEndElement()
                        Else
                            writer.WriteStartElement("AcctType")
                            writer.WriteString(strAccountType)
                            writer.WriteEndElement()
                        End If

                        writer.WriteStartElement("EntryType")
                        writer.WriteString(strType)
                        writer.WriteEndElement()

                        writer.WriteStartElement("EffDt")
                        writer.WriteString(dtDateTime)
                        writer.WriteEndElement()

                        writer.WriteStartElement("ApplType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("CurCode")
                        writer.WriteString(strCurrency)
                        writer.WriteEndElement()

                        writer.WriteStartElement("Amt")
                        writer.WriteString(dblAmount.ToString)
                        writer.WriteEndElement()
                        straccount = "01-01"
                        If oRecItem.Fields.Item("TransType").Value <> "46" Then
                            'sAPAccount = oRecItem.Fields.Item("Account").Value
                            If strPostAccount <> "" Then
                                sAPAccount = strPostAccount
                            Else
                                sAPAccount = oRecItem.Fields.Item("Account").Value
                            End If
                            CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                            CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                            If CostCenter1.Length > 0 Then
                                If CostCenter1.Length > 3 Then
                                    CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                Else
                                    CostCenter1 = CostCenter1
                                End If
                            Else
                                CostCenter1 = ""
                            End If
                            If CostCenter2.Length > 0 Then
                                If CostCenter2.Length > 3 Then
                                    CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                Else
                                    CostCenter2 = CostCenter2
                                End If
                            Else
                                CostCenter2 = ""
                            End If
                            CostCenter2 = "001"
                            straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                        Else

                            If strType = "Dr" Then
                                straccount = oRecItem.Fields.Item("OutGoing").Value
                                If straccount = "" Then
                                    straccount = "01-01"
                                    sAPAccount = oRecItem.Fields.Item("Account").Value
                                    'If strPostAccount <> "" Then
                                    '    sAPAccount = strPostAccount
                                    'Else
                                    '    sAPAccount = oRecItem.Fields.Item("Account").Value
                                    'End If
                                    CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                                    CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                                    If CostCenter1.Length > 0 Then
                                        If CostCenter1.Length > 3 Then
                                            CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                        Else
                                            CostCenter1 = CostCenter1
                                        End If
                                    Else
                                        CostCenter1 = ""
                                    End If
                                    If CostCenter2.Length > 0 Then
                                        If CostCenter2.Length > 3 Then
                                            CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                        Else
                                            CostCenter2 = CostCenter2
                                        End If
                                    Else
                                        CostCenter2 = ""
                                    End If
                                    CostCenter2 = "001"
                                    straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                                End If
                            Else
                                '  straccount = sAPAccount ' oRecItem.Fields.Item("OutGoing").Value
                                straccount = "01-01"
                                ' sAPAccount = oRecItem.Fields.Item("Account").Value
                                If strPostAccount <> "" Then
                                    sAPAccount = strPostAccount
                                Else
                                    sAPAccount = oRecItem.Fields.Item("Account").Value
                                End If
                                CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                                CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                                If CostCenter1.Length > 0 Then
                                    If CostCenter1.Length > 3 Then
                                        CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                    Else
                                        CostCenter1 = CostCenter1
                                    End If
                                Else
                                    CostCenter1 = ""
                                End If
                                If CostCenter2.Length > 0 Then
                                    If CostCenter2.Length > 3 Then
                                        CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                    Else
                                        CostCenter2 = CostCenter2
                                    End If
                                Else
                                    CostCenter2 = ""
                                End If
                                CostCenter2 = "001"
                                straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount

                            End If

                        End If

                      
                        writer.WriteStartElement("AcctId")
                        writer.WriteString(straccount)
                        writer.WriteEndElement()

                        writer.WriteStartElement("EntryType")
                        writer.WriteString(strType)
                        writer.WriteEndElement()

                        writer.WriteStartElement("EffDt")
                        writer.WriteString(dtDateTime)
                        writer.WriteEndElement()

                        '  If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                        If oRecItem.Fields.Item("TransType").Value = "46" Then
                            If strType = "Dr" Then
                                writer.WriteStartElement("ApplType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                                writer.WriteStartElement("AcctType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                            Else
                                If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                                    writer.WriteStartElement("ApplType")
                                    If straccount.StartsWith("1") = True Then
                                        writer.WriteString("CK")
                                    ElseIf straccount.StartsWith("2") = True Then
                                        writer.WriteString("SV")
                                    Else
                                        writer.WriteString("CK")
                                    End If

                                    writer.WriteEndElement()
                                    writer.WriteStartElement("AcctType")
                                    Dim Test As SAPbobsCOM.Recordset
                                    Test = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim sa As String
                                    sa = "Select * from OCRB where  Account='" & straccount & "'"
                                    Test.DoQuery(sa)
                                    If Test.RecordCount > 0 Then
                                        If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                            writer.WriteString("MGC")
                                        Else
                                            writer.WriteString(Test.Fields.Item("MandateID").Value)
                                        End If
                                    Else

                                        If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                            writer.WriteString("MGC")
                                        Else
                                            writer.WriteString(Test.Fields.Item("CUR").Value)
                                            writer.WriteString("CUR")
                                        End If
                                        'writer.WriteString("CUR")
                                    End If



                                    writer.WriteEndElement()

                                Else
                                    writer.WriteStartElement("ApplType")
                                    writer.WriteString("GL")
                                    writer.WriteEndElement()
                                    writer.WriteStartElement("AcctType")
                                    writer.WriteString("GL")
                                    writer.WriteEndElement()
                                End If

                                strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                                writer.WriteStartElement("ChkNum")
                                writer.WriteString(strCheckNo)
                                writer.WriteEndElement()
                            End If

                        Else
                            writer.WriteStartElement("AcctType")
                            writer.WriteString(strAccountType)
                            writer.WriteEndElement()
                        End If

                        If oRecItem.Fields.Item("TransType").Value = "46" Then
                            Dim strs As SAPbobsCOM.Recordset
                            strs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            strs.DoQuery("Select * from OVPM where DocEntry=" & oRecItem.Fields.Item("BaseRef").Value)

                            If strs.Fields.Item("CheckSum").Value > 0 Then
                                strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                                If strType <> "Dr" Then
                                    strCheckNo = "MANAGER CHEQUE NO:" & strCheckNo & " Branch:911"
                                Else
                                    strCheckNo = "MGC:" & strCheckNo & " Br:911" & " Bnf :" & strs.Fields.Item("Comments").Value 'oRecItem.Fields.Item("LineMemo").Value
                                End If
                            Else
                                strCheckNo = ""
                            End If
                            writer.WriteStartElement("Description")
                            If strCheckNo = "" Then
                                writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                            Else
                                writer.WriteString(strCheckNo)
                            End If
                            '  
                            writer.WriteEndElement()
                        Else
                            writer.WriteStartElement("Description")
                            writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                            writer.WriteEndElement()
                        End If

                    End If
                    writer.WriteEndElement()
                    'update JE Lines UDF
                    If strType = "Dr" Then
                        OJEUpdate.DoQuery("Update JDT1 set U_Z_FCDRAc='" & sAPAccount & "' , U_Z_FCDRAm='" & dblAmount & "',U_Z_FCDRCUR='" & strCurrency & "' where TransId=" & oRecItem.Fields.Item("TransId").Value & " and Line_ID=" & oRecItem.Fields.Item("Line_ID").Value)
                    Else
                        OJEUpdate.DoQuery("Update JDT1 set U_Z_FCCRAc='" & sAPAccount & "' , U_Z_FCCRAmt='" & dblAmount & "',U_Z_FCCRCUR='" & strCurrency & "' where TransId=" & oRecItem.Fields.Item("TransId").Value & " and Line_ID=" & oRecItem.Fields.Item("Line_ID").Value)

                    End If


                End If
                'End Foreign Currency XML
                oRecItem.MoveNext()
            Next
            ' writer.WriteString(dtDateTime)
            writer.WriteStartElement("BankInfo")
            writer.WriteStartElement("BranchId")
            writer.WriteString("911")
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("EmployeeIdent")
            writer.WriteStartElement("EmployeeIdentlNum")
            writer.WriteString(strPhxId)
            writer.WriteEndElement()
            writer.WriteStartElement("SuperEmployeeIdentlNum")
            writer.WriteString("0")
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("RqUID")
            writer.WriteString("85fb650d-336e-4c1d-bea3-6d9551b9fb08")
            writer.WriteEndElement()

            writer.WriteEndElement()

            writer.WriteStartElement("RqUID")
            writer.WriteString("85fb650d-336e-4c1d-bea3-6d9551b9fb08")
            writer.WriteEndElement()


            writer.WriteEndElement()
            writer.WriteEndElement()
            'writer.WriteEndElement()

            writer.Flush()
            MyXMLString = myStringWriter.ToString()
            myStringWriter.Close()
            writer.Close()

            ' SendXMLtoIFX(MyXMLString)
            Dim IFXResponse As String = SendXMLtoIFX(MyXMLString)
            Dim doc As New XmlDocument
            doc.LoadXml(IFXResponse)
            Try
                'Dim x As String = doc.SelectSingleNode("IFX/BankScvRs/RqUID)").InnerText.ToString
                Dim locx, locy, locy1 As String
                locx = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusCode").InnerText).ToString
                locy = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/Severity").InnerText).ToString
                locy1 = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusDesc").InnerText).ToString
                If locx = "0" Then
                    oRecItem.DoQuery("Update OJDT set U_Export='Y',U_Z_IFXReply='" & locy1 & "' where TransId in (" & strJNo & ")")
                    strMessage = "Export Jounral Entry  Compleated . Journal Entry No : " & strJNo
                    oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Return True
                Else
                    '     strMessage = "Error while transfering Journal Entry No : " & strJNo & " : Error : " & locy1
                    strMessage = "Error while transfering Journal Entry No : " & strJNo & " : Error : " & locy1 & " . The transaction has been reversed. Create a new Transaction"

                    oApplication.SBO_Application.MessageBox(strMessage)
                    oRecItem.DoQuery("Update OJDT set U_Z_IFXReply='" & locy1 & "' where TransId in (" & strJNo & ")")
                    Return False
                End If
                'WriteErrorlog(strMessage, strErrorFileName)
                oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception
            End Try
            oMainRec.MoveNext()
        Next
        Return True
        'If File.Exists(strTRGFileName) Then
        '    File.Delete(strTRGFileName)
        'End If
    End Function

    Public Sub ExportJournalEntries_Testing_Reverse(ByVal aPath As String)
        Dim oRecItem, oRecItemCode, oTemp, oMainRec As SAPbobsCOM.Recordset
        Dim strSQL, strFilename, strUserID As String
        Dim sValue As String
        Dim sPath, strLogDirectory, strPath, strMessage, strSelectedFolderPath, strExportFilePaty As String
        Dim blnErrorflag As Boolean
        Dim FILedatetiem As String
        strPath = System.Windows.Forms.Application.StartupPath
        strFilename = Now.ToLongDateString
        strPath = aPath
        ' strPath = aFileName
        oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMainRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'aPath = 1

        strSQL = "Update OJDT set U_Export='N' ,U_Z_IFXReply='' where Transid=" & CInt(aPath)
        oMainRec.DoQuery(strSQL)

        strSQL = "Select T1.TransId ,T1.RefDate,Count(*) from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where Isnull(T1.U_Export,'N')='N' and T0.TransId=" & CInt(aPath) & " group by T1.TransId,T1.RefDate" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
        oMainRec.DoQuery(strSQL)
        Dim strTransID As String
        Dim dtJEDate As Date
        For intLoop As Integer = 0 To oMainRec.RecordCount - 1
            strTransID = oMainRec.Fields.Item(0).Value
            dtJEDate = oMainRec.Fields.Item(1).Value

            Dim strPhxId As String = ""
            'BP customer Master
            'strSQL = "Select * from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  where Isnull(T1.U_Export,'N')='N'" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"

            strSQL = "Select *,T0.Ref2 'OutGoing',isnull(T2.U_PhxId,'911') 'Uid' from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where T1.TransId=" & oMainRec.Fields.Item(0).Value & " and  Isnull(T1.U_Export,'N')='N'" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
            strSQL = "Select *,T0.Ref2 'OutGoing',T1.TransType 'TransType',T0.ShortName 'BPName',isnull(T2.U_PhxId,'911') 'Uid',T0.Ref3Line 'CheckNo',T0.BaseRef 'BaseDoc' from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where  isnull(T0.VatGroup,'')<>'X1' and Isnull(T1.U_Export,'N')='N' and  T1.TransId=" & oMainRec.Fields.Item(0).Value  'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"


            oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecItem.DoQuery(strSQL)
            FILedatetiem = Now.ToString("yyyyMMdd_hh_mm_ss")
            Dim transtype As String
            Dim strTRGFileName As String
            Dim dtDateTime As String
            Dim strJNo As String = ""
            Dim strCheckNo As String = ""
            dtDateTime = dtJEDate.ToString("yyyy-MM-dd") & " _" & strTransID ' Format(Now.Date, "yyyy-MM-dd") & Now.ToLongTimeString.Replace(":", "")
            FILedatetiem = dtDateTime
            dtDateTime = dtJEDate.ToString("yyyy-MM-dd")
            strTRGFileName = strExportFilePaty & "\JE_" & FILedatetiem & ".trg"
            strFilename = strExportFilePaty & "\JE_" & FILedatetiem & ".xml"
            Dim MyXMLString As String
            Dim myStringWriter As New StringWriter
            Dim writer As New XmlTextWriter(myStringWriter)
            '(strFilename, System.Text.Encoding.UTF8)
            writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2
            writer.WriteStartElement("IFX")

            writer.WriteStartElement("SignonRq")
            writer.WriteStartElement("RqUID")
            writer.WriteString("6a8b5973-ec37-47e0-855e-4cc020ab7f02")
            writer.WriteEndElement()
            writer.WriteStartElement("ClientDt")
            writer.WriteString(dtJEDate.ToString("yyyy-MM-dd"))
            writer.WriteEndElement()
            writer.WriteStartElement("ClientApp")
            writer.WriteString("SAP")
            writer.WriteEndElement()
            writer.WriteStartElement("OperatorId")
            writer.WriteString("Manager")
            writer.WriteEndElement()
            writer.WriteEndElement()

            'BankSvcRq
            writer.WriteStartElement("BankSvcRq")
            writer.WriteStartElement("FinancialMessageAddRq")
            Dim strFCCurrency, strLocalCurrency, strCurrency, strType, strAccountType As String
            Dim dblAmount As Double
            oTemp.DoQuery("Select MainCurncy from OADM")
            strLocalCurrency = oTemp.Fields.Item(0).Value

            For intRow As Integer = 0 To oRecItem.RecordCount - 1

                strPhxId = oRecItem.Fields.Item("U_PhxId").Value
                If strJNo = "" Then
                    strJNo = oRecItem.Fields.Item("TransId").Value
                Else
                    strJNo = strJNo & "," & oRecItem.Fields.Item("TransId").Value
                End If
                If oRecItem.Fields.Item("Account").Value <> oRecItem.Fields.Item("ShortName").Value Then
                    strAccountType = "GL"
                Else
                    strAccountType = "GL"
                End If
                strFCCurrency = oRecItem.Fields.Item("FCCurrency").Value
                If strFCCurrency <> "" Then
                    strCurrency = strFCCurrency
                    If oRecItem.Fields.Item("FCDebit").Value <> 0 Then
                        dblAmount = oRecItem.Fields.Item("FCDebit").Value
                        strType = "Dr"
                    Else
                        dblAmount = oRecItem.Fields.Item("FCCredit").Value
                        strType = "Cr"
                    End If
                Else
                    strCurrency = strLocalCurrency
                    If oRecItem.Fields.Item("Debit").Value <> 0 Then
                        dblAmount = oRecItem.Fields.Item("Debit").Value
                        strType = "Dr"
                    Else
                        dblAmount = oRecItem.Fields.Item("Credit").Value
                        strType = "Cr"
                    End If
                End If
                writer.WriteStartElement("FinancialEntry")

                Dim straccount, sAPAccount, CostCenter1, CostCenter2 As String

                '                If oRecItem.Fields.Item("OutGoing").Value = "" Then
                If oRecItem.Fields.Item("TransType").Value <> "46" Then
                    writer.WriteStartElement("CurCode")
                    writer.WriteString(strCurrency)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Amt")
                    writer.WriteString(dblAmount.ToString)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Description")
                    Dim strRemk As String = GetJournalRemarks()
                    If strRemk <> "" Then
                        strRemk = strRemk & "-" & oRecItem.Fields.Item("LineMemo").Value
                    Else
                        strRemk = oRecItem.Fields.Item("LineMemo").Value
                    End If

                    writer.WriteString(strRemk)
                    'writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                    writer.WriteEndElement()


                    straccount = "01-01"
                    'If oRecItem.Fields.Item("OutGoing").Value = "" Then
                    If oRecItem.Fields.Item("TransType").Value <> "46" Then
                        sAPAccount = oRecItem.Fields.Item("Account").Value
                        CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                        CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                        If CostCenter1.Length > 0 Then
                            If CostCenter1.Length > 3 Then
                                CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                            Else
                                CostCenter1 = CostCenter1
                            End If
                        Else
                            CostCenter1 = ""
                        End If
                        If CostCenter2.Length > 0 Then
                            If CostCenter2.Length > 3 Then
                                CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                            Else
                                CostCenter2 = CostCenter2
                            End If
                        Else
                            CostCenter2 = ""
                        End If
                        straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                    Else
                        straccount = oRecItem.Fields.Item("OutGoing").Value
                    End If


                    writer.WriteStartElement("AcctId")
                    '   writer.WriteString(oRecItem.Fields.Item("Account").Value)
                    writer.WriteString(straccount)
                    writer.WriteEndElement()

                    '  If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                    If oRecItem.Fields.Item("TransType").Value = "46" Then
                        writer.WriteStartElement("ApplType")
                        writer.WriteString("SV")
                        writer.WriteEndElement()
                        writer.WriteStartElement("AcctType")
                        writer.WriteString("SAV")
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("AcctType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    End If

                    writer.WriteStartElement("EntryType")
                    writer.WriteString(strType)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EffDt")
                    writer.WriteString(dtDateTime)
                    writer.WriteEndElement()

                    writer.WriteStartElement("ApplType")
                    writer.WriteString(strAccountType)
                    writer.WriteEndElement()
                Else 'OutGoing Payment Reversal
                    writer.WriteStartElement("CurCode")
                    writer.WriteString(strCurrency)
                    writer.WriteEndElement()

                    writer.WriteStartElement("Amt")
                    writer.WriteString(dblAmount.ToString)
                    writer.WriteEndElement()

                    straccount = "01-01"
                    '   If oRecItem.Fields.Item("OutGoing").Value = "" Then
                    If oRecItem.Fields.Item("TransType").Value <> "46" Then
                        sAPAccount = oRecItem.Fields.Item("Account").Value
                        CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                        CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                        If CostCenter1.Length > 0 Then
                            If CostCenter1.Length > 3 Then
                                CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                            Else
                                CostCenter1 = CostCenter1
                            End If
                        Else
                            CostCenter1 = ""
                        End If
                        If CostCenter2.Length > 0 Then
                            If CostCenter2.Length > 3 Then
                                CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                            Else
                                CostCenter2 = CostCenter2
                            End If
                        Else
                            CostCenter2 = ""
                        End If
                        straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                    Else
                        If strType <> "Cr" Then
                            straccount = oRecItem.Fields.Item("OutGoing").Value
                            If straccount = "" Then
                                straccount = "01-01"
                                sAPAccount = oRecItem.Fields.Item("Account").Value
                                CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                                CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                                If CostCenter1.Length > 0 Then
                                    If CostCenter1.Length > 3 Then
                                        CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                    Else
                                        CostCenter1 = CostCenter1
                                    End If
                                Else
                                    CostCenter1 = ""
                                End If
                                If CostCenter2.Length > 0 Then
                                    If CostCenter2.Length > 3 Then
                                        CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                    Else
                                        CostCenter2 = CostCenter2
                                    End If
                                Else
                                    CostCenter2 = ""
                                End If
                                straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount
                            End If
                        Else
                            '  straccount = sAPAccount ' oRecItem.Fields.Item("OutGoing").Value
                            straccount = "01-01"
                            sAPAccount = oRecItem.Fields.Item("Account").Value
                            CostCenter1 = oRecItem.Fields.Item("ProfitCode").Value
                            CostCenter2 = oRecItem.Fields.Item("OcrCode2").Value
                            If CostCenter1.Length > 0 Then
                                If CostCenter1.Length > 3 Then
                                    CostCenter1 = CostCenter1.Substring(CostCenter1.Length - 3, 3)
                                Else
                                    CostCenter1 = CostCenter1
                                End If
                            Else
                                CostCenter1 = ""
                            End If
                            If CostCenter2.Length > 0 Then
                                If CostCenter2.Length > 3 Then
                                    CostCenter2 = CostCenter2.Substring(CostCenter2.Length - 3, 3)
                                Else
                                    CostCenter2 = CostCenter2
                                End If
                            Else
                                CostCenter2 = ""
                            End If
                            straccount = straccount & "-" & CostCenter1 & "-" & CostCenter2 & "-" & sAPAccount

                        End If

                    End If


                    writer.WriteStartElement("AcctId")
                    writer.WriteString(straccount)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EntryType")
                    writer.WriteString(strType)
                    writer.WriteEndElement()

                    writer.WriteStartElement("EffDt")
                    writer.WriteString(dtDateTime)
                    writer.WriteEndElement()

                    '  If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                    If oRecItem.Fields.Item("TransType").Value = "46" Then
                        If strType <> "Dr" Then
                            writer.WriteStartElement("ApplType")
                            writer.WriteString("GL")
                            writer.WriteEndElement()
                            writer.WriteStartElement("AcctType")
                            writer.WriteString("GL")
                            writer.WriteEndElement()
                        Else
                            If oRecItem.Fields.Item("OutGoing").Value <> "" Then
                                writer.WriteStartElement("ApplType")
                                If straccount.StartsWith("1") = True Then
                                    writer.WriteString("CK")
                                ElseIf straccount.StartsWith("2") = True Then
                                    writer.WriteString("SV")
                                Else
                                    writer.WriteString("CK")
                                End If

                                writer.WriteEndElement()
                                writer.WriteStartElement("AcctType")
                                Dim Test As SAPbobsCOM.Recordset
                                Test = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim sa As String
                                sa = "Select * from OCRB where  Account='" & straccount & "'"
                                Test.DoQuery(sa)
                                If Test.RecordCount > 0 Then
                                    If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                        writer.WriteString("MGC")
                                    Else
                                        writer.WriteString(Test.Fields.Item("MandateID").Value)
                                    End If
                                Else

                                    If oRecItem.Fields.Item("OutGoing").Value = "100000027275" Then
                                        writer.WriteString("MGC")
                                    Else
                                        writer.WriteString(Test.Fields.Item("CUR").Value)
                                        writer.WriteString("CUR")
                                    End If
                                    'writer.WriteString("CUR")
                                End If



                                writer.WriteEndElement()

                            Else
                                writer.WriteStartElement("ApplType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                                writer.WriteStartElement("AcctType")
                                writer.WriteString("GL")
                                writer.WriteEndElement()
                            End If

                            strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                            writer.WriteStartElement("ChkNum")
                            writer.WriteString(strCheckNo)
                            writer.WriteEndElement()
                        End If

                    Else
                        writer.WriteStartElement("AcctType")
                        writer.WriteString(strAccountType)
                        writer.WriteEndElement()
                    End If

                    If oRecItem.Fields.Item("TransType").Value = "46" Then
                        Dim strs As SAPbobsCOM.Recordset
                        strs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        strs.DoQuery("Select * from OVPM where DocEntry=" & oRecItem.Fields.Item("BaseRef").Value)

                        If strs.Fields.Item("CheckSum").Value > 0 Then
                            strCheckNo = oRecItem.Fields.Item("CheckNo").Value
                            If strType = "Dr" Then
                                strCheckNo = "MANAGER CHEQUE NO:" & strCheckNo & " Branch:911"
                            Else
                                strCheckNo = "MGC:" & strCheckNo & " Br:911" & " Bnf :" & strs.Fields.Item("Comments").Value & "- Reverse" ' oRecItem.Fields.Item("LineMemo").Value
                            End If
                        Else
                            strCheckNo = ""
                        End If

                        writer.WriteStartElement("Description")
                        If strCheckNo = "" Then
                            writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                        Else
                            writer.WriteString(strCheckNo)
                        End If
                        '  
                        writer.WriteEndElement()
                    Else
                        writer.WriteStartElement("Description")
                        writer.WriteString(oRecItem.Fields.Item("LineMemo").Value)
                        writer.WriteEndElement()
                    End If

                End If
                writer.WriteEndElement()
                oRecItem.MoveNext()
            Next
            ' writer.WriteString(dtDateTime)
            writer.WriteStartElement("BankInfo")
            writer.WriteStartElement("BranchId")
            writer.WriteString("911")
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("EmployeeIdent")
            writer.WriteStartElement("EmployeeIdentlNum")
            writer.WriteString(strPhxId)
            writer.WriteEndElement()
            writer.WriteStartElement("SuperEmployeeIdentlNum")
            writer.WriteString("0")
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("RqUID")
            writer.WriteString("85fb650d-336e-4c1d-bea3-6d9551b9fb08")
            writer.WriteEndElement()

            writer.WriteEndElement()

            writer.WriteStartElement("RqUID")
            writer.WriteString("85fb650d-336e-4c1d-bea3-6d9551b9fb08")
            writer.WriteEndElement()


            writer.WriteEndElement()
            writer.WriteEndElement()
            'writer.WriteEndElement()

            writer.Flush()
            MyXMLString = myStringWriter.ToString()

            myStringWriter.Close()
            writer.Close()
            ' SendXMLtoIFX(MyXMLString)

            Dim IFXResponse As String = SendXMLtoIFX(MyXMLString)
            Dim doc As New XmlDocument
            doc.LoadXml(IFXResponse)
            Try
                'Dim x As String = doc.SelectSingleNode("IFX/BankScvRs/RqUID)").InnerText.ToString
                Dim locx, locy, locy1 As String
                locx = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusCode").InnerText).ToString
                locy = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/Severity").InnerText).ToString
                locy1 = (doc.SelectSingleNode("IFX/BankSvcRs/FinancialMessageAddRs/Status/StatusDesc").InnerText).ToString
                If locx = "0" Then
                    oRecItem.DoQuery("Update OJDT set U_Export='Y',U_Z_IFXReply='" & locy1 & "' where TransId in (" & strJNo & ")")
                    strMessage = "Export Jounral Entry  Compleated . Journal Entry No : " & strJNo
                    oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    strMessage = "Error while transfering Journal Entry No : " & strJNo & " : Error : " & locy1 & " . The transaction has been reversed. Create a new Transaction"
                    oApplication.SBO_Application.MessageBox(strMessage)
                    oRecItem.DoQuery("Update OJDT set U_Z_IFXReply='" & locy1 & "' where TransId in (" & strJNo & ")")
                End If

                '  WriteErrorlog(strMessage, strErrorFileName)
                oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Catch ex As Exception
            End Try
            oMainRec.MoveNext()
        Next
        'If File.Exists(strTRGFileName) Then
        '    File.Delete(strTRGFileName)
        'End If
    End Sub


    Public Function ValidateBPAccunt(ByVal aform As SAPbouiCOM.Form) As Boolean

        Dim oRecItem, oRecItemCode, oTemp, oMainRec As SAPbobsCOM.Recordset
        Dim strSQL, strFilename As String
        Dim sValue As String
        Dim sPath, strLogDirectory, strPath, strMessage, strSelectedFolderPath, strExportFilePaty As String
        Dim blnErrorflag As Boolean
        Dim FILedatetiem As String
        strPath = System.Windows.Forms.Application.StartupPath
        strFilename = Now.ToLongDateString
        ' strPath = aFileName
        oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMainRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'aPath = 1

        '  strSQL = "Select T1.TransId ,T1.RefDate,Count(*) from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where Isnull(T1.U_Export,'N')='N' and T0.TransId=" & CInt(aPath) & " group by T1.TransId,T1.RefDate" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
        ' oMainRec.DoQuery(strSQL)
        Dim strTransID As String
        Dim dtJEDate As Date
        Dim strPhxId As String = ""
        Dim ACCID, ACCTTYPE As String

        Dim omatrix As SAPbouiCOM.Matrix
        omatrix = aform.Items.Item("3").Specific
        For intRow As Integer = 1 To omatrix.RowCount
            ACCID = getMatrixValues(omatrix, "Account", intRow)
            ACCTTYPE = getMatrixValues(omatrix, "MandateID", intRow)
            If getMatrixValues(omatrix, "BankCode", intRow) <> "" Then
                If ACCID.Length <> 12 Then
                    Message("Account code should be 12 digit :  Line ID : " & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If ACCTTYPE = "" Then
                    Message("ManddateID field is missing... : Line ID : &" & intRow, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                FILedatetiem = Now.ToString("yyyyMMdd_hh_mm_ss")
                Dim transtype As String
                Dim strTRGFileName As String
                Dim dtDateTime As String
                Dim strJNo As String = ""
                dtDateTime = dtJEDate.ToString("yyyy-MM-dd") & " _" & strTransID ' Format(Now.Date, "yyyy-MM-dd") & Now.ToLongTimeString.Replace(":", "")
                FILedatetiem = dtDateTime
                dtDateTime = dtJEDate.ToString("yyyy-MM-dd")
                strTRGFileName = strExportFilePaty & "\JE_" & FILedatetiem & ".trg"
                strFilename = strExportFilePaty & "\JE_" & FILedatetiem & ".xml"
                Dim MyXMLString As String
                Dim myStringWriter As New StringWriter
                Dim writer As New XmlTextWriter(myStringWriter)
                '(strFilename, System.Text.Encoding.UTF8)
                writer.WriteStartDocument(True)
                writer.Formatting = Formatting.Indented
                writer.Indentation = 2
                writer.WriteStartElement("IFX")

                writer.WriteStartElement("SignonRq")

                writer.WriteStartElement("RqUID")
                writer.WriteString("1")
                writer.WriteEndElement()

                writer.WriteStartElement("ClientDt")
                writer.WriteString(dtJEDate.ToString("yyyy-MM-dd"))
                writer.WriteEndElement()

                writer.WriteEndElement()


                'BankSvcRq
                writer.WriteStartElement("BankSvcRq")
                writer.WriteStartElement("DepAccountBalanceInqRq")

                writer.WriteStartElement("RqUID")
                writer.WriteString("1")
                writer.WriteEndElement()

                writer.WriteStartElement("ApplType")
                If ACCID.StartsWith("1") Then
                    writer.WriteString("CK")
                ElseIf ACCID.StartsWith("2") Then
                    writer.WriteString("SV")
                Else
                    writer.WriteString("CK")
                End If


                writer.WriteEndElement()

                writer.WriteStartElement("AcctId")
                writer.WriteString(ACCID)
                writer.WriteEndElement()

                writer.WriteStartElement("AcctType")
                writer.WriteString(ACCTTYPE)
                writer.WriteEndElement()
                ' writer.WriteString(dtDateTime)
                writer.WriteEndElement()
                writer.WriteStartElement("RqUID")
                writer.WriteString("1")
                writer.WriteEndElement()
                writer.WriteEndElement()
                writer.WriteEndElement()
                writer.Flush()
                MyXMLString = myStringWriter.ToString()
                myStringWriter.Close()
                writer.Close()
                ' SendXMLtoIFX(MyXMLString)

                Dim IFXResponse As String = SendXMLtoIFX(MyXMLString)
                Dim doc As New XmlDocument
                doc.LoadXml(IFXResponse)
                Try
                    Dim strBPName As String
                    Dim strActive11 As String
                    'Dim x As String = doc.SelectSingleNode("IFX/BankScvRs/RqUID)").InnerText.ToString
                    Dim locx As String = (doc.SelectSingleNode("IFX/BankSvcRs/DepAccountBalanceInqRs/Status/StatusCode").InnerText).ToString
                    Dim locY As String = (doc.SelectSingleNode("IFX/BankSvcRs/DepAccountBalanceInqRs/Status/Severity").InnerText).ToString
                    Dim locY1 As String = (doc.SelectSingleNode("IFX/BankSvcRs/DepAccountBalanceInqRs/Status/StatusDesc").InnerText).ToString

                    Try
                        strActive11 = (doc.SelectSingleNode("IFX/BankSvcRs/DepAccountBalanceInqRs/BankAcctStatusCode").InnerText).ToString
                    Catch ex As Exception
                        strActive11 = locY1 '(doc.SelectSingleNode("IFX/BankSvcRs/DepAccountBalanceInqRs/BankAcctStatusCode").InnerText).ToString
                    End Try

                    Try
                        strBPName = (doc.SelectSingleNode("IFX/BankSvcRs/DepAccountBalanceInqRs/TitleLine1").InnerText).ToString
                    Catch ex As Exception
                        strBPName = ""
                    End Try
                    If locx.Trim() = "0" And strActive11.Trim().ToUpper = "ACTIVE" Then
                        SetMatrixValues(omatrix, "AcctName", intRow, strBPName)
                    Else
                        strBPName = ""
                        strMessage = "Error in Account ID: " & ACCID & " : Error : " & locY1 & " : Bank Account Status : " & strActive11
                        oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        SetMatrixValues(omatrix, "AcctName", intRow, strBPName)
                        Return False
                    End If
                    '   WriteErrorlog(strMessage, strErrorFileName)
                Catch ex As Exception

                End Try
            End If
        Next

        Return True
        'If File.Exists(strTRGFileName) Then
        '    File.Delete(strTRGFileName)
        'End If
    End Function

    Public Function ValidateBPAccunt_EmployeeMaster(ByVal aform As SAPbouiCOM.Form) As Boolean

        Dim oRecItem, oRecItemCode, oTemp, oMainRec As SAPbobsCOM.Recordset
        Dim strSQL, strFilename As String
        Dim sValue As String
        Dim sPath, strLogDirectory, strPath, strMessage, strSelectedFolderPath, strExportFilePaty As String
        Dim blnErrorflag As Boolean
        Dim FILedatetiem As String
        strPath = System.Windows.Forms.Application.StartupPath
        strFilename = Now.ToLongDateString
        ' strPath = aFileName
        oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMainRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'aPath = 1

        '  strSQL = "Select T1.TransId ,T1.RefDate,Count(*) from JDT1 T0 inner Join OJDT T1  On T0.TransId = T1.TransId  inner Join OUSR T2 on T0.UserSign = T2.userid where Isnull(T1.U_Export,'N')='N' and T0.TransId=" & CInt(aPath) & " group by T1.TransId,T1.RefDate" 'CARDCODE IN (select CardCode from OCRD where Cardtype='S' and cardcode='" & aCardCode & "')"
        ' oMainRec.DoQuery(strSQL)
        Dim strTransID As String
        Dim dtJEDate As Date
        Dim strPhxId As String = ""
        Dim ACCID, ACCTTYPE As String

        Dim omatrix As SAPbouiCOM.Matrix
        ' omatrix = aform.Items.Item("3").Specific
        For intRow As Integer = 1 To 1 'omatrix.RowCount
            ACCID = getEdittextvalue(aform, "77")
            ACCTTYPE = getEdittextvalue(aform, "79")
            If ACCID <> "" Then
                If ACCID.Length <> 12 Then
                    Message("Account code should be 12 digit in  Finance Details ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
                If ACCTTYPE = "" Then
                    Message("Account Type field is missing in Finance Details... ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                FILedatetiem = Now.ToString("yyyyMMdd_hh_mm_ss")
                Dim transtype As String
                Dim strTRGFileName As String
                Dim dtDateTime As String
                Dim strJNo As String = ""
                dtDateTime = dtJEDate.ToString("yyyy-MM-dd") & " _" & strTransID ' Format(Now.Date, "yyyy-MM-dd") & Now.ToLongTimeString.Replace(":", "")
                FILedatetiem = dtDateTime
                dtDateTime = dtJEDate.ToString("yyyy-MM-dd")
                strTRGFileName = strExportFilePaty & "\JE_" & FILedatetiem & ".trg"
                strFilename = strExportFilePaty & "\JE_" & FILedatetiem & ".xml"
                Dim MyXMLString As String
                Dim myStringWriter As New StringWriter
                Dim writer As New XmlTextWriter(myStringWriter)
                '(strFilename, System.Text.Encoding.UTF8)
                writer.WriteStartDocument(True)
                writer.Formatting = Formatting.Indented
                writer.Indentation = 2
                writer.WriteStartElement("IFX")

                writer.WriteStartElement("SignonRq")

                writer.WriteStartElement("RqUID")
                writer.WriteString("1")
                writer.WriteEndElement()

                writer.WriteStartElement("ClientDt")
                writer.WriteString(dtJEDate.ToString("yyyy-MM-dd"))
                writer.WriteEndElement()

                writer.WriteEndElement()


                'BankSvcRq
                writer.WriteStartElement("BankSvcRq")
                writer.WriteStartElement("DepAccountBalanceInqRq")

                writer.WriteStartElement("RqUID")
                writer.WriteString("1")
                writer.WriteEndElement()

                writer.WriteStartElement("ApplType")
                If ACCID.StartsWith("1") Then
                    writer.WriteString("CK")
                ElseIf ACCID.StartsWith("2") Then
                    writer.WriteString("SV")
                Else
                    writer.WriteString("CK")
                End If


                writer.WriteEndElement()

                writer.WriteStartElement("AcctId")
                writer.WriteString(ACCID)
                writer.WriteEndElement()

                writer.WriteStartElement("AcctType")
                writer.WriteString(ACCTTYPE)
                writer.WriteEndElement()
                ' writer.WriteString(dtDateTime)
                writer.WriteEndElement()
                writer.WriteStartElement("RqUID")
                writer.WriteString("1")
                writer.WriteEndElement()
                writer.WriteEndElement()
                writer.WriteEndElement()
                writer.Flush()
                MyXMLString = myStringWriter.ToString()
                myStringWriter.Close()
                writer.Close()
                ' SendXMLtoIFX(MyXMLString)

                Dim IFXResponse As String = SendXMLtoIFX(MyXMLString)
                Dim doc As New XmlDocument
                doc.LoadXml(IFXResponse)
                Try
                    Dim strBPName As String
                    'Dim x As String = doc.SelectSingleNode("IFX/BankScvRs/RqUID)").InnerText.ToString
                    Dim locx As String = (doc.SelectSingleNode("IFX/BankSvcRs/DepAccountBalanceInqRs/Status/StatusCode").InnerText).ToString
                    Dim locY As String = (doc.SelectSingleNode("IFX/BankSvcRs/DepAccountBalanceInqRs/Status/Severity").InnerText).ToString
                    Dim locY1 As String = (doc.SelectSingleNode("IFX/BankSvcRs/DepAccountBalanceInqRs/Status/StatusDesc").InnerText).ToString
                    Dim strActive11 As String = (doc.SelectSingleNode("IFX/BankSvcRs/DepAccountBalanceInqRs/BankAcctStatusCode").InnerText).ToString
                    Try
                        strBPName = (doc.SelectSingleNode("IFX/BankSvcRs/DepAccountBalanceInqRs/TitleLine1").InnerText).ToString

                    Catch ex As Exception
                        strBPName = ""
                    End Try
                    If locx = "0" And strActive11.Trim().ToUpper = "ACTIVE" Then
                        '   oRecItem.DoQuery("Update OJDT set U_Export='Y' where TransId in (" & strJNo & ")")
                        '   strMessage = "Export Jounral Entry  Compleated . Journal Entry No : " & strJNo
                        ' Return True
                        ' SetMatrixValues(omatrix, "AcctName", intRow, strBPName)
                    Else
                        strBPName = ""
                        strMessage = "Error in Account ID: " & ACCID & " : Error : " & locY1 & " : Bank Account Status : " & strActive11
                        oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        ' SetMatrixValues(omatrix, "AcctName", intRow, strBPName)
                        Return False
                    End If
                    '   WriteErrorlog(strMessage, strErrorFileName)
                Catch ex As Exception

                End Try
            End If
        Next

        Return True
        'If File.Exists(strTRGFileName) Then
        '    File.Delete(strTRGFileName)
        'End If
    End Function

    Public Function ValidateGLAccunt(ByVal aGL As String, ByVal aCurrency As String, ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oRecItem, oRecItemCode, oTemp, oMainRec As SAPbobsCOM.Recordset
        Dim strSQL, strFilename As String
        Dim sValue As String
        Dim sPath, strLogDirectory, strPath, strMessage, strSelectedFolderPath, strExportFilePaty As String
        Dim blnErrorflag As Boolean
        Dim FILedatetiem As String
        strPath = System.Windows.Forms.Application.StartupPath
        strFilename = Now.ToLongDateString
        ' strPath = aFileName
        oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMainRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'aPath = 1

        '  strSQL = "Select *  from OCRB  where "
        ' oMainRec.DoQuery(strSQL)
        Dim strTransID As String
        Dim dtJEDate As Date
        Dim strPhxId As String = ""
        Dim ACCID, ACCTTYPE, Dept, branch As String
        Dept = getEdittextvalue(aform, "2001")
        branch = getEdittextvalue(aform, "38")
        If Dept.Length > 3 Then
            Dept = Dept.Substring(Dept.Length - 3, 3)
        End If
        If branch.Length > 3 Then
            branch = branch.Substring(branch.Length - 3, 3)
        End If
        ACCID = aGL
        ACCID = "01-01-" & branch & "-" & Dept & "-" & ACCID
        ACCTTYPE = aCurrency
        oRecItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecItemCode = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        FILedatetiem = Now.ToString("yyyyMMdd_hh_mm_ss")
        Dim transtype As String
        Dim strTRGFileName As String
        Dim dtDateTime As String
        Dim strJNo As String = ""
        dtDateTime = dtJEDate.ToString("yyyy-MM-dd") & " _" & strTransID ' Format(Now.Date, "yyyy-MM-dd") & Now.ToLongTimeString.Replace(":", "")
        FILedatetiem = dtDateTime
        dtDateTime = dtJEDate.ToString("yyyy-MM-dd")
        strTRGFileName = strExportFilePaty & "\JE_" & FILedatetiem & ".trg"
        strFilename = strExportFilePaty & "\JE_" & FILedatetiem & ".xml"
        Dim MyXMLString As String
        Dim myStringWriter As New StringWriter
        Dim writer As New XmlTextWriter(myStringWriter)
        '(strFilename, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement("IFX")

        writer.WriteStartElement("SignonRq")

        writer.WriteStartElement("RqUID")
        writer.WriteString("1")
        writer.WriteEndElement()

        writer.WriteStartElement("ClientDt")
        writer.WriteString(dtJEDate.ToString("yyyy-MM-dd"))
        writer.WriteEndElement()

        writer.WriteStartElement("ClientApp")
        writer.WriteString("SAP")
        writer.WriteEndElement()

        writer.WriteStartElement("OperatorId")
        writer.WriteString("1")
        writer.WriteEndElement()

        writer.WriteEndElement()


        'BankSvcRq
        writer.WriteStartElement("BankSvcRq")
        writer.WriteStartElement("GLAcctCheckInqRq")

        writer.WriteStartElement("RqUID")
        writer.WriteString("1")
        writer.WriteEndElement()

        'writer.WriteStartElement("ApplType")
        'writer.WriteString("CK")
        'writer.WriteEndElement()

        writer.WriteStartElement("AcctId")
        writer.WriteString(ACCID)
        writer.WriteEndElement()

        'writer.WriteStartElement("AcctType")
        'writer.WriteString(ACCTTYPE)
        'writer.WriteEndElement()
        ' writer.WriteString(dtDateTime)
        writer.WriteEndElement()
        writer.WriteStartElement("RqUID")
        writer.WriteString("1")
        writer.WriteEndElement()
        writer.WriteEndElement()
        writer.WriteEndElement()
        writer.Flush()
        MyXMLString = myStringWriter.ToString()
        myStringWriter.Close()
        writer.Close()
        ' SendXMLtoIFX(MyXMLString)

        Dim IFXResponse As String = SendXMLtoIFX(MyXMLString)
        Dim doc As New XmlDocument
        doc.LoadXml(IFXResponse)
        Try
            'Dim x As String = doc.SelectSingleNode("IFX/BankScvRs/RqUID)").InnerText.ToString
            Dim locx As String = (doc.SelectSingleNode("IFX/BankSvcRs/GLAcctCheckInqRs/Status/StatusCode").InnerText).ToString
            Dim locY As String = (doc.SelectSingleNode("IFX/BankSvcRs/GLAcctCheckInqRs/Status/Severity").InnerText).ToString
            Dim locY1 As String = (doc.SelectSingleNode("IFX/BankSvcRs/GLAcctCheckInqRs/Status/StatusDesc").InnerText).ToString
            If locx = "0" Then
                '   oRecItem.DoQuery("Update OJDT set U_Export='Y' where TransId in (" & strJNo & ")")
                '   strMessage = "Export Jounral Entry  Compleated . Journal Entry No : " & strJNo
                '    oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                Return True
            Else
                strMessage = "Error in Account ID: " & ACCID & " : Error : " & locY1
                oApplication.SBO_Application.MessageBox(strMessage)

                oApplication.Utilities.Message(strMessage, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False

            End If


            '   WriteErrorlog(strMessage, strErrorFileName)
        Catch ex As Exception
            Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Return True
        'If File.Exists(strTRGFileName) Then
        '    File.Delete(strTRGFileName)
        'End If
    End Function

    Private Function SendXMLtoIFX(ByVal aString As String) As String




        Dim MyService1 As New BisBIntegration.WebReference.UBSWebservice
        Dim myCredentials As New System.Net.CredentialCache
        Dim netCred As New System.Net.NetworkCredential()
        Dim otest As SAPbobsCOM.Recordset
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select * from [@Z_IFXSetup]")
        If otest.RecordCount > 0 Then
            MyService1.Url = otest.Fields.Item("U_Z_URL").Value '  "http://10.10.151.52:7001/UBSWebservice/UBSWebservice.jws?WSDL"
            netCred.UserName = otest.Fields.Item("U_Z_UID").Value ' "FluidUser"
            netCred.Password = otest.Fields.Item("U_Z_PWD").Value '"FluidUser"
            Dim strPwd As String = getLoginPassword(otest.Fields.Item("U_Z_PWD").Value)
            netCred.Password = strPwd
        Else
            MyService1.Url = "http://10.10.151.52:7001/UBSWebservice/UBSWebservice.jws?WSDL"
            netCred.UserName = "FluidUser"
            netCred.Password = "FluidUser"
        End If

        'MyService1.Url = "http://10.10.151.52:7001/UBSWebservice/UBSWebservice.jws?WSDL"
        'netCred.UserName = "FluidUser"
        'netCred.Password = "FluidUser"


        myCredentials.Add(New Uri(MyService1.Url), "Basic", netCred)

        MyService1.Credentials = netCred
        Dim RequestXMLString As New adaptorService
        Dim ResponseXMLString As New adaptorServiceResponse
        RequestXMLString.ifxMessage = aString
        ResponseXMLString = MyService1.adaptorService(RequestXMLString)

        Return (ResponseXMLString.adaptorServiceResult)

    End Function
#End Region

#Region "Add to Import UDT"
    Public Sub AddtoExportUDT(ByVal strCode As String, ByVal strMastercode As String, ByVal strchoice As String, ByVal transType As String)
        Try
            Dim oUsertable As SAPbobsCOM.UserTable
            Dim strsql, sCode, strUpdateQuery As String
            Dim oRec, oRecUpdate As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecUpdate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("Select * from [@Z_EXPORT] where U_Z_DocType='" & strchoice & "' and U_Z_MasterCode='" & strCode & "' and U_Z_Exported='N'")
            If oRec.RecordCount <= 0 Then
                strsql = getMaxCode("@Z_EXPORT", "CODE")
                oUsertable = oApplication.Company.UserTables.Item("Z_EXPORT")
                oUsertable.Code = strsql
                oUsertable.Name = strsql & "M"
                oUsertable.UserFields.Fields.Item("U_Z_DocType").Value = strchoice
                oUsertable.UserFields.Fields.Item("U_Z_MasterCode").Value = strCode
                oUsertable.UserFields.Fields.Item("U_Z_DocNum").Value = strMastercode
                oUsertable.UserFields.Fields.Item("U_Z_Action").Value = transType 'strAction '"A"
                oUsertable.UserFields.Fields.Item("U_Z_CreateDate").Value = Now.Date
                oUsertable.UserFields.Fields.Item("U_Z_CreateTime").Value = Now.ToShortTimeString.Replace(":", "")
                oUsertable.UserFields.Fields.Item("U_Z_Exported").Value = "N"
                If oUsertable.Add <> 0 Then
                    MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            oApplication.SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub




#End Region

#Region "Connect remote Company"
    Public Function ConnectRemoteCompany(ByVal aCompDB As String, ByVal aSAPUID As String, ByVal aSAPPWD As String) As SAPbobsCOM.Company
        Dim oRemCompany As SAPbobsCOM.Company
        oRemCompany = New SAPbobsCOM.Company
        With oRemCompany
            .Server = oApplication.Company.Server
            .DbServerType = oApplication.Company.DbServerType
            .LicenseServer = oApplication.Company.LicenseServer
            .UserName = aSAPUID
            .Password = aSAPPWD
            .CompanyDB = aCompDB
            If .Connect <> 0 Then
                Message("Connection to : " & aCompDB & " faild", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Function
            Else
                Return oRemCompany
            End If
        End With
        Return oRemCompany
    End Function

    Public Function CheckConnection(ByVal aCompDB As String, ByVal aSAPUID As String, ByVal aSAPPWD As String) As Boolean
        Dim oRemCompany As SAPbobsCOM.Company
        oRemCompany = New SAPbobsCOM.Company
        With oRemCompany
            .Server = oApplication.Company.Server
            .DbServerType = oApplication.Company.DbServerType
            .LicenseServer = oApplication.Company.LicenseServer
            .UserName = aSAPUID
            .Password = aSAPPWD
            .CompanyDB = aCompDB
            If .Connect <> 0 Then
                Message("Connection to : " & aCompDB & " faild", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                Return True
            End If
        End With
        Return True
    End Function



#End Region


#Region "Check the Filepaths"
    Private Function ValidateFilePaths(ByVal aPath As String) As Boolean
        Dim strMessage, strpath, strFilename, strErrorLogPath As String
        strErrorLogPath = aPath
        strpath = strErrorLogPath ' System.Windows.Forms.Application.StartupPath
        If Directory.Exists(strpath) = False Then
            System.IO.Directory.CreateDirectory(strpath)
            Return False
        End If

        Return True
    End Function
#End Region
#Region "Write into ErrorLog File"
    Public Sub WriteErrorHeader(ByVal apath As String, ByVal strMessage As String)
        Dim aSw As System.IO.StreamWriter
        Dim aMessage As String
        aMessage = Now.Date.ToString("dd/MM/yyyy") & ":" & Now.ToShortTimeString.Replace(":", "") & " --> " & strMessage
        If File.Exists(apath) Then
        End If
        aSw = New StreamWriter(apath, True)
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
#End Region

#Region "Export Documents Details"
    Public Sub ExportSKU(ByVal aPath As String, ByVal aChoice As String)
        If aChoice <> "SKU" Then
            Exit Sub
        End If
        Dim strPath, strFilename, strMessage, strExportFilePaty, strErrorLog, strTime As String
        strPath = aPath ' System.Windows.Forms.Application.StartupPath
        strTime = Now.ToShortTimeString.Replace(":", "")
        strFilename = Now.Date.ToString("ddMMyyyy")
        strFilename = strFilename & strTime
        Dim stHours, stMin As String

        strErrorLog = ""
        If aChoice = "SKU" Then
            strErrorLog = strPath & "\Logs\SKU Import"
            strPath = strPath & "\Export\SKU Import"
        End If
        strExportFilePaty = strPath
        If Directory.Exists(strPath) = False Then
            Directory.CreateDirectory(strPath)
        End If
        If Directory.Exists(strErrorLog) = False Then
            Directory.CreateDirectory(strErrorLog)
        End If
        strFilename = "Export SKU_" & strFilename
        strErrorLog = strErrorLog & "\" & strFilename & ".txt"
        Message("Processing SKU's Exporting...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        WriteErrorHeader(strErrorLog, "Export SKU's Starting..")
        If Directory.Exists(strExportFilePaty) = False Then
            Directory.CreateDirectory(strPath)
        End If
        Try
            Dim s As System.Text.StringBuilder = New System.Text.StringBuilder
            Dim strRecquery, strdocnum As String
            Dim oCheckrs As SAPbobsCOM.Recordset
            oCheckrs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strRecquery = "SELECT T0.[ItemCode], T0.[ItemName], T1.[ItmsGrpNam], T0.[ItemType], T0.[SWeight1], T0.[SVolume], T0.[CodeBars] FROM OITM T0  INNER JOIN OITB T1 ON T0.ItmsGrpCod = T1.ItmsGrpCod  and T0.ItemCode in (Select U_Z_Mastercode from [@Z_EXPORT] where U_Z_DocType='SKU' and U_Z_Exported='N')"
            oCheckrs.DoQuery(strRecquery)
            oApplication.Utilities.Message("Exporting SKU's in process.....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If oCheckrs.RecordCount > 0 Then
                Dim otemprec As SAPbobsCOM.Recordset
                otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                otemprec.DoQuery(strRecquery)
                If 1 = 1 Then
                    s.Remove(0, s.Length)
                    Dim cols As Integer = 2 ' Me.DataSet1.SalesOrder.Columns.Count
                    strdocnum = ""

                    s.Remove(0, s.Length)
                    Dim strItem As String
                    strItem = ""
                    For intRow As Integer = 0 To otemprec.RecordCount - 1
                        Dim strQt, strStoreKey, strName, groupname, itemtype, weight, volume, expirable, codebars, packkey, defaultuom As String
                        strQt = CStr(otemprec.Fields.Item(1).Value)
                        If strItem = "" Then
                            strItem = "'" & otemprec.Fields.Item("ItemCode").Value & "'"
                        Else
                            strItem = strItem & ",'" & otemprec.Fields.Item("ItemCode").Value & "'"
                        End If
                        strName = otemprec.Fields.Item("ItemName").Value
                        strStoreKey = ""
                        expirable = ""
                        s.Remove(0, s.Length)
                        s.Append("'" + otemprec.Fields.Item(0).Value + "'")
                        s.Append(",'" + otemprec.Fields.Item(1).Value + "'")
                        s.Append(",'" + otemprec.Fields.Item(2).Value + "'")
                        s.Append(",'" + otemprec.Fields.Item(3).Value.ToString + "'")
                        s.Append(",'" + otemprec.Fields.Item(4).Value.ToString + "'")
                        s.Append(",'" + otemprec.Fields.Item(5).Value.ToString + "'")
                        s.Append(",'" + otemprec.Fields.Item(6).Value.ToString + "'")
                        Dim strLine, strTableQuery, strfields As String
                        strLine = s.ToString
                        strfields = "([SKU],[DESCR],[ItemGroup],[ItemType],[Weight],[Volume],[Barcode])"
                        ' strTableQuery = "Insert into  " & strSKUExportTable & strfields & " values (" & strLine & ")"
                        Dim oInserQuery As SAPbobsCOM.Recordset
                        oInserQuery = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oInserQuery.DoQuery(strTableQuery)
                        otemprec.MoveNext()
                    Next

                    Dim filename As String
                    strMessage = strItem & "--> SKU's  Exported "
                    Dim oUpdateRS As SAPbobsCOM.Recordset
                    Dim strUpdate As String
                    oUpdateRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    filename = ""
                    strUpdate = "Update [@Z_EXPORT] set U_Z_Exported='Y' ,U_Z_ExportMethod='M',U_Z_ExportFile='" & filename & "',U_Z_ExportDate=getdate() where U_Z_MasterCode in (" & strItem & ") and U_Z_DocType='SKU'"
                    oUpdateRS.DoQuery(strUpdate)
                    WriteErrorlog(strMessage, strErrorLog)
                End If
            Else
                strMessage = ("No new SKUs!")
                WriteErrorlog(strMessage, strErrorLog)
            End If
        Catch ex As Exception
            strMessage = ("An Error Occured. A log entry has been created." & ex.Message)
            WriteErrorlog(strMessage, strErrorLog)
        Finally
            strMessage = "Export process compleated"
            WriteErrorlog(strMessage, strErrorLog)
        End Try
        '   System.Windows.Forms.Application.Exit()
    End Sub


#End Region

#Region "Import Documents"

    Public Sub ImportASNFiles(ByVal apath As String)

    End Sub



#Region "Get StoreKey"
    Public Function getStoreKey() As String
        Dim stStorekey As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'oTemp.DoQuery("Select isnull(U_Z_Storekey,'') from OADM")
        'Return oTemp.Fields.Item(0).Value
        Return ""
    End Function
#End Region
#End Region

#Region "Close Open Sales Order Lines"


    Public Sub WriteErrorlog(ByVal aMessage As String, ByVal aPath As String)
        Dim aSw As System.IO.StreamWriter
        Try
            If File.Exists(aPath) Then
            End If
            aSw = New StreamWriter(aPath, True)
            aMessage = Now.Date.ToString("dd/MM/yyyy") & ":" & Now.ToShortTimeString.Replace(":", "") & " --> " & aMessage
            aSw.WriteLine(aMessage)
            aSw.Flush()
            aSw.Close()
            aSw.Dispose()
        Catch ex As Exception
            Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

    Public Sub createARINvoice()
        Dim strCardcode, stritemcode As String
        Dim intbaseEntry, intbaserow As Integer
        Dim oInv As SAPbobsCOM.Documents
        strCardcode = "C20000"
        intbaseEntry = 66
        intbaserow = 1
        oInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        oInv.DocDate = Now.Date
        oInv.CardCode = strCardcode
        oInv.Lines.BaseType = 17
        oInv.Lines.BaseEntry = intbaseEntry
        oInv.Lines.BaseLine = intbaserow
        oInv.Lines.Quantity = 1
        If oInv.Add <> 0 Then
            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Else
            oApplication.Utilities.Message("AR Invoice added", SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        End If

    End Sub
    Public Sub CloseOpenSOLines()
        Try
            Dim oDoc As SAPbobsCOM.Documents
            Dim oTemp As SAPbobsCOM.Recordset
            Dim strSQL, strSQL1, spath As String
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            spath = System.Windows.Forms.Application.StartupPath & "\Sales Order Matching ErrorLog.txt"
            If File.Exists(spath) Then
                File.Delete(spath)
            End If
            blnError = False
            ' oTemp.DoQuery("Select DocEntry,LineNum from RDR1 where isnull(trgetentry,0)=0 and  LineStatus='O' and Quantity = isnull(U_RemQty,0) order by DocEntry,LineNum")
            '            oTemp.DoQuery("Select DocEntry,VisOrder,LineNum from RDR1 where isnull(trgetentry,0)=0 and  LineStatus='O' and Quantity = isnull(U_RemQty,0) order by DocEntry,LineNum")
            oTemp.DoQuery("Select DocEntry,VisOrder,LineNum from RDR1 where   LineStatus='O' and Quantity = isnull(U_RemQty,0) order by DocEntry,LineNum")
            oApplication.Utilities.Message("Processing closing Sales order Lines", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Dim numb As Integer
            For introw As Integer = 0 To oTemp.RecordCount - 1
                oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                numb = oTemp.Fields.Item(1).Value
                '  numb = oTemp.Fields.Item(2).Value
                If oDoc.GetByKey(oTemp.Fields.Item("DocEntry").Value) Then
                    oApplication.Utilities.Message("Processing Sales order :" & oDoc.DocNum, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oDoc.Comments = oDoc.Comments & "XXX1"
                    If oDoc.Update() <> 0 Then
                        WriteErrorlog(" Error in Closing Sales order Line  SO No :" & oDoc.DocNum & " : Line No : " & oDoc.Lines.LineNum & " : Error : " & oApplication.Company.GetLastErrorDescription, spath)
                        blnError = True
                    Else
                        oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                        If oDoc.GetByKey(oTemp.Fields.Item("DocEntry").Value) Then
                            Dim strcomments As String
                            strcomments = oDoc.Comments
                            strcomments = strcomments.Replace("XXX1", "")
                            oDoc.Comments = strcomments
                            oDoc.Lines.SetCurrentLine(numb)
                            '  MsgBox(oDoc.Lines.VisualOrder)
                            If oDoc.Lines.LineStatus <> SAPbobsCOM.BoStatus.bost_Close Then
                                oDoc.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                            End If
                            If oDoc.Update <> 0 Then
                                WriteErrorlog(" Error in Closing Sales order Line  SO No :" & oDoc.DocNum & " : Line No : " & oDoc.Lines.LineNum & " : Error : " & oApplication.Company.GetLastErrorDescription, spath)
                                blnError = True
                                'oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Else
                                WriteErrorlog(" Sales order Line  SO No :" & oDoc.DocNum & " : Line No : " & oDoc.Lines.LineNum & " : Closed successfully  ", spath)
                            End If
                        End If
                    End If

                End If
                oTemp.MoveNext()
            Next
            oApplication.Utilities.Message("Operation completed succesfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            blnError = True
            ' oApplication.SBO_Application.MessageBox("Error Occured...")\
            spath = System.Windows.Forms.Application.StartupPath & "\Sales Order Matching ErrorLog.txt"
            If File.Exists(spath) Then
                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                x.FileName = spath
                System.Diagnostics.Process.Start(x)
                x = Nothing
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region


#Region "Connect to Company"
    Public Sub Connect()
        Dim strCookie As String
        Dim strConnectionContext As String

        Try
            strCookie = oApplication.Company.GetContextCookie
            strConnectionContext = oApplication.SBO_Application.Company.GetConnectionContext(strCookie)

            If oApplication.Company.SetSboLoginContext(strConnectionContext) <> 0 Then
                Throw New Exception("Wrong login credentials.")
            End If

            'Open a connection to company
            If oApplication.Company.Connect() <> 0 Then
                Throw New Exception("Cannot connect to company database. : " & oApplication.Company.GetLastErrorDescription)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Genral Functions"

#Region "Get MaxCode"
    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            ExecuteSQL(oRS, strSQL)

            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region "Status Message"
    Public Sub Message(ByVal sMessage As String, ByVal StatusType As SAPbouiCOM.BoStatusBarMessageType)
        oApplication.SBO_Application.StatusBar.SetText(sMessage, SAPbouiCOM.BoMessageTime.bmt_Short, StatusType)
    End Sub
#End Region

#Region "Add Choose from List"
    Public Sub AddChooseFromList(ByVal FormUID As String, ByVal CFL_Text As String, ByVal CFL_Button As String, _
                                        ByVal ObjectType As SAPbouiCOM.BoLinkedObject, _
                                            Optional ByVal AliasName As String = "", Optional ByVal CondVal As String = "", _
                                                    Optional ByVal Operation As SAPbouiCOM.BoConditionOperation = SAPbouiCOM.BoConditionOperation.co_EQUAL)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            If ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items Then
                oCFLCreationParams.MultiSelection = True
            Else
                oCFLCreationParams.MultiSelection = False
            End If

            oCFLCreationParams.ObjectType = ObjectType
            oCFLCreationParams.UniqueID = CFL_Text

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            If Not AliasName = "" Then
                oCon = oCons.Add()
                oCon.Alias = AliasName
                oCon.Operation = Operation
                oCon.CondVal = CondVal
                oCFL.SetConditions(oCons)
            End If

            oCFLCreationParams.UniqueID = CFL_Button
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Linked Object Type"
    Public Function getLinkedObjectType(ByVal Type As SAPbouiCOM.BoLinkedObject) As String
        Return CType(Type, String)
    End Function

#End Region

#Region "Execute Query"
    Public Sub ExecuteSQL(ByRef oRecordSet As SAPbobsCOM.Recordset, ByVal SQL As String)
        Try
            If oRecordSet Is Nothing Then
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            End If

            oRecordSet.DoQuery(SQL)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Application path"
    Public Function getApplicationPath() As String

        Return Application.StartupPath.Trim

        'Return IO.Directory.GetParent(Application.StartupPath).ToString
    End Function
#End Region

#Region "Date Manipulation"

#Region "Convert SBO Date to System Date"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	ConvertStrToDate
    'Parameter          	:   ByVal oDate As String, ByVal strFormat As String
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	07/12/05
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To convert Date according to current culture info
    '********************************************************************
    Public Function ConvertStrToDate(ByVal strDate As String, ByVal strFormat As String) As DateTime
        Try
            Dim oDate As DateTime
            Dim ci As New System.Globalization.CultureInfo("en-GB", False)
            Dim newCi As System.Globalization.CultureInfo = CType(ci.Clone(), System.Globalization.CultureInfo)

            System.Threading.Thread.CurrentThread.CurrentCulture = newCi
            oDate = oDate.ParseExact(strDate, strFormat, ci.DateTimeFormat)

            Return oDate
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region

#Region " Get SBO Date Format in String (ddmmyyyy)"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	StrSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(ddmmyy value) as applicable to SBO
    '********************************************************************
    Public Function StrSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String, GetDateFormat As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yy"
                Case 1
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yyyy"
                Case 2
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yy"
                Case 3
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yyyy"
                Case 4
                    GetDateFormat = "yyyy" & DateSep & "dd" & DateSep & "MM"
                Case 5
                    GetDateFormat = "dd" & DateSep & "MMM" & DateSep & "yyyy"
            End Select
            Return GetDateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Get SBO date Format in Number"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	IntSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(integer value) as applicable to SBO
    '********************************************************************
    Public Function NumSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    NumSBODateFormat = 3
                Case 1
                    NumSBODateFormat = 103
                Case 2
                    NumSBODateFormat = 1
                Case 3
                    NumSBODateFormat = 120
                Case 4
                    NumSBODateFormat = 126
                Case 5
                    NumSBODateFormat = 130
            End Select
            Return NumSBODateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#End Region

#Region "Get Rental Period"
    Public Function getRentalDays(ByVal Date1 As String, ByVal Date2 As String, ByVal IsWeekDaysBilling As Boolean) As Integer
        Dim TotalDays, TotalDaysincSat, TotalBillableDays As Integer
        Dim TotalWeekEnds As Integer
        Dim StartDate As Date
        Dim EndDate As Date
        Dim oRecordset As SAPbobsCOM.Recordset

        StartDate = CType(Date1.Insert(4, "/").Insert(7, "/"), Date)
        EndDate = CType(Date2.Insert(4, "/").Insert(7, "/"), Date)

        TotalDays = DateDiff(DateInterval.Day, StartDate, EndDate)

        If IsWeekDaysBilling Then
            strSQL = " select dbo.WeekDays('" & Date1 & "','" & Date2 & "')"
            oApplication.Utilities.ExecuteSQL(oRecordset, strSQL)
            If oRecordset.RecordCount > 0 Then
                TotalBillableDays = oRecordset.Fields.Item(0).Value
            End If
            Return TotalBillableDays
        Else
            Return TotalDays + 1
        End If

    End Function

    Public Function WorkDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long
        Try
            Dim dtFirstSunday As Date
            Dim dtLastSaturday As Date
            Dim lngWorkDays As Long

            ' get first sunday in range
            dtFirstSunday = dtBegin.AddDays((8 - Weekday(dtBegin)) Mod 7)

            ' get last saturday in range
            dtLastSaturday = dtEnd.AddDays(-(Weekday(dtEnd) Mod 7))

            ' get work days between first sunday and last saturday
            lngWorkDays = (((DateDiff(DateInterval.Day, dtFirstSunday, dtLastSaturday)) + 1) / 7) * 5

            ' if first sunday is not begin date
            If dtFirstSunday <> dtBegin Then

                ' assume first sunday is after begin date
                ' add workdays from begin date to first sunday
                lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))

            End If

            ' if last saturday is not end date
            If dtLastSaturday <> dtEnd Then

                ' assume last saturday is before end date
                ' add workdays from last saturday to end date
                lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)

            End If

            WorkDays = lngWorkDays
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function

#End Region

#Region "Get Item Price with Factor"
    Public Function getPrcWithFactor(ByVal CardCode As String, ByVal ItemCode As String, ByVal RntlDays As Integer, ByVal Qty As Double) As Double
        Dim oItem As SAPbobsCOM.Items
        Dim Price, Expressn As Double
        Dim oDataSet, oRecSet As SAPbobsCOM.Recordset

        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oApplication.Utilities.ExecuteSQL(oDataSet, "Select U_RentFac, U_NumDys From [@REN_FACT] order by U_NumDys ")
        If oItem.GetByKey(ItemCode) And oDataSet.RecordCount > 0 Then

            oApplication.Utilities.ExecuteSQL(oRecSet, "Select ListNum from OCRD where CardCode = '" & CardCode & "'")
            oItem.PriceList.SetCurrentLine(oRecSet.Fields.Item(0).Value - 1)
            Price = oItem.PriceList.Price
            Expressn = 0
            oDataSet.MoveFirst()

            While RntlDays > 0

                If oDataSet.EoF Then
                    oDataSet.MoveLast()
                End If

                If RntlDays < oDataSet.Fields.Item(1).Value Then
                    Expressn += (oDataSet.Fields.Item(0).Value * RntlDays * Price * Qty)
                    RntlDays = 0
                    Exit While
                End If
                Expressn += (oDataSet.Fields.Item(0).Value * oDataSet.Fields.Item(1).Value * Price * Qty)
                RntlDays -= oDataSet.Fields.Item(1).Value
                oDataSet.MoveNext()

            End While

        End If
        If oItem.UserFields.Fields.Item("U_Rental").Value = "Y" Then
            Return CDbl(Expressn / Qty)
        Else
            Return Price
        End If


    End Function
#End Region

#Region "Get WareHouse List"
    Public Function getUsedWareHousesList(ByVal ItemCode As String, ByVal Quantity As Double) As DataTable
        Dim oDataTable As DataTable
        Dim oRow As DataRow
        Dim rswhs As SAPbobsCOM.Recordset
        Dim LeftQty As Double
        Try
            oDataTable = New DataTable
            oDataTable.Columns.Add(New System.Data.DataColumn("ItemCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("WhsCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("Quantity"))

            strSQL = "Select WhsCode, ItemCode, (OnHand + OnOrder - IsCommited) As Available From OITW Where ItemCode = '" & ItemCode & "' And " & _
                        "WhsCode Not In (Select Whscode From OWHS Where U_Reserved = 'Y' Or U_Rental = 'Y') Order By (OnHand + OnOrder - IsCommited) Desc "

            ExecuteSQL(rswhs, strSQL)
            LeftQty = Quantity

            While Not rswhs.EoF
                oRow = oDataTable.NewRow()

                oRow.Item("WhsCode") = rswhs.Fields.Item("WhsCode").Value
                oRow.Item("ItemCode") = rswhs.Fields.Item("ItemCode").Value

                LeftQty = LeftQty - CType(rswhs.Fields.Item("Available").Value, Double)

                If LeftQty <= 0 Then
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double) + LeftQty
                    oDataTable.Rows.Add(oRow)
                    Exit While
                Else
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double)
                End If

                oDataTable.Rows.Add(oRow)
                rswhs.MoveNext()
                oRow = Nothing
            End While

            'strSQL = ""
            'For count As Integer = 0 To oDataTable.Rows.Count - 1
            '    strSQL += oDataTable.Rows(count).Item("WhsCode") & " : " & oDataTable.Rows(count).Item("Quantity") & vbNewLine
            'Next
            'MessageBox.Show(strSQL)

            Return oDataTable

        Catch ex As Exception
            Throw ex
        Finally
            oRow = Nothing
        End Try
    End Function
#End Region

#Region "GetDocumentQuantity"
    Public Function getDocumentQuantity(ByVal strQuantity As String) As Double
        Dim dblQuant As Double
        Dim strTemp, strtemp1 As String
        strTemp = CompanyDecimalSeprator
        strtemp1 = strQuantity
        If strtemp1 = "" Then
            Return 0
        End If
        Dim otest As SAPbobsCOM.Recordset
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select * from OCRN")
        For introw As Integer = 0 To otest.RecordCount - 1
            strQuantity = strQuantity.Replace(otest.Fields.Item("CurrCode").Value, "")
            otest.MoveNext()
        Next

        If CompanyDecimalSeprator <> "." Then
            If CompanyThousandSeprator <> strTemp Then
            End If
            strQuantity = strQuantity.Replace(".", ",")
        End If
        Try
            dblQuant = Convert.ToDouble(strQuantity)
        Catch ex As Exception
            dblQuant = CDbl(strtemp1)
        End Try

        Return dblQuant
    End Function
#End Region


#Region "Set / Get Values from Matrix"
    Public Function getMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer) As String
        Return aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value
    End Function
    Public Sub SetMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer, ByVal strvalue As String)
        aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value = strvalue
    End Sub
#End Region

#Region "Get Edit Text"
    Public Function getEdittextvalue(ByVal aform As SAPbouiCOM.Form, ByVal UID As String) As String
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aform.Items.Item(UID).Specific
        Return objEdit.String
    End Function
    Public Sub setEdittextvalue(ByVal aform As SAPbouiCOM.Form, ByVal UID As String, ByVal newvalue As String)
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aform.Items.Item(UID).Specific
        objEdit.String = newvalue
    End Sub
#End Region

#End Region

    Public Function GetCode(ByVal sTableName As String) As String
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim sQuery As String
        oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        sQuery = "SELECT Top 1 DocEntry FROM " & sTableName + " ORDER BY Convert(Int,DocEntry) desc"
        oRecSet.DoQuery(sQuery)
        If Not oRecSet.EoF Then
            GetCode = Convert.ToInt32(oRecSet.Fields.Item(0).Value.ToString()) + 1
        Else
            GetCode = "1"
        End If
    End Function

#Region "Functions related to Load XML"

#Region "Add/Remove Menus "
    Public Sub AddRemoveMenus(ByVal sFileName As String)
        Dim oXMLDoc As New Xml.XmlDocument
        Dim sFilePath As String
        Try
            sFilePath = getApplicationPath() & "\XML Files\" & sFileName
            oXMLDoc.Load(sFilePath)
            oApplication.SBO_Application.LoadBatchActions(oXMLDoc.InnerXml)
        Catch ex As Exception
            Throw ex
        Finally
            oXMLDoc = Nothing
        End Try
    End Sub
#End Region

#Region "Load XML File "
    Private Function LoadXMLFiles(ByVal sFileName As String) As String
        Dim oXmlDoc As Xml.XmlDocument
        Dim oXNode As Xml.XmlNode
        Dim oAttr As Xml.XmlAttribute
        Dim sPath As String
        Dim FrmUID As String
        Try
            oXmlDoc = New Xml.XmlDocument

            sPath = getApplicationPath() & "\XML Files\" & sFileName

            oXmlDoc.Load(sPath)
            oXNode = oXmlDoc.GetElementsByTagName("form").Item(0)
            oAttr = oXNode.Attributes.GetNamedItem("uid")
            oAttr.Value = oAttr.Value & FormNum
            FormNum = FormNum + 1
            oApplication.SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
            FrmUID = oAttr.Value

            Return FrmUID

        Catch ex As Exception
            Throw ex
        Finally
            oXmlDoc = Nothing
        End Try
    End Function
#End Region

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        'Return LoadForm(XMLFile, FormType.ToString(), FormType & "_" & oApplication.SBO_Application.Forms.Count.ToString)
        LoadXMLFiles(XMLFile)
        Return Nothing
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : LoadForm
    'Parameter          : XmlFile,FormType,FormUID
    'Return Value       : SBO Form
    'Author             : Senthil Kumar B Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load XML file 
    '*****************************************************************

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String, ByVal FormUID As String) As SAPbouiCOM.Form

        Dim oXML As System.Xml.XmlDocument
        Dim objFormCreationParams As SAPbouiCOM.FormCreationParams
        Try
            oXML = New System.Xml.XmlDocument
            oXML.Load(XMLFile)
            objFormCreationParams = (oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams))
            objFormCreationParams.XmlData = oXML.InnerXml
            objFormCreationParams.FormType = FormType
            objFormCreationParams.UniqueID = FormUID
            Return oApplication.SBO_Application.Forms.AddEx(objFormCreationParams)
        Catch ex As Exception
            Throw ex

        End Try

    End Function



#Region "Load Forms"
    Public Sub LoadForm(ByRef oObject As Object, ByVal XmlFile As String)
        Try
            oObject.FrmUID = LoadXMLFiles(XmlFile)
            oObject.Form = oApplication.SBO_Application.Forms.Item(oObject.FrmUID)
            If Not oApplication.Collection.ContainsKey(oObject.FrmUID) Then
                oApplication.Collection.Add(oObject.FrmUID, oObject)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#End Region

#Region "Functions related to System Initilization"

#Region "Create Tables"
    Public Sub CreateTables()
        Dim oCreateTable As clsTable
        'Dim lRetCode As Integer
        'Dim sErrMsg As String
        'Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try
            oCreateTable = New clsTable
            oCreateTable.CreateTables()


            'oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'oUserTablesMD.TableName = "Z_EXPORT"
            'oUserTablesMD.TableDescription = "WMS Implementation"
            'oUserTablesMD.TableType = SAPbobsCOM.BoUTBTableType.bott_NoObject
            'lRetCode = oUserTablesMD.Add
            ''// check for errors in the process
            'If lRetCode <> 0 Then
            '    If lRetCode = -1 Then
            '    Else
            '        oApplication.Company.GetLastError(lRetCode, sErrMsg)
            '        MsgBox(sErrMsg)
            '    End If
            'Else
            '    MsgBox("Table: " & oUserTablesMD.TableName & " was added successfully")
            'End If
        Catch ex As Exception
            Throw ex
        Finally
            oCreateTable = Nothing
        End Try
    End Sub
#End Region
#Region "AddControls"
    Public Sub AddControls(ByVal objForm As SAPbouiCOM.Form, ByVal ItemUID As String, ByVal SourceUID As String, ByVal ItemType As SAPbouiCOM.BoFormItemTypes, ByVal position As String, Optional ByVal fromPane As Integer = 1, Optional ByVal toPane As Integer = 1, Optional ByVal linkedUID As String = "", Optional ByVal strCaption As String = "", Optional ByVal dblWidth As Double = 0, Optional ByVal dblTop As Double = 0, Optional ByVal Hight As Double = 0)
        Dim objNewItem, objOldItem As SAPbouiCOM.Item
        Dim ostatic As SAPbouiCOM.StaticText
        Dim oButton As SAPbouiCOM.Button
        Dim oCheckbox As SAPbouiCOM.CheckBox
        Dim ofolder As SAPbouiCOM.Folder
        objOldItem = objForm.Items.Item(SourceUID)
        objNewItem = objForm.Items.Add(ItemUID, ItemType)
        With objNewItem
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                .Left = objOldItem.Left - 15
                .Top = objOldItem.Top + 1
                .LinkTo = linkedUID
            Else
                If position.ToUpper = "RIGHT" Then
                    .Left = objOldItem.Left + objOldItem.Width + 2
                    .Top = objOldItem.Top
                    .Height = objOldItem.Height

                ElseIf position.ToUpper = "DOWN" Then
                    .Top = objOldItem.Top + objOldItem.Height + 3
                    .Left = objOldItem.Left
                    .Width = objOldItem.Width

                End If
            End If
            .FromPane = fromPane
            .ToPane = toPane
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                .LinkTo = linkedUID
            End If
            .LinkTo = linkedUID
        End With
        If (ItemType = SAPbouiCOM.BoFormItemTypes.it_EDIT Or ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC) Then
            objNewItem.Width = objOldItem.Width
        End If
        If ItemType = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
            objNewItem.Width = objOldItem.Width '+ 50
            oButton = objNewItem.Specific
            oButton.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_FOLDER Then
            ofolder = objNewItem.Specific
            ofolder.Caption = strCaption
            ofolder.GroupWith(linkedUID)
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
            ostatic = objNewItem.Specific
            ostatic.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX Then
            oCheckbox = objNewItem.Specific
            oCheckbox.Caption = strCaption

        End If
        If dblWidth <> 0 Then
            objNewItem.Width = dblWidth
        End If

        If dblTop <> 0 Then
            objNewItem.Top = objNewItem.Top + dblTop
        End If
        If Hight <> 0 Then
            objNewItem.Height = objNewItem.Height + Hight
        End If
    End Sub
#End Region

#Region "Notify Alert"
    Public Sub NotifyAlert()
        'Dim oAlert As clsPromptAlert

        'Try
        '    oAlert = New clsPromptAlert
        '    oAlert.AlertforEndingOrdr()
        'Catch ex As Exception
        '    Throw ex
        'Finally
        '    oAlert = Nothing
        'End Try

    End Sub
#End Region

#End Region

#Region "Function related to Quantities"

#Region "Get Available Quantity"
    Public Function getAvailableQty(ByVal ItemCode As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset

        strSQL = "Select SUM(T1.OnHand + T1.OnOrder - T1.IsCommited) From OITW T1 Left Outer Join OWHS T3 On T3.Whscode = T1.WhsCode " & _
                    "Where T1.ItemCode = '" & ItemCode & "'"
        Me.ExecuteSQL(rsQuantity, strSQL)

        If rsQuantity.Fields.Item(0) Is System.DBNull.Value Then
            Return 0
        Else
            Return CLng(rsQuantity.Fields.Item(0).Value)
        End If

    End Function
#End Region

#Region "Get Rented Quantity"
    Public Function getRentedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim RentedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_RDR1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_ORDR] Where U_Status = 'R') " & _
                    " and '" & StartDate & "' between [@REN_RDR1].U_ShipDt1 and [@REN_RDR1].U_ShipDt2 "
        '" and [@REN_RDR1].U_ShipDt1 between '" & StartDate & "' and '" & EndDate & "'"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            RentedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return RentedQty

    End Function
#End Region

#Region "Get Reserved Quantity"
    Public Function getReservedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim ReservedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_QUT1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_OQUT] Where U_Status = 'R' And Status = 'O') " & _
                    " and '" & StartDate & "' between [@REN_QUT1].U_ShipDt1 and [@REN_QUT1].U_ShipDt2"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            ReservedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return ReservedQty

    End Function
#End Region

#End Region

#Region "Functions related to Tax"

#Region "Get Tax Codes"
    Public Sub getTaxCodes(ByRef oCombo As SAPbouiCOM.ComboBox)
        Dim rsTaxCodes As SAPbobsCOM.Recordset

        strSQL = "Select Code, Name From OVTG Where Category = 'O' Order By Name"
        Me.ExecuteSQL(rsTaxCodes, strSQL)

        oCombo.ValidValues.Add("", "")
        If rsTaxCodes.RecordCount > 0 Then
            While Not rsTaxCodes.EoF
                oCombo.ValidValues.Add(rsTaxCodes.Fields.Item(0).Value, rsTaxCodes.Fields.Item(1).Value)
                rsTaxCodes.MoveNext()
            End While
        End If
        oCombo.ValidValues.Add("Define New", "Define New")
        'oCombo.Select("")
    End Sub
#End Region

#Region "Get Applicable Code"

    Public Function getApplicableTaxCode1(ByVal CardCode As String, ByVal ItemCode As String, ByVal Shipto As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    strSQL = "select LicTradNum from CRD1 where Address ='" & Shipto & "' and CardCode ='" & CardCode & "'"
                    Me.ExecuteSQL(rsExempt, strSQL)
                    If rsExempt.RecordCount > 0 Then
                        rsExempt.MoveFirst()
                        TaxGroup = rsExempt.Fields.Item(0).Value
                    Else
                        TaxGroup = ""
                    End If
                    'TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If




        Return TaxGroup

    End Function


    Public Function getApplicableTaxCode(ByVal CardCode As String, ByVal ItemCode As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If

        'If oBP.GetByKey(CardCode.Trim) Then
        '    If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
        '        If oBP.VatGroup.Trim <> "" Then
        '            TaxGroup = oBP.VatGroup.Trim
        '        Else
        '            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '            If oItem.GetByKey(ItemCode.Trim) Then
        '                TaxGroup = oItem.SalesVATGroup.Trim
        '            End If
        '        End If
        '    ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
        '        strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
        '        Me.ExecuteSQL(rsExempt, strSQL)
        '        If rsExempt.RecordCount > 0 Then
        '            rsExempt.MoveFirst()
        '            TaxGroup = rsExempt.Fields.Item(0).Value
        '        Else
        '            TaxGroup = ""
        '        End If
        '    End If
        'End If
        Return TaxGroup

    End Function
#End Region

#End Region

#Region "Log Transaction"
    Public Sub LogTransaction(ByVal DocNum As Integer, ByVal ItemCode As String, _
                                    ByVal FromWhs As String, ByVal TransferedQty As Double, ByVal ProcessDate As Date)
        Dim sCode As String
        Dim sColumns As String
        Dim sValues As String
        Dim rsInsert As SAPbobsCOM.Recordset

        sCode = Me.getMaxCode("@REN_PORDR", "Code")

        sColumns = "Code, Name, U_DocNum, U_WhsCode, U_ItemCode, U_Quantity, U_RetQty, U_Date"
        sValues = "'" & sCode & "','" & sCode & "'," & DocNum & ",'" & FromWhs & "','" & ItemCode & "'," & TransferedQty & ", 0, Convert(DateTime,'" & ProcessDate.ToString("yyyyMMdd") & "')"

        strSQL = "Insert into [@REN_PORDR] (" & sColumns & ") Values (" & sValues & ")"
        oApplication.Utilities.ExecuteSQL(rsInsert, strSQL)

    End Sub

    Public Sub LogCreatedDocument(ByVal DocNum As Integer, ByVal CreatedDocType As SAPbouiCOM.BoLinkedObject, ByVal CreatedDocNum As String, ByVal sCreatedDate As String)
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim sCode As String
        Dim CreatedDate As DateTime
        Try
            oUserTable = oApplication.Company.UserTables.Item("REN_DORDR")

            sCode = Me.getMaxCode("@REN_DORDR", "Code")

            If Not oUserTable.GetByKey(sCode) Then
                oUserTable.Code = sCode
                oUserTable.Name = sCode

                With oUserTable.UserFields.Fields
                    .Item("U_DocNum").Value = DocNum
                    .Item("U_DocType").Value = CInt(CreatedDocType)
                    .Item("U_DocEntry").Value = CInt(CreatedDocNum)

                    If sCreatedDate <> "" Then
                        CreatedDate = CDate(sCreatedDate.Insert(4, "/").Insert(7, "/"))
                        .Item("U_Date").Value = CreatedDate
                    Else
                        .Item("U_Date").Value = CDate(Format(Now, "Long Date"))
                    End If

                End With

                If oUserTable.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserTable = Nothing
        End Try
    End Sub
#End Region

    Public Function FormatDataSourceValue(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If

            If Value.IndexOf(CompanyThousandSeprator) > -1 Then
                Value = Value.Replace(CompanyThousandSeprator, "")
            End If
        Else
            Value = "0"

        End If

        ' NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue


        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue
    End Function

    Public Function FormatScreenValues(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If
        Else
            Value = "0"
        End If

        'NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue

        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue

    End Function

    Public Function SetScreenValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

    Public Function SetDBValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function



End Class
