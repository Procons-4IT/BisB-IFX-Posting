Public NotInheritable Class clsTable

#Region "Private Functions"
    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Tables in DB. This function shall be called by 
    '                     public functions to create a table
    '**************************************************************************************************************
    Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try
            oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'Adding Table
            If Not oUserTablesMD.GetByKey(strTab) Then
                oUserTablesMD.TableName = strTab
                oUserTablesMD.TableDescription = strDesc
                oUserTablesMD.TableType = nType
                If oUserTablesMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
            oUserTablesMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Private Sub AddFields(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO)
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            If Not (strTab = "OUSR" Or strTab = "OCRD" Or strTab = "OJDT" Or strTab = "OINV" Or strTab = "OCRN" Or strTab = "OHEM" Or strTab = "INV1" Or strTab = "RDR1" Or strTab = "OINV" Or strTab = "ORDR" Or strTab = "JDT1") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                oUserFieldMD.Description = strDesc
                oUserFieldMD.Name = strCol
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory
                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If



            If (Not IsColumnExists(TableName, ColumnName)) Then
                objUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    '  MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            Else


            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

            objUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try


    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : IsColumnExists
    'Parameter          : ByVal Table As String, ByVal Column As String
    'Return Value       : Boolean
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Function to check if the Column already exists in Table
    '**************************************************************************************************************
    Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & Table & "' AND AliasID = '" & Column & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strSQL)

            If oRecordSet.Fields.Item(0).Value = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
        End Try
    End Function

    Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        Try
            '// The meta-data object must be initialized with a
            '// regular UserKeys object
            oUserKeysMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

            If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                '// Set the table name and the key name
                oUserKeysMD.TableName = strTab
                oUserKeysMD.KeyName = strKey

                '// Set the column's alias
                oUserKeysMD.Elements.ColumnAlias = strColumn
                oUserKeysMD.Elements.Add()
                oUserKeysMD.Elements.ColumnAlias = "RentFac"

                '// Determine whether the key is unique or not
                oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                '// Add the key
                If oUserKeysMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
            oUserKeysMD = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    '********************************************************************
    'Type		            :   Function    
    'Name               	:	AddUDO
    'Parameter          	:   
    'Return Value       	:	Boolean
    'Author             	:	
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To Add a UDO for Transaction Tables
    '********************************************************************
    Private Sub AddUDO(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal strChildTbl As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document, Optional ByVal strChildTb2 As String = "")

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.LogTableName = ""
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If

                If strChildTb2 <> "" Then
                    oUserObjectMD.ChildTables.Add()
                    oUserObjectMD.ChildTables.TableName = strChildTb2
                End If

                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub

#End Region

#Region "Public Functions"
    '*************************************************************************************************************
    'Type               : Public Function
    'Name               : CreateTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Creating Tables by calling the AddTables & AddFields Functions
    '**************************************************************************************************************
    Public Sub CreateTables()
        Dim oProgressBar As SAPbouiCOM.ProgressBar
        Try

            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            addField("OJDT", "Export", "Exported", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            '---- User Defined Object's
            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            AddTables("Z_SQLDETAIL", "GFS SQL Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_SQLDETAIL", "Z_Server", "Server Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_SQLDETAIL", "Z_DBName", "Database Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_SQLDETAIL", "Z_UID", "User ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_SQLDETAIL", "Z_Pwd", "Password", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
          
            AddFields("OJDT", "Z_IFXReply", "IFX Response", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            AddFields("OJDT", "Export", "Export", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, 1)

            AddTables("Z_IFXSetup", "IFX Setup", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_IFXSetup", "Z_URL", "URL", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_IFXSetup", "Z_UID", "User ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_IFXSetup", "Z_PWD", "Password", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_IFXSetup", "Z_JERemarks", "SAP Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_IFXSetup", "Z_XMLPath", "XML File Path", SAPbobsCOM.BoFieldTypes.db_Memo)

            AddFields("OUSR", "PhxId", "Phoneix ID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            AddFields("OCRN", "Z_Acct", "FC Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("OCRN", "Z_Acct1", "Local Currency Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            addField("@Z_PAYROLL1", "Z_IFXPosting", "IFX Posting Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "C,P", "Posted,Pending", "P")
            AddFields("Z_PAYROLL1", "Z_IFXResponse", "IFX Posting Response", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)


            addField("@Z_PAYROLL1", "Z_IFXPAY", "IFX Emp.Pay Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "C,P", "Posted,Pending", "P")
            AddFields("Z_PAYROLL1", "Z_IFXEmpRes", "IFX Emp.Posting Response", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)

            addField("@Z_PAYROLL1", "Z_IFXDebit", "IFX Debit.Pay Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "C,P", "Posted,Pending", "P")

            addField("OUSR", "FApprover", "First Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("OUSR", "SApprover", "Second Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("OUSR", "TApprover", "Third Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            addField("OBTF", "FAproval", "First Approved", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,A,R", "Open,Approved,Rejected", "O")
            addField("OBTF", "FApprTime", "First Approval Time", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("OBTF", "FAppRmks", "First Approval Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            addField("OBTF", "SAproval", "Second Approved", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,A,R", "Open,Approved,Rejected", "O")
            addField("OBTF", "SApprTime", "Second Approval Time", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("OBTF", "SAppRmks", "Second Approval Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            addField("OBTF", "TAproval", "Third Approved", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,A,R", "Open,Approved,Rejected", "O")
            addField("OBTF", "TApprTime", "Third Approval Time", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("OBTF", "TAppRmks", "Third Approval Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            addField("OBTF", "FoAproval", "Fourth Approved", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,A,R", "Open,Approved,Rejected", "O")
            addField("OBTF", "FoApprTime", "Fourth Approval Time", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("OBTF", "FoAppRmks", "Fourth Approval Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            addField("OBTF", "Approval", "Final Approved", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,A,R", "Open,Approved,Rejected", "O")

            addField("OBTF", "FApprover", "First Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("OBTF", "SApprover", "Second Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("OBTF", "TApprover", "Third Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
            addField("OBTF", "FoApprover", "Fourth Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")

            AddTables("OJAT", "JV Approval Template", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddTables("JAT1", "JV Approval - Originator", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            AddFields("JAT1", "OUser", "Origin - User", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            AddFields("JAT1", "OName", "Origin - Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 155)
            AddTables("JAT2", "JV Approval - Authorizer", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
            AddFields("JAT2", "AUser", "Authorizer - User", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            AddFields("JAT2", "AName", "Authorizer - Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 155)

            AddFields("JDT1", "Z_FCCRAc", "FC Credit A/C", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("JDT1", "Z_FCDRAc", "FC Debit A/C", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("JDT1", "Z_FCCRCUR", "FC Credit Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)
            AddFields("JDT1", "Z_FCDRCUR", "FC Debit A/C", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)
            AddFields("JDT1", "Z_FCCRAmt", "FC Credit Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("JDT1", "Z_FCDRAm", "FC Debit Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)

            AddFields("OJDT", "Z_AttRef", "Attachment Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddTables("Z_OJVATT", "Journal Attachment Master", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddTables("Z_JVATT", "Journal Attachment Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_JVATT", "Z_RefCode", "Reference Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_JVATT", "Z_FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            AddFields("Z_JVATT", "Z_FilePath", "File Path", SAPbobsCOM.BoFieldTypes.db_Memo)

            AddFields("Z_JVATT", "Z_CreatedBy", "Created By", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            AddFields("Z_JVATT", "Z_CreateDate", "Create Date", SAPbobsCOM.BoFieldTypes.db_Date)

            CreateUDO()
            oApplication.Utilities.Message("Initializing Database Completed...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Function UDOSetup(ByVal strUDO As String, _
                    ByVal strDesc As String, _
                        ByVal strTable As String, _
                            ByVal intFind As Integer, _
                                Optional ByVal strCode As String = "", _
                                    Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""
                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()

                oUserObjects.FormColumns.FormColumnAlias = "U_Z_BraDB"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_BraDB"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_BraSQL"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_BraSQL"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_SQLUID"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_SQLUID"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_SQLPWD"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_SQLPWD"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_SAPUID"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_SAPUID"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_SAPPWD"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_SAPPWD"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                    ' oUserObjects.FindColumns.Add()
                    'oUserObjects.FindColumns.SetCurrentLine(1)
                    'oUserObjects.FindColumns.ColumnAlias = strName
                    'oUserObjects.FindColumns.Add()
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function

    Public Sub CreateUDO()
        Try
            ' UDOSetup("Z_AL_OADM", "Branch DB Setups", "Z_AL_OADM", 1, "U_Z_BraDB", )
            AddUDO("UDO_OJAT", "JV Approval Template", "OJAT", "Code", "Name", "JAT1", SAPbobsCOM.BoUDOObjType.boud_MasterData, "JAT2")
             Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

End Class
