Public Module modVariables

    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String
    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public frmSourceFormUD As String
    Public frmSourceForm As SAPbouiCOM.Form
    Public frmSourcePMForm As SAPbouiCOM.Form
    Public frmSourceQCOR As SAPbouiCOM.Form

    Public LoalDB As String
    Public intCurrentRow As Integer = 10000


    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public strCardCode As String = ""
    Public blnDraft As Boolean = False
    Public blnError As Boolean = False
    Public strDocEntry As String
    Public strImportErrorLog As String = ""
    Public companyStorekey As String = ""

    Public intSelectedMatrixrow As Integer = 0
    Public strSourceformEmpID As String = ""
    Public strApprovalType As String = ""
    Public oAssetTransactionNumber As String = ""

    Public strMdbFilePath As String
    Dim strFileName As String
    Public strSelectedFilepath, sPath, strSelectedFolderPath, strFilepath As String

    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Enum DocumentType As Integer
        RENTAL_QUOTATION = 1
        RENTAL_ORDER
        RENTAL_RETURN
    End Enum
    Public Const frm_GFCSetup As String = "frm_GFCSetup"
    Public Const mnu_GFCSetup As String = "Z_mnu_IFX103"
    Public Const xml_GFCSetup As String = "xml_GFCSetup.xml"

    Public Const frm_JournalVoucher As String = "393"
    Public Const frm_JournalEntry As String = "392"

    Public Const frm_Attachment As String = "frm_Attach"
    Public Const xml_Attachment As String = "xml_Attachment.xml"

    Public Const mnu_IFXPosting As String = "Z_mnu_IFX002"
    Public Const frm_IFXPosting As String = "frm_IFXPosting"
    Public Const xml_IFXPosting As String = "xml_IFXPosting.xml"

    Public Const frm_IFXSetup As String = "frm_IFXSetup"
    Public Const mnu_IFXSetup As String = "Z_mnu_IFX003"
    Public Const xml_IFXSetup As String = "xml_IFXSetup.xml"


    Public Const frm_JVPostings As String = "229"
    Public Const mnu_JVApproval As String = "Z_mnu_JVApp"
    Public Const frm_JVApproval As String = "frm_JVApproval"
    Public Const xml_JVApproval As String = "xml_JVApproval.xml"

    Public Const frm_OJAT As String = "frm_OJAT"
    Public Const mnu_OJAT As String = "Z_mnu_OJAT"
    Public Const xml_OJAT As String = "xml_OJAT.xml"
  
    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"
    Public Const mnu_DuplicateRow As String = "1294"

    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"
   

End Module
