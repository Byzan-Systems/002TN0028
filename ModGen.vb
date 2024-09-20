Option Explicit On

Module ModGen


    Public blnErrorLog As Boolean
    Public strAuditFolderPath As String
    Public strErrorFolderPath As String

    Public strInputFolderPath As String
    Public gstrInputFile As String
    Public gstrInputFolder As String

    Public gstrOutputFile_Name As String

    Public strOutputFolderPath As String

    ''Res
    Public gstrResOutputfile As String
    Public gstrResponseInputFolder As String
    Public gstrResponseInputFile As String

    Public strResponseFolderPath As String             ' Response folder path
    Public strReverseResponseFolderPath As String            ' RevResponse folder path

    ''Archive
    Public strArchivedFolderSuc As String
    Public strArchivedFolderUnSuc As String
    ''''''''''''''''''
    Public strReportFolderPath As String

    Public strProceed As String
    Public strInvalidTrans As String
    Public FileCounter As String

    Public strValidationPath As String
    Public strTransactionNo As String

    '-Client Details-
    Public strClientCode As String
    Public strClientName As String
    Public strInputDateFormat As String

    Public strFileFormat As String
     
End Module


