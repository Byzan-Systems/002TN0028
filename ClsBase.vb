Imports System
Imports System.Data
Imports System.IO

Public Class ClsBase
    Inherits ClsShared

    Public Sub New(ByVal _StrIniPath As String)

        Try
            gstrIniPath = _StrIniPath

            'General
            strAuditFolderPath = GetINISettings("General", "Audit Log", _StrIniPath)
            strErrorFolderPath = GetINISettings("General", "Error Log", _StrIniPath)

            strInputFolderPath = GetINISettings("General", "Input Folder", _StrIniPath)
            strOutputFolderPath = GetINISettings("General", "Output Folder", _StrIniPath)
            strReportFolderPath = GetINISettings("General", "Report Folder", _StrIniPath)

            strResponseFolderPath = GetINISettings("General", "Response Folder", _StrIniPath)
            strReverseResponseFolderPath = GetINISettings("General", "RevResponse Folder", _StrIniPath)

            strValidationPath = GetINISettings("General", "Validation", _StrIniPath)
            strArchivedFolderSuc = GetINISettings("General", "Archived FolderSuc", _StrIniPath)
            strArchivedFolderUnSuc = GetINISettings("General", "Archived FolderUnSuc", _StrIniPath)

            strProceed = GetINISettings("General", "Process Output File Ignoring Invalid Transactions", _StrIniPath)

            strClientName = GetINISettings("Client Details", "Client Name", _StrIniPath)
            strClientCode = GetINISettings("Client Details", "Client Code", _StrIniPath)
            strInputDateFormat = GetINISettings("Client Details", "Input Date Format", _StrIniPath)

            strTransactionNo = GetINISettings("Client Details", "Number of Records in Per Output File", _StrIniPath)

            Reset_Counter(_StrIniPath)

        Catch ex As Exception

        End Try

    End Sub
    Public Function FileDelete(ByVal SourceFilePath As String) As Boolean

        Try
            If File.Exists(SourceFilePath) Then
                File.Delete(SourceFilePath)
            End If

            FileDelete = True

        Catch ex As Exception
            FileDelete = False
            Call Handle_Error(ex, "ClsBase", "FileDelete - Source File =" & SourceFilePath)
        End Try

    End Function
    Public Function Reset_Counter(ByVal _StrIniPath1 As String)

        Try
            Dim strSettingsdate As String, dtsettings As Date
            Dim intresult As Integer

            strSettingsdate = (GetINISettings("General", "Date", _StrIniPath1))
            GetValidateSettingDate(strSettingsdate)
            dtsettings = strSettingsdate
            intresult = DetermineNumberofDays(dtsettings)
            If intresult > 0 Then
                Call SetINISettings("General", "Date", Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year, _StrIniPath1)
                Call SetINISettings("General", "File Counter", 0, _StrIniPath1)
                LogEntry("[Counter Reseted]")
            End If

        Catch ex As Exception
            Call Me.Handle_Error(ex, "ClsBase", Err.Number, "Reset_Counter")

        End Try

    End Function

    Private Function GetValidateSettingDate(ByRef pStrDate As String) As Boolean

        Try

            Dim striniDate As String
            striniDate = "DD/MM/YYYY"

            Dim TempStrDateValue() As String = pStrDate.Split(" ")

            If InStr(pStrDate, "/") > 0 Then
                TempStrDateValue = TempStrDateValue(0).Split("/")
            ElseIf InStr(pStrDate, "-") > 0 Then
                TempStrDateValue = TempStrDateValue(0).Split("-")
            End If



            
            Dim TmpstrInputDateFormat() As String = striniDate.Split("/")

            Dim HsUserDate As New Hashtable
            Dim HsSystemDate As New Hashtable
            Dim StrFinalDate As String

            If TempStrDateValue.Length = 3 Then
                For IntStr As Integer = 0 To TempStrDateValue.Length - 1
                    HsUserDate.Add(GetShortINI(TmpstrInputDateFormat(IntStr)), TempStrDateValue(IntStr))
                Next


                Dim SysDate() As String
                Dim dtSys As String = System.Globalization.DateTimeFormatInfo.CurrentInfo.ShortDatePattern.ToUpper()
                If InStr(dtSys, "/") > 0 Then
                    SysDate = dtSys.Split("/")
                ElseIf InStr(dtSys, "-") > 0 Then
                    SysDate = dtSys.Split("-")
                End If

                StrFinalDate = ""

                For IntStr As Integer = 0 To SysDate.Length - 1
                    If StrFinalDate = "" Then
                        StrFinalDate += HsUserDate(GetShortINI(SysDate(IntStr))).ToString().Trim()
                    Else
                        StrFinalDate += "/" & HsUserDate(GetShortINI(SysDate(IntStr))).ToString().Trim()
                    End If

                Next

                Try
                    pStrDate = CDate(StrFinalDate)
                    GetValidateSettingDate = True
                Catch ex As Exception
                    GetValidateSettingDate = False
                End Try
            Else
                GetValidateSettingDate = False
            End If

        Catch ex As Exception
            GetValidateSettingDate = False

        End Try

    End Function

    Private Function GetShortINI(ByVal pStr As String) As String

        pStr = pStr.ToUpper

        If InStr(pStr, "D") > 0 Then
            GetShortINI = "D"
        ElseIf InStr(pStr, "M") > 0 Then
            GetShortINI = "M"
        ElseIf InStr(pStr, "Y") > 0 Then
            GetShortINI = "Y"
        End If

    End Function

    Private Function DetermineNumberofDays(ByVal dtStartDate As Date) As Integer

        Dim tsTimeSpan As TimeSpan
        Dim iNumberOfDays As Integer

        tsTimeSpan = Now.Subtract(dtStartDate)
        iNumberOfDays = tsTimeSpan.Days
        DetermineNumberofDays = iNumberOfDays

    End Function


    Public Function openConn_String_XL_UpdateMaster(ByVal sFileName As String) As String

        ' if connection is aready open then close and re-open the same connection
        Try
            Dim strConn As String
            If Right(sFileName, 4).ToUpper() = ".XLS" Then
                ' strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sFileName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'"
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sFileName + ";Extended Properties='Excel 8.0;HDR=Yes;'"
            ElseIf Right(sFileName, 5).ToUpper() = ".XLSX" Then
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFileName + ";Extended Properties='Excel 12.0 Xml;HDR=Yes;IMEX=1'"
            Else
                MessageBox.Show("Please install office 2007 or above." & vbCrLf & "  XLSX file format is not supported.", "Generic RBI Converter")
                End
            End If

            'strConn = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & sFileName & ";Extended Properties='Excel 8.0;IMEX=1';"

            openConn_String_XL_UpdateMaster = strConn

        Catch ex As Exception

            openConn_String_XL_UpdateMaster = "Error"
            Call Me.Handle_Error(ex, "clsBase", Err.Number, "openConn_String_XL")
        Finally

        End Try
    End Function
    
    Public Sub WriteAuditLogFile(ByVal pDesc As String)

        Dim StrFileName As String
        Dim obOpenFile As FileStream
        Dim SwOpenFile As StreamWriter

        Dim strAuditLogPath As String
        Dim Strheading As String

        Try

            strAuditLogPath = strAuditFolderPath

            strAuditLogPath = padSlash(strAuditLogPath)

            StrFileName = strAuditLogPath & "Log" & Today.Day & Today.Month & Today.Year & ".log"

            'check for the existence of the text file
            If File.Exists(StrFileName) Then
                obOpenFile = New FileStream(StrFileName, FileMode.Append, FileAccess.Write, FileShare.Write)
                Strheading = ""
                SwOpenFile = New StreamWriter(obOpenFile)
            Else
                obOpenFile = New FileStream(StrFileName, FileMode.Create, FileAccess.Write, FileShare.Write)
                ' Strheading = "AutoName | Frequency | ActionDescription | start Date and Time "
                SwOpenFile = New StreamWriter(obOpenFile)
                ' SwOpenFile.WriteLine(Strheading)
            End If
            SwOpenFile.WriteLine(pDesc)
            ObjectFlush(obOpenFile)
            ObjectDispose(SwOpenFile)
            StrFileName = ""

        Catch ex As Exception

            Call Handle_Error(ex, "ClsBase", Err.Number, "WriteAuditLogFile")

        Finally
            ObjectFlush(obOpenFile)
            ObjectDispose(SwOpenFile)
        End Try

    End Sub

    Public Function openConn_String_XL(ByVal sFileName As String) As String

        ' if connection is aready open then close and re-open the same connection
        Try
            Dim strConn As String
            'If Right(sFileName, 4).ToUpper() = ".XLS" Then
            '    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sFileName + ";Extended Properties='Excel 8.0;IMEX=1'"
            'ElseIf Right(sFileName, 5).ToUpper() = ".XLSX" Then
            '    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFileName + ";Extended Properties='Excel 12.0 Xml;HDR=Yes;IMEX=1'"
            'Else
            '    MessageBox.Show("Please install office 2007 or above." & vbCrLf & "  XLSX file format is not supported.", "Generic RBI Converter")
            '    End
            'End If

            'strConn = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & sFileName & ";Extended Properties='Excel 8.0;IMEX=1';"

            strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFileName + ";Extended Properties='Excel 12.0 Xml;HDR=Yes;IMEX=1'"

            openConn_String_XL = strConn

        Catch ex As Exception

            openConn_String_XL = "Error"
            Call Me.Handle_Error(ex, "clsBase", Err.Number, "openConn_String_XL")
        Finally

        End Try
    End Function

   Public Function GetDatatable_Text(ByVal StrFilePath As String) As DataTable

        Dim strTemp() As String
        Dim TmpLineStr As String
        Dim DtInput As DataTable
        Dim strReader As New StreamReader(StrFilePath)

        Try

            Do While strReader.EndOfStream = False

                TmpLineStr = strReader.ReadLine

                'strTemp = GetInArrayByComma(TmpLineStr) 'TmpLineStr.Split("@")
                strTemp = TmpLineStr.Split(",")
                AddColumnToTable(DtInput, strTemp.Length)
                DtInput.Rows.Add(strTemp)

            Loop

            GetDatatable_Text = DtInput.Copy

        Catch ex As Exception

        Finally
            If Not strReader Is Nothing Then
                strReader.Close()
                strReader.Dispose()
            End If
            strReader = Nothing

            If Not DtInput Is Nothing Then
                DtInput.Dispose()
            End If
            DtInput = Nothing

        End Try

    End Function

    Public Function MyGetDatatable_Text(ByVal StrFilePath As String, ByVal Parameter As String) As DataTable

        Dim strTemp() As String
        Dim TmpLineStr As String
        Dim DtInput As New DataTable
        Dim strReader As New StreamReader(StrFilePath)


        Try

            Do While strReader.EndOfStream = False
                TmpLineStr = strReader.ReadLine

                strTemp = GetInArrayByComma(TmpLineStr, Parameter) 'TmpLineStr.Split("@")
                AddColumnToTable(DtInput, strTemp.Length)
                DtInput.Rows.Add(strTemp)
                ' End If

            Loop

            MyGetDatatable_Text = DtInput.Copy

        Catch ex As Exception

        Finally
            If Not strReader Is Nothing Then
                strReader.Close()
                strReader.Dispose()
            End If
            strReader = Nothing

            If Not DtInput Is Nothing Then
                DtInput.Dispose()
            End If
            DtInput = Nothing

        End Try

    End Function

    Private Function GetInArrayByComma(ByVal pStrValue As String, ByVal strDelimiter As String) As String()

        Try

            Dim Tmpstr As String = ""
            Dim Index_S, Index_E, TmpIndex As Integer

            Index_E = InStr(pStrValue, Chr(34))

            If Index_E > 0 Then

                Index_S = 0
                Tmpstr = ""
                While True

                    Index_E = InStr(Index_S + 1, pStrValue, Chr(34))

                    If Index_E > 0 Then

                        Tmpstr += pStrValue.Substring(Index_S, Index_E - Index_S - 1).Replace(strDelimiter, "|")
                        Index_S = Index_E
                        Index_E = InStr(Index_E + 1, pStrValue, Chr(34))
                        Tmpstr += pStrValue.Substring(Index_S, (Index_E - Index_S) - 1)
                        Index_S = Index_E

                    Else

                        Tmpstr += pStrValue.Substring(Index_S, pStrValue.Length - Index_S).Replace(strDelimiter, "|")
                        GetInArrayByComma = Tmpstr.Split("|")
                        Exit While

                    End If

                End While
            Else
                GetInArrayByComma = pStrValue.Split(strDelimiter)
            End If

        Catch ex As Exception
            Call Handle_Error(ex, "ClsBase", Err.Number, "GetInArrayByComma")
        End Try

    End Function
    

    Public Sub AddColumnToTable(ByRef pDt As DataTable, ByVal pCols As Integer)

        If pDt Is Nothing Then
            pDt = New DataTable("Input")
        End If

        If pDt.Columns.Count < pCols Then
            pDt.Columns.Add(New DataColumn("Column_" & pDt.Columns.Count))
            AddColumnToTable(pDt, pCols)
        End If

    End Sub

    Public Function Execute_Batch_file(ByVal tempstrBatchFilePath As String) As Boolean

        Dim batchExecute As New Process
        Dim batchExecuteInfo As New ProcessStartInfo(tempstrBatchFilePath & "\" & "Test.bat")


        Try

            batchExecuteInfo.UseShellExecute = True
            batchExecuteInfo.CreateNoWindow = False
            batchExecute.StartInfo = batchExecuteInfo
            batchExecute.Start()
            batchExecute.WaitForExit(20000)


            Execute_Batch_file = True


        Catch ex As Exception

        End Try

    End Function


    Public Function isCompleteFileAvailable(ByVal szFilePath As String) As Boolean

        Dim fsObj As FileStream
        Dim obOpenFile As StreamWriter
        Try

            While True
                Try
                    If File.Exists(szFilePath) Then
                        fsObj = New FileStream(szFilePath, FileMode.Append, FileAccess.Write, FileShare.None)
                        obOpenFile = New StreamWriter(fsObj)
                        isCompleteFileAvailable = True
                    Else
                        isCompleteFileAvailable = False
                        Exit While
                    End If
                Catch ex As Exception
                    isCompleteFileAvailable = False
                    Threading.Thread.Sleep(1000)
                Finally
                    If Not fsObj Is Nothing Then fsObj.Flush()
                    If Not obOpenFile Is Nothing Then obOpenFile.Dispose()
                    fsObj = Nothing
                    obOpenFile = Nothing
                End Try
                If isCompleteFileAvailable = True Then Exit While
            End While

        Catch ex As Exception

            Call Me.Handle_Error(ex, "ClsBase", Err.Number, "isCompleteFileAvailable")
        End Try

    End Function

    Public Function FileMove(ByVal SourceFilePath As String, ByVal DestinFilePath As String) As Boolean

        Try
            If File.Exists(SourceFilePath) Then
                If File.Exists(DestinFilePath) Then
                    File.Delete(DestinFilePath)
                End If
                File.Move(SourceFilePath, DestinFilePath)
            End If
            FileMove = True

        Catch ex As Exception

            FileMove = False
            Call Handle_Error(ex, "ClsBase", Err.Number, "FileMove : " & "Source File =" & SourceFilePath & "Destination File =" & DestinFilePath)
        End Try

    End Function
    Public Function FileCopy(ByVal SourceFilePath As String, ByVal DestinFilePath As String) As Boolean

        Try
            If File.Exists(SourceFilePath) Then
                If File.Exists(DestinFilePath) Then
                    File.Delete(DestinFilePath)
                End If
                File.Copy(SourceFilePath, DestinFilePath)
            End If
            FileCopy = True

        Catch ex As Exception

            FileCopy = False
            Call Handle_Error(ex, "ClsBase", Err.Number, "FileCopy : " & "Source File =" & SourceFilePath & "Destination File =" & DestinFilePath)
        End Try

    End Function

    Public Function GetDataTable_DistinctColoumData(ByVal fileName As String, ByVal sheetName As String) As DataTable
        Dim conn As System.Data.OleDb.OleDbConnection
        Dim dataResult As New DataTable
        'Dim DtWithoutblank As New DataTable

        Try
            'conn = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=Excel 8.0;")
            conn = New System.Data.OleDb.OleDbConnection(openConn_String_XL(fileName))
            conn.Open()

            Dim command As New System.Data.OleDb.OleDbCommand(" SELECT [Advisor Code],sum([Net Amount]) as [Net Amount] FROM [" + sheetName + "$] group by [Advisor Code]")
            command.Connection = conn
            Dim adaperForExcelBook As New System.Data.OleDb.OleDbDataAdapter
            adaperForExcelBook.SelectCommand = command
            adaperForExcelBook.Fill(dataResult)
            conn.Close()

            GetDataTable_DistinctColoumData = dataResult.Copy()

            ObjectDispose(adaperForExcelBook)
            ObjectDispose(dataResult)
            '  ObjectDispose(DtWithoutblank)
            ObjectDispose(conn)

        Catch ex As Exception

            Call Handle_Error(ex, "ClsBase", Err.Number, "GetDataTable_DistinctColoumData")
        End Try

    End Function

    Public Function GetDataTable_ExcelSQL(ByVal FilePathName As String, ByVal IntSheetNo As Integer, ByVal StrSQLFilterOrder As String) As DataTable
        ''Added on dtd 31-03-2011
        Dim conn As System.Data.OleDb.OleDbConnection
        Dim dataResult As New DataTable
        Dim StrSheetName(0) As String

        Try
            conn = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FilePathName + ";Extended Properties='Excel 8.0;IMEX=1'")
            conn.Open()
            Dim dt As DataTable = conn.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, Nothing)

            If Not dt Is Nothing Then

                If IntSheetNo > dt.Rows.Count Then
                    IntSheetNo = dt.Rows.Count
                ElseIf IntSheetNo < 0 Then
                    IntSheetNo = 1
                End If
                For Each Dr As DataRow In dt.Rows
                    ReDim Preserve StrSheetName(UBound(StrSheetName) + 1)
                    StrSheetName(UBound(StrSheetName) - 1) = Dr("TABLE_NAME").ToString()
                Next
            Else
                Throw New ApplicationException(FilePathName & " Excel file content 0 sheet")
            End If

            Dim command As New System.Data.OleDb.OleDbCommand("Select * from [" & StrSheetName(IntSheetNo - 1) & "] " & StrSQLFilterOrder)
            ''command.CommandTimeout = Gstrcomtimeout
            command.Connection = conn
            Dim adaperForExcelBook As New System.Data.OleDb.OleDbDataAdapter
            adaperForExcelBook.SelectCommand = command
            adaperForExcelBook.Fill(dataResult)
            conn.Close()

            GetDataTable_ExcelSQL = dataResult.Copy()

            ObjectDispose(adaperForExcelBook)
            ObjectDispose(dataResult)
            ObjectDispose(conn)

        Catch ex As Exception
            Call Handle_Error(ex, "ClsBase", Err.Number, "GetDataTable_ExcelSQL")

        End Try
    End Function

    Public Function GetDataTable_ExcelSheet(ByVal fileName As String, ByVal sheetName As String, Optional ByVal Filter As String = "") As DataTable

        Dim conn As System.Data.OleDb.OleDbConnection
        Dim dataResult As New DataTable
        Dim sheetName1 As String
        Try
            conn = New System.Data.OleDb.OleDbConnection(openConn_String_XL(fileName))
            conn.Open()

            Dim dt As DataTable = conn.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, Nothing)
            Dim dr As DataRow = dt.Rows(0)
            sheetName1 = dr(2).ToString()
            '----------------end here
            Dim command As New System.Data.OleDb.OleDbCommand("SELECT * FROM [" + sheetName + "$]")
            If sheetName = "" Then
                command = New System.Data.OleDb.OleDbCommand("SELECT * FROM [" + sheetName1 + "]")
            End If



            'Dim command As New System.Data.OleDb.OleDbCommand("SELECT * FROM [" + sheetName + "$]")
            command.Connection = conn
            Dim adaperForExcelBook As New System.Data.OleDb.OleDbDataAdapter
            adaperForExcelBook.SelectCommand = command
            adaperForExcelBook.Fill(dataResult)
            conn.Close()

            GetDataTable_ExcelSheet = dataResult.Copy()

            ObjectDispose(adaperForExcelBook)

        Catch ex As Exception
            Call Handle_Error(ex, "ClsBase", Err.Number, "GetDataTable_ExcelSheet")

        Finally
            'ObjectDispose(adaperForExcelBook)
            ObjectDispose(dataResult)
            ObjectDispose(conn)
        End Try

    End Function

    Public Function GetDataTable_ExcelSheet_Head(ByVal fileName As String, ByVal sheetName As String, Optional ByVal Filter As String = "") As DataTable
        Dim conn As System.Data.OleDb.OleDbConnection
        Dim dataResult As New DataTable
        'Dim DtWithoutblank As New DataTable

        Try

            conn = New System.Data.OleDb.OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & fileName & ";Extended Properties='Excel 8.0;IMEX=1;HDR=No;';")
            'conn = New System.Data.OleDb.OleDbConnection(openConn_String_XL(fileName))
            conn.Open()

            Dim command As New System.Data.OleDb.OleDbCommand("SELECT * FROM [" + sheetName + "$]")
            command.Connection = conn
            Dim adaperForExcelBook As New System.Data.OleDb.OleDbDataAdapter
            adaperForExcelBook.SelectCommand = command
            adaperForExcelBook.Fill(dataResult)
            conn.Close()

            GetDataTable_ExcelSheet_Head = dataResult.Copy()

            ObjectDispose(adaperForExcelBook)

        Catch ex As Exception

            Call Handle_Error(ex, "ClsBase", Err.Number, "GetDataTable_ExcelSheet_Head")
        Finally
            'ObjectDispose(adaperForExcelBook)
            ObjectDispose(dataResult)
            'ObjectDispose(DtWithoutblank)
            ObjectDispose(conn)
        End Try

    End Function

    Public Function GetDataTable_ExcelSheetInput(ByVal fileName As String, ByVal sheetName As String, Optional ByVal Filter As String = "") As DataTable

        Dim conn As System.Data.OleDb.OleDbConnection
        Dim dataResult As New DataTable
        Dim dt As DataTable

        Try
            conn = New System.Data.OleDb.OleDbConnection(openConn_String_XLInput(fileName))
            conn.Open()

            Dim myTableName = conn.GetSchema("Tables").Rows(0)("TABLE_NAME")
            sheetName = myTableName
            Dim command As New System.Data.OleDb.OleDbCommand("SELECT * FROM [" + sheetName + "]")
            command.Connection = conn
            Dim adaperForExcelBook As New System.Data.OleDb.OleDbDataAdapter
            adaperForExcelBook.SelectCommand = command
            adaperForExcelBook.Fill(dataResult)
            conn.Close()

            GetDataTable_ExcelSheetInput = dataResult.Copy()

            ObjectDispose(adaperForExcelBook)

        Catch ex As Exception
            Call Handle_Error(ex, "ClsBase", Err.Number, "GetDataTable_ExcelSheetInput")

        Finally
            'ObjectDispose(adaperForExcelBook)
            ObjectDispose(dataResult)
            ObjectDispose(conn)
        End Try

    End Function
    Public Function openConn_String_XLInput(ByVal sFileName As String) As String

        ' if connection is aready open then close and re-open the same connection
        Try
            Dim strConn As String
            strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sFileName + ";Extended Properties='Excel 12.0 Xml;HDR=No;IMEX=1'"

            openConn_String_XLInput = strConn

        Catch ex As Exception

            openConn_String_XLInput = "Error"
            Call Me.Handle_Error(ex, "clsBase", Err.Number, "openConn_String_XL")
        Finally

        End Try
    End Function
    Public Function funcGetRange(ByVal intCol As Integer) As String
        Dim intReminder As Integer, intValue As Integer
        Try
            intCol += 1
            intReminder = intCol Mod 26
            intValue = intCol / 26
            If intCol > 26 Then
                funcGetRange = Chr(64 + intValue) & Chr(64 + intReminder)
            Else
                funcGetRange = Chr(64 + intCol)
            End If

        Catch ex As Exception
            Call Handle_Error(ex, "ClsBase", Err.Number, "funcGetRange")

        End Try
    End Function

    Public Sub ObjectDispose(ByRef Obj As Object)
        Try
            If Not Obj Is Nothing Then
                Try
                    Obj.close()
                Catch ex As Exception
                    ' Debug.Print("Error")
                End Try
                Obj.dispose()
                Obj = Nothing
            End If
        Catch ex As Exception
            Obj = Nothing
        Finally
            GC.Collect()
        End Try

    End Sub

    Public Sub ObjectFlush(ByRef Obj As Object)
        Try
            If Not Obj Is Nothing Then
                Obj.flush()
                Obj = Nothing
            End If
        Catch ex As Exception
            Obj = Nothing
        Finally
            GC.Collect()
        End Try

    End Sub

    Public Sub ObjectDispose_Excel(ByRef obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
            GC.RemoveMemoryPressure(GC.MaxGeneration)

        End Try
    End Sub

    Public Overloads Sub Dispose()
        Me.Finalize()
        MyBase.Dispose()
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

Public Class ClsShared
    Inherits ClsErrLog

#Region " API Decalration"

    '----API Declaration
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    ' for copy paste operations
    Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Integer) As Integer
    ' Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

    '------------------------------
    Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer
    Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
    Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer

    Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Object, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
    Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Object, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As Object, ByVal lpvSource As Object, ByVal cbCopy As Long)

    Public Const INFINITE As Short = -1
    Public Const SYNCHRONIZE As Integer = &H100000
    '------------------------------

    Public Connstring As String
    Public gstrErrLogPath As String
    Public gstrIniPath As String
    Public ProcDelimeter As Char

    Protected ConnDBF As System.Data.OleDb.OleDbConnection

#End Region

    Public Function padSlash(ByRef szfPath As String) As String

        Try
            If Right(szfPath, 1) <> "\" Then szfPath = szfPath & "\"
            padSlash = Trim(szfPath)

        Catch ex As Exception

            padSlash = szfPath
            Call Handle_Error(ex, "ClsBase", Err.Number, "padSlash")
        Finally

        End Try

    End Function

    Public Overloads Sub Dispose()
        Me.Finalize()
        MyBase.Dispose()
        GC.SuppressFinalize(Me)
    End Sub

    Public Function SetINISettings(ByVal sectionName As String, ByVal strkeyName As String, ByVal strkeyValue As String, ByVal appPath As String) As Boolean

        Try
            Dim lgStatus As Integer

            lgStatus = WritePrivateProfileString(sectionName, strkeyName, strkeyValue, appPath)
            If lgStatus = 0 Then
                SetINISettings = False
            Else
                SetINISettings = True
            End If

        Catch ex As Exception

            SetINISettings = False
            Call Me.Handle_Error(ex, "ClsBase", "SettINISettings", sectionName & ", " & strkeyName & ", " & strkeyValue & ", " & appPath)
        Finally

        End Try

    End Function

    Public Function GetINISettings(ByVal sHeader As String, ByVal sKey As String, ByVal sININame As String) As String

        Dim iRetval As Short
        Dim lpBuffer As String   ' New VB6.FixedLengthString(255)
        Dim sRetval As String

        Try

            lpBuffer = ""
            For i As Int16 = 1 To 255
                lpBuffer = lpBuffer & Chr(16) ''"" ''Chr(0)
            Next

            iRetval = GetPrivateProfileString(sHeader, sKey, "", lpBuffer, 255, sININame)
            sRetval = Left(lpBuffer, iRetval)
            GetINISettings = sRetval

        Catch ex As Exception

            GetINISettings = ""
            Call Handle_Error(ex, "ClsBase", Err.Number, "GetINISettings")
        Finally

        End Try

    End Function



End Class

Public Class ClsErrLog
    Implements IDisposable

    Public Sub Handle_Error(ByVal oErr As Exception, ByVal strFormName As String, ByVal errno As Int64, Optional ByVal strFunctionName As String = "")
        Try

            WriteErrorToTxtFile(Err.Number, oErr.Message, strFormName, strFunctionName) ', strEnvtVars)

        Catch ex As Exception

        End Try
    End Sub

    Public Sub LogEntry(ByVal StrMessage As String, Optional ByVal IsError As Boolean = False)

        Try

            Dim LogPath As String
            Dim LogFileName As String
            StrMessage = "[" & Now.Day.ToString().PadLeft(2, "0") & "-" & Now.Month.ToString().PadLeft(2, "0") & "-" & Now.Year.ToString().PadLeft(4, "0") & " " & Now.Hour.ToString().PadLeft(2, "0") & ":" & Now.Minute.ToString().PadLeft(2, "0") & ":" & Now.Second.ToString().PadLeft(2, "0") & "]" & StrDup(3, " ") & StrMessage

            If IsError = True Then
                LogPath = strErrorFolderPath & "\"
                LogFileName = LogPath & "Error_" & Format(Date.Now, "ddMMyyyy") & ".log"
            Else
                LogPath = strAuditFolderPath & "\"
                LogFileName = LogPath & "Log_" & Format(Date.Now, "ddMMyyyy") & ".log"
            End If

            If Not Directory.Exists(LogPath) Then
                Directory.CreateDirectory(LogPath)
            End If

            Dim fsObj As FileStream
            Dim SwOpenFile As StreamWriter

            If File.Exists(LogFileName) Then
                fsObj = New FileStream(LogFileName, FileMode.Append, FileAccess.Write, FileShare.Read)
            Else
                fsObj = New FileStream(LogFileName, FileMode.Create, FileAccess.Write, FileShare.Read)
            End If
            SwOpenFile = New StreamWriter(fsObj)
            SwOpenFile.WriteLine(StrMessage)

            fsObj.Flush()
            SwOpenFile.Dispose()
            fsObj = Nothing
            SwOpenFile = Nothing


        Catch ex As Exception

        End Try

    End Sub

    Public Sub WriteErrorToTxtFile(ByVal ErrorNumber As String, ByVal ErrorDesc As String, ByVal ModuleName As String, ByVal ProcName As String)

        Dim strfilename As String
        Dim strErrorString As String

        Try
            ''Change by Jaiwant dtd 31-05-2011
            ''strErrorString = "[" & Format(DateTime.Now, "dd MM yyyy") & "] [" & ErrorNumber & " " & ErrorDesc & "] [ " & ModuleName & "]"
            strErrorString = "[" & Format(DateTime.Now, "dd-MM-yyyy hh:mm:ss") & "] [" & ErrorNumber & " " & ErrorDesc & "] [ " & ModuleName & "] [ " & ProcName & "]"
            '--

            If Len(strErrorFolderPath) = 0 Then
                strErrorFolderPath = strErrorFolderPath
            End If

            If Right$(strErrorFolderPath, 1) <> "\" Then
                strErrorFolderPath = strErrorFolderPath & "\"
            End If

            strfilename = strErrorFolderPath & ModuleName & ".log"

            Dim fsObj As FileStream
            Dim SwOpenFile As StreamWriter

            If File.Exists(strfilename) Then
                fsObj = New FileStream(strfilename, FileMode.Append, FileAccess.Write, FileShare.Read)
            Else
                fsObj = New FileStream(strfilename, FileMode.Create, FileAccess.Write, FileShare.Read)
            End If

            SwOpenFile = New StreamWriter(fsObj)
            SwOpenFile.WriteLine(strErrorString)
            SwOpenFile.Dispose()
            fsObj = Nothing

        Catch er As Exception

        End Try
    End Sub
    ' IDisposable
    Protected Sub Dispose() Implements System.IDisposable.Dispose
        Me.Finalize()
        GC.SuppressFinalize(Me)
    End Sub


    'Public Sub New()
    '    InitializeComponent()
    'End Sub
End Class