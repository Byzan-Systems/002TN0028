Imports System.IO
Imports System.Text
Imports System.Drawing
Imports System
Imports System.Data
Imports Microsoft.Office.Interop

Module GenrateOutput

    Dim objLogCls As New ClsErrLog
    Dim objGetSetINI As ClsShared
    Dim objBaseClass As ClsBase
    Dim objValidationClass As ClsValidation
    Dim SumOfAmount As Double = 0
   
    Public Function GenerateOutPutFile(ByRef dtOutput As DataTable, ByVal strFileName As String) As Boolean
        Dim gstrA2Afile As String = String.Empty
        Dim strMethodCalForEpay As Boolean = False
        Try
            objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
            objValidationClass = New ClsValidation(strFileName, objBaseClass.gstrIniPath)
            FileCounter = objBaseClass.GetINISettings("General", "File Counter", My.Application.Info.DirectoryPath & "\settings.ini")
            FileCounter = FileCounter + 1

            If Len(FileCounter) < 3 Then
                FileCounter = FileCounter.PadLeft(4, "0").Trim()
                FileCounter = FileCounter.Substring(FileCounter.Length - 3, 3)
            End If

            strFileName = (objValidationClass.IsJustAlpha(Path.GetFileNameWithoutExtension(gstrInputFile), 10, "N"))
            '   strFileName = (objValidationClass.IsJustAlpha(gstrInputFile, 10, "N"))

            '  gstrOutputFile_Name = "Cheque_" & strFileName & ".xls "

            Dim strOptFileNameSplitArr() As String = Nothing
            Dim strOptFileName As String = ""

            strOptFileNameSplitArr = strFileName.Split("-")

            If strOptFileNameSplitArr.Length > 0 Then
                strOptFileName = strOptFileNameSplitArr(0).Substring(0, 3) & strOptFileNameSplitArr(1)
            End If


            Dim FileCount As Integer = dtOutput.DefaultView.ToTable(True, "File_No").Rows.Count
            For index = 1 To FileCount
                If FileCount = 1 Then
                    gstrOutputFile_Name = strOptFileName & ".xls"
                Else
                    gstrOutputFile_Name = strOptFileName & "_" & index & ".xls"
                End If

                If Generate_Output(dtOutput, gstrOutputFile_Name, index) = False Then
                    GenerateOutPutFile = False
                Else
                    GenerateOutPutFile = True
                    Call objBaseClass.SetINISettings("General", "File Counter", Val(FileCounter), My.Application.Info.DirectoryPath & "\settings.ini")
                End If
            Next


        Catch ex As Exception
            GenerateOutPutFile = False
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "GenerateOutput", "GenerateOutPutFile")
        End Try


    End Function

    Private Function Generate_Output(ByRef dt As DataTable, ByVal strOptFileName As String, ByVal FileNo As Integer) As Boolean

        Dim ExlApp As New Excel.Application
        Dim ExlWb As Excel.Workbook
        Dim ExSht As Excel.Worksheet
        Dim strLog As String = """"
        Dim ColDisplay As String() = {"MAJOR_HEAD", "BASIC_TAX", "SUR_CHARGE", "EDU_CESS", "INTEREST", "PENALTY", "OTHERS", "TOTAL_AMT"}

        Try

            strLog = "Output"
            '   gstrOutputFile_Name = strFileName & ".xls "

            If dt.Rows.Count > 0 Then
                objBaseClass.LogEntry(strLog & " File Generation Process Start")

                Dim RowNo As Integer = 0
                Dim ColNo As Integer = 1
                Dim DrRow As DataRow() = Nothing

                ExlApp.Visible = False
                ExlWb = ExlApp.Workbooks.Add
                ExSht = DirectCast(ExlWb.ActiveSheet, Excel.Worksheet)
                ExSht.Name = "Sheet1"

                Dim RecordCount As Integer = 0

                '-Header Section'''Commented by swati dtd 2022-11-25
                'For Index As Int32 = 0 To dt.Columns.Count - 6
                '    ExSht.Cells(RowNo, ColNo) = dt.Columns(Index).ToString()
                '    ColNo += 1
                'Next
                ''--
                '-Details Section
                For Each drRBI As DataRow In dt.Select("File_No=" & FileNo)
                    RowNo += 1
                    ColNo = 1
                    RecordCount += 1
                    For index = 0 To dt.Columns.Count - 6
                        If ColDisplay.Contains(dt.Columns(index).ColumnName.ToString().Trim().ToUpper()) Then
                            ExSht.Cells(RowNo, ColNo) = "'" & drRBI(index).ToString().Trim() & ""
                        Else
                            ExSht.Cells(RowNo, ColNo) = drRBI(index).ToString().Trim()
                        End If
                        ColNo += 1
                    Next
                Next
                '-

                ExSht.Columns.AutoFit()
                ExSht = Nothing
                ExlApp.DisplayAlerts = False

                ExlWb.SaveAs(strOutputFolderPath & "\" & strOptFileName, Excel.XlFileFormat.xlWorkbookNormal)

                ExlWb.Close()
                ExlApp.Quit()

                objBaseClass.LogEntry(strLog & " File [" & strOptFileName & "] Generated Successfully")

                Generate_Output = True
            Else
                objBaseClass.LogEntry(strLog & " Record Not Found")
                Generate_Output = False
            End If


        Catch ex As Exception
            Generate_Output = False
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "GenerateOutput", "Generate_Output")
            ExlApp.Quit()
        Finally

            '  ExlWb.Close()
            ExlApp.Quit()
        End Try
    End Function
    Public Function Generate_Output_Response(ByRef _dtRes As DataTable, ByRef _dtHeader As DataTable, ByVal strRespFileName As String) As Boolean
        Dim strOutPutLine As String
        Dim objStrmWriter As StreamWriter
        Dim strRevResp_OptFileName As String = ""
        Try
            If _dtRes.Rows.Count > 0 Then

                objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
                strRevResp_OptFileName = Path.GetFileNameWithoutExtension(strRespFileName) & ".txt"

                objStrmWriter = New StreamWriter(strReverseResponseFolderPath & "\" & strRevResp_OptFileName)
                objBaseClass.LogEntry("Reverse Output File generating process Started...")

                '''''Header section
                strOutPutLine = ""
                For Inti As Int32 = 0 To _dtHeader.Columns.Count - 6
                    strOutPutLine = strOutPutLine & (_dtHeader.Rows(0)(Inti).ToString()) & "~"
                Next

                strOutPutLine = Left(strOutPutLine, strOutPutLine.Length - 1)
                objStrmWriter.WriteLine(strOutPutLine, strRevResp_OptFileName)

                '''''Detail section
                For Each drRow As DataRow In _dtRes.Rows
                    strOutPutLine = ""

                    For Inti As Int32 = 0 To drRow.ItemArray.Length - 6
                        strOutPutLine = strOutPutLine & (drRow.ItemArray(Inti).ToString()) & "~"
                    Next

                    strOutPutLine = Left(strOutPutLine, strOutPutLine.Length - 1)
                    objStrmWriter.WriteLine(strOutPutLine, strRevResp_OptFileName)
                Next

                If Not objStrmWriter Is Nothing Then
                    objStrmWriter.Close()
                    objStrmWriter.Dispose()

                End If
                objBaseClass.LogEntry("Reverse Response Output File [" & strRevResp_OptFileName & "] is generated successfully", False)

                Generate_Output_Response = True
            Else
                objBaseClass.LogEntry("Reverse Response Record Not Found")
            End If


        Catch ex As Exception
            Generate_Output_Response = False
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "GenerateOutput", "Generate_Output_Response")
        Finally
            If Not objStrmWriter Is Nothing Then
                objStrmWriter.Close()
                objStrmWriter.Dispose()

            End If
        End Try
    End Function

    Public Function Check_Comma(ByVal strTemp) As String
        Try
            If InStr(strTemp, ",") > 0 Then

                ' Check_Comma = Chr(34) & strTemp & Chr(34) & ","
                Check_Comma = strTemp
            Else
                Check_Comma = strTemp & ","
            End If

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Payment", "Check_Comma")

        End Try
    End Function

    Private Function Pad_Length(ByVal strtemp As String, ByVal intLen As Integer) As String
        Try
            Pad_Length = Microsoft.VisualBasic.Left(strtemp & StrDup(intLen, " "), intLen)

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call objBaseClass.Handle_Error(ex, "frmGenericRBI", Err.Number, "Pad_Length")

        End Try
    End Function


    Function RemoveCharacter(ByVal stringToCleanUp As String)
        Dim characterToRemove As String = ""
        characterToRemove = Chr(34) + "=~^!#$%&'()*+,-@`/\:{}[]"

        Dim firstThree As Char() = characterToRemove.Take(30).ToArray()
        For index = 0 To firstThree.Length - 1
            stringToCleanUp = stringToCleanUp.ToString.Replace(firstThree(index), "")
        Next
        Return stringToCleanUp
    End Function
End Module
