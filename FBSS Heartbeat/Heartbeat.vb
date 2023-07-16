Imports System
Imports System.Threading
Imports System.Diagnostics
Imports System.Data.SqlClient
Imports System.Environment
Imports System.Math
Imports System.Xml
Imports System.IO
Public Class Heartbeat
    Private Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    Protected Overrides Sub OnStart(ByVal args() As String)
        Dim autoEvent As New AutoResetEvent(False)
        Dim objMainframe As New Mainframe
        objMainframe.ServiceName = Me.ServiceName
        Dim thdCheckStatus As New System.Threading.Thread(AddressOf objMainframe.CheckStatus)
        thdCheckStatus.Start()
    End Sub
    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
    End Sub
End Class

Public Class Mainframe
    Private Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
    Dim strServiceName As String
    Public Property ServiceName()
        Get
            ServiceName = strServiceName
        End Get
        Set(ByVal value)
            strServiceName = value
        End Set
    End Property
    Sub CheckStatus()
        Dim blnCancel As Boolean
        Dim strComputerName As String
        Dim strServerName As String
        Dim dblRandomStart As Double
        Dim dblMaximum_Minutes_Before_First_Logging As Double
        Dim dblInterval_Minutes_Before_Next_Logging As Double
        Dim strStatus As String

        dblMaximum_Minutes_Before_First_Logging = GetValue("FBSS_Heartbeat/Config", "Maximum_Minutes_Before_First_Logging")

        strComputerName = ""
        strServerName = ""
        strStatus = ""
        Try
            strComputerName = MachineName
        Catch ex As Exception

        End Try
        If Len(strComputerName) >= 8 Then
            strServerName = "XQ" & MachineName.Substring(1, 5)
        Else
            strServerName = "Invalid LU (" & strComputerName & ")"
        End If

        Randomize()
        dblRandomStart = Round((Rnd() * dblMaximum_Minutes_Before_First_Logging) + 1, 0)
        dblRandomStart = dblRandomStart * 60000
        Sleep(dblRandomStart)

        Do Until blnCancel = True
            Try
                strStatus = getStatus()
                LogHeartbeatToDatabase(strServerName, strComputerName, strStatus, Now)
                LogHeartbeatToFile(strServerName, strComputerName, strStatus, Now)
                System.Diagnostics.EventLog.WriteEntry(strServiceName, strStatus)
            Catch ex As Exception
                
            End Try

            dblInterval_Minutes_Before_Next_Logging = GetValue("FBSS_Heartbeat/Config", "Interval_Minutes_Before_Next_Logging")
            Sleep(dblInterval_Minutes_Before_Next_Logging * 60000)
        Loop
    End Sub
    Private Function getStatus() As String
        Try
            Dim strReturn As String
            Dim strIMSPingPath As String

            strIMSPingPath = GetValue("FBSS_Heartbeat/Config", "IMSPing_Path")

            strReturn = ReadCmdOutput(strIMSPingPath & "\Imsping.exe", "", strIMSPingPath, True)
            If strReturn.IndexOf("TIME") <> -1 Then
                Return "Available"
            Else
                Return "Unavailable"
            End If
        Catch ex As Exception
            Return "Unavailable"
        End Try
    End Function
    Friend Function ReadCmdOutput(ByVal applicationName As String, Optional ByVal applicationArgs As String = "", Optional ByVal workingDirectory As String = "", Optional ByVal showWindow As Boolean = False) As String
        Try
            Dim processObj As New Process

            processObj.StartInfo.UseShellExecute = False
            processObj.StartInfo.RedirectStandardOutput = True
            processObj.StartInfo.FileName = applicationName
            processObj.StartInfo.Arguments = applicationArgs
            processObj.StartInfo.WorkingDirectory = workingDirectory

            If showWindow = True Then
                processObj.StartInfo.CreateNoWindow = False
            Else
                processObj.StartInfo.CreateNoWindow = True
            End If

            processObj.Start()
            processObj.WaitForExit()

            Return processObj.StandardOutput.ReadToEnd
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Private Function LogHeartbeatToDatabase(ByVal strServerName As String, ByVal strClientName As String, ByVal strStatus As String, ByVal dteDateAndTime As Date) As Boolean
        Try
            Dim strSQL_Server As String
            Dim strSQL_User_Name As String
            Dim strSQL_Password As String
            Dim strSQL_Database As String

            strSQL_Server = GetValue("FBSS_Heartbeat/Config", "SQL_Server")
            strSQL_User_Name = GetValue("FBSS_Heartbeat/Config", "SQL_User_Name")
            strSQL_Password = GetValue("FBSS_Heartbeat/Config", "SQL_Password")
            strSQL_Database = GetValue("FBSS_Heartbeat/Config", "SQL_Database")

            Dim objSqlConnection As SqlConnection
            Dim objSqlCommand As New SqlCommand

            objSqlConnection = New SqlConnection("Server=" & strSQL_Server & ";uid=" & strSQL_User_Name & ";pwd=" & strSQL_Password & ";database=" & strSQL_Database)
            objSqlConnection.Open()

            objSqlCommand.Connection = objSqlConnection
            objSqlCommand.CommandType = CommandType.StoredProcedure
            objSqlCommand.CommandText = "LogHeartbeat"

            Dim parServerName As New SqlParameter
            Dim parClientName As New SqlParameter
            Dim parStatus As New SqlParameter
            Dim parDateAndTime As New SqlParameter

            parServerName.ParameterName = "@ServerName"
            parServerName.DbType = DbType.String
            parServerName.Value = strServerName

            parClientName.ParameterName = "@ClientName"
            parClientName.DbType = DbType.String
            parClientName.Value = strClientName

            parStatus.ParameterName = "@Status"
            parStatus.DbType = DbType.String
            parStatus.Value = strStatus

            parDateAndTime.ParameterName = "@DateAndTime"
            parDateAndTime.DbType = DbType.DateTime
            parDateAndTime.Value = dteDateAndTime

            objSqlCommand.Parameters.Add(parServerName)
            objSqlCommand.Parameters.Add(parClientName)
            objSqlCommand.Parameters.Add(parStatus)
            objSqlCommand.Parameters.Add(parDateAndTime)

            objSqlCommand.ExecuteNonQuery()

            objSqlConnection.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Function GetValue(ByVal strPath As String, ByVal strValueName As String)
        Try
            Dim objXmlDocument As New XmlDocument
            Dim objXmlNode As XmlNode
            Dim strServicePath As String

            strServicePath = GetServicePath()
            objXmlDocument.Load(GetServicePath() & "Config.xml")
            objXmlNode = objXmlDocument.SelectSingleNode(strPath)
            GetValue = objXmlNode.SelectSingleNode(strValueName).InnerText.Trim
            objXmlDocument = Nothing
        Catch ex As Exception
            End
        End Try     
    End Function
    Private Sub SetValue(ByVal strPath As String, ByVal strValueName As String, ByVal strValue As String)
        Try
            Dim objXmlDocument As New XmlDocument
            Dim objXmlNode As XmlNode
            Dim strServicePath As String

            strServicePath = GetServicePath()
            objXmlDocument.Load(strServicePath & "Config.xml")
            objXmlNode = objXmlDocument.SelectSingleNode(strPath)
            objXmlNode.SelectSingleNode(strValueName).InnerText = strValue
            objXmlDocument.Save(strServicePath & "Config.xml")
            objXmlDocument = Nothing
        Catch ex As Exception
            End
        End Try
    End Sub
    Private Function LogHeartbeatToFile(ByVal strServerName As String, ByVal strClientName As String, ByVal strStatus As String, ByVal dteDateAndTime As Date) As Boolean
        Try
            Dim objLogFile As StreamWriter
            objLogFile = New StreamWriter(GetLogFileName, True)
            objLogFile.WriteLine(strServerName & vbTab & strClientName & vbTab & strStatus & vbTab & dteDateAndTime)
            objLogFile.Close()
        Catch ex As Exception

        End Try
    End Function
    Private Function GetLogFileName() As String
        Dim intMaximumNumberOfLogFiles As Integer
        Dim intMinimumSizeBeforeNewLogFile As Integer
        Dim intCnt As Integer
        Dim arrLogFiles As String()
        Dim strReturn As String
        Dim arrLastWriteTime As String()
        Dim intCurrentLogFileIndex As Integer
        Dim objFileInfo As FileInfo

        intMaximumNumberOfLogFiles = GetValue("FBSS_Heartbeat/Config", "Maximum_Number_Of_Log_Files")
        intMinimumSizeBeforeNewLogFile = GetValue("FBSS_Heartbeat/Config", "Minimum_Size_Before_New_Log_File")
        intCurrentLogFileIndex = GetValue("FBSS_Heartbeat/Config", "Current_Log_File_Index")

        ReDim arrLogFiles(intMaximumNumberOfLogFiles - 1)
        ReDim arrLastWriteTime(intMaximumNumberOfLogFiles - 1)
        For intCnt = 0 To intMaximumNumberOfLogFiles - 1
            arrLogFiles(intCnt) = GetEnvironmentVariable("SystemDrive") & "\FBSS_Heartbeat_" & (intCnt + 1) & ".log"
        Next
        strReturn = ""

        If File.Exists(arrLogFiles(intCurrentLogFileIndex)) = True Then
            objFileInfo = New FileInfo(arrLogFiles(intCurrentLogFileIndex))
            If objFileInfo.Length < intMinimumSizeBeforeNewLogFile Then
                Return arrLogFiles(intCurrentLogFileIndex)
            Else
                If intCurrentLogFileIndex >= intMaximumNumberOfLogFiles - 1 Then
                    intCurrentLogFileIndex = 0
                Else
                    intCurrentLogFileIndex = intCurrentLogFileIndex + 1
                End If
            End If
        Else
            Return arrLogFiles(intCurrentLogFileIndex)
        End If
        If File.Exists(arrLogFiles(intCurrentLogFileIndex)) = True Then
            File.Delete(arrLogFiles(intCurrentLogFileIndex))
        End If
        SetValue("FBSS_Heartbeat/Config", "Current_Log_File_Index", intCurrentLogFileIndex)
        Return arrLogFiles(intCurrentLogFileIndex)
    End Function
    Private Function GetServicePath()
        Return System.Environment.CommandLine.Substring(1, System.Environment.CommandLine.LastIndexOf("\"))
    End Function
End Class

