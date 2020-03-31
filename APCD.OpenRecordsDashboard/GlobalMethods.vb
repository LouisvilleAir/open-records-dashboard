Imports APCD.Common
'Imports APCD.AdminAssistant.GlobalVariables 'connnectionMode
Imports System.IO

Public Class GlobalMethods

    Friend Shared Function GetFriendlyUserName() As String

        Dim strUserName As String

        If (My.User.Name.Contains("\")) Then
            strUserName = My.User.Name.Substring(My.User.Name.IndexOf("\") + 1)
        Else
            strUserName = My.User.Name
        End If

        Return strUserName

    End Function

#Region "----- methods to handle a disconnected session -----"

    Friend Shared Sub RunInDisconnectedMode(ByVal configFilePath As String)

        GlobalVariables.DatabaseMode = GlobalVariables.DatabaseModeEnum.Development
        Dim userName As String = GlobalMethods.GetFriendlyUserName
        Call GlobalMethods.SetOfflineLogFilesAndConnectionStrings(configFilePath)
        Call GlobalMethods.InitializeApplicationObject()
        Call GlobalMethods.ValidateTheUser(userName)
        Call GlobalMethods.AssignUserRole()
        GlobalVariables.ConnectionStatusColor = Color.Green

    End Sub

    Friend Shared Sub SetOfflineLogFilesAndConnectionStrings(ByVal configFilePath As String)

        'read the file into a hashtable
        Dim sr As New StreamReader(configFilePath)
        Dim keyValuePairTable As New Hashtable
        Do While (sr.Peek > -1)
            Dim tempLine As String = sr.ReadLine()
            If (Not tempLine Is Nothing) AndAlso (tempLine.Length > 0) Then
                Dim indexOfFirstEqualSign As Int32 = tempLine.IndexOf(APCD.Tools.Constants.Character.EqualSign)
                Dim key As String = tempLine.Substring(0, indexOfFirstEqualSign)
                Dim value As String = tempLine.Substring(indexOfFirstEqualSign + 1, tempLine.Length - (indexOfFirstEqualSign + 1))
                keyValuePairTable.Add(key, value)
            End If
        Loop
        sr.Close()

        'all apps need the log file and Applications and People for authentication
        GlobalVariables.ApplicationLog = CStr(keyValuePairTable("logDirectory")) & "\" & Format(Now, "yyyyMMdd") & ".log"
        Applications.Globals.GlobalVariables.ConnectionString = CStr(keyValuePairTable(GlobalVariables.DatabaseName.Applications))
        APCD.PeopleLibrary.Globals.GlobalVariable.ConnectionString = CStr(keyValuePairTable(GlobalVariables.DatabaseName.People))

    End Sub

#End Region '----- methods to handle a disconnected session -----


#Region "----- methods to handle a connected session -----"

    Friend Shared Sub SetMainConnnection(ByVal mainConnection As String)

        Try
            Applications.Globals.GlobalVariables.ConnectionString = mainConnection
        Catch ex As Exception
            Throw New ApplicationException("Unable to set main connection.")
        End Try

    End Sub

    Friend Shared Sub InitializeApplicationObject()

        Try
            GlobalVariables.Application = Applications.Utility.ApplicationUtility.GetByLookupName(My.Application.Info.ProductName)
            If (GlobalVariables.Application Is Nothing) Then
                Throw New ApplicationException("Unable to obtain application information.")
            End If
        Catch ex As Exception
            Throw New ApplicationException("Unable to obtain application information.")
        End Try

    End Sub

    Friend Shared Sub InitializeApplicationDatabases()
        Try
            'set all of the DBs for a given application
            Dim dtb As DataTable = Applications.Utility.ApplicationDatabasseUtility.GetByApplicationID(GlobalVariables.Application.ApplicationID)
            For Each row As DataRow In dtb.Rows
                Dim databaseID As Int32 = CInt(row(Applications.Constants.ApplicationDatabasseConstants.FieldName.DatabasseID))
                Dim objDatabase As Applications.Business.Databasse = Applications.Utility.DatabasseUtility.GetByPrimaryKey(databaseID)

                Select Case GlobalVariables.DatabaseMode

                    'the only connection that should be set to TEST here is PeopleLibrary
                    Case GlobalVariables.DatabaseModeEnum.Test
                        Select Case objDatabase.DatabasseName
                            Case GlobalVariables.DatabaseName.Applications
                                Applications.Globals.GlobalVariables.ConnectionString = objDatabase.ProductionConnectionString
                            Case GlobalVariables.DatabaseName.People
                                PeopleLibrary.Globals.GlobalVariable.ConnectionString = objDatabase.TestConnectionString
                        End Select

                    Case GlobalVariables.DatabaseModeEnum.Production
                        Select Case objDatabase.DatabasseName
                            Case GlobalVariables.DatabaseName.Applications
                                Applications.Globals.GlobalVariables.ConnectionString = objDatabase.ProductionConnectionString
                            Case GlobalVariables.DatabaseName.People
                                PeopleLibrary.Globals.GlobalVariable.ConnectionString = objDatabase.ProductionConnectionString
                        End Select
                End Select
            Next
        Catch ex As Exception
            Throw New ApplicationException("Unable initialize application databases.")
        End Try
    End Sub

    Friend Shared Sub ValidateTheUser(ByVal userName As String)

        Try
            GlobalVariables.Employee = PeopleLibrary.Utility.EmployeeUtility.GetByUserName(userName)
        Catch ex As Exception
            Throw New ApplicationException("Unable to obtain user information.")
        End Try

        Try
            GlobalVariables.EmployeeApplication = Applications.Utility.EmployeeApplicationUtility.GetByEmployeeID_ApplicationID(GlobalVariables.Employee.EmployeeID, GlobalVariables.Application.ApplicationID)
        Catch ex As Exception
            Throw New ApplicationException("Unable to obtain employee application information.")
        End Try

        If (GlobalVariables.EmployeeApplication Is Nothing) Then
            Throw New ApplicationException("You are not a user of this system.")
        ElseIf (GlobalVariables.EmployeeApplication.IsActive = False) Then
            Throw New ApplicationException("Your access to this application has been disabled. Please contact APCD IT for assistance.")
        End If

    End Sub

#End Region '----- methods to handle a connected session -----

    Friend Shared Sub HandleError(ByVal ex As Exception)

        Dim message As New System.Text.StringBuilder
        Dim messagePrefix As New System.Text.StringBuilder

        With messagePrefix
            .Append(Format(Date.Now, "HH:mm:ss"))
            .Append(APCD.Tools.Constants.Character.Space)
            'If GlobalVariables.Employee.UserName Is Nothing Then
            '    .Append("Username unavailable.")
            'Else
            '    .Append(GlobalVariables.Employee.UserName)
            'End If
            .Append(GetFriendlyUserName())
            .Append(vbTab)
        End With

        message.Append(messagePrefix.ToString())
        message.Append("Message: ")
        message.AppendLine(ex.Message)
        message.Append(messagePrefix.ToString())
        message.AppendLine("------------------------------------------------------------------------------------------------------------------------------------------")

        Try
            APCD.Tools.WindowsForms.WriteToApplicationLogFile(message.ToString, GlobalVariables.ApplicationLog)
        Catch e As Exception
            Dim outerMessage As New System.Text.StringBuilder
            With outerMessage
                .AppendLine("An error occurred while writing to the application log file:")
                .AppendLine(e.Message)
                .AppendLine("Original error: ")
                .Append(message.ToString)
            End With
            APCD.Tools.WindowsForms.WriteToWindowsEventLog(My.Computer.Name, My.Application.Info.ProductName, outerMessage.ToString, EventLogEntryType.Error, 1)
        End Try

    End Sub

    Friend Shared Sub HandleError(ByVal message As String)

        Dim messagePrefix As New System.Text.StringBuilder
        With messagePrefix
            .Append(Format(Date.Now, "HH:mm:ss"))
            .Append(Space(1))
            .Append(GlobalVariables.Employee.UserName)
            .Append(Space(3))
        End With

        Dim errorMessage As New System.Text.StringBuilder
        errorMessage.Append(messagePrefix.ToString())
        errorMessage.AppendLine(message)

        Try
            APCD.Tools.WindowsForms.WriteToApplicationLogFile(errorMessage.ToString, GlobalVariables.ApplicationLog)
        Catch e As Exception
            Dim outerMessage As New System.Text.StringBuilder
            With outerMessage
                .AppendLine("An error occurred while writing the the application log file:")
                .AppendLine(e.Message)
                .AppendLine("Original error: ")
                .Append(message.ToString)
            End With
            APCD.Tools.WindowsForms.WriteToWindowsEventLog(My.Computer.Name, My.Application.Info.ProductName, outerMessage.ToString, EventLogEntryType.Error, 1)
        End Try

    End Sub

    Friend Shared Sub AssignUserRole()

        Select Case GlobalVariables.EmployeeApplication.RoleID
            Case 1
                GlobalVariables.UserRole = GlobalVariables.Role.User
            Case 2
                GlobalVariables.UserRole = GlobalVariables.Role.PowerUser
            Case 3
                GlobalVariables.UserRole = GlobalVariables.Role.Administrator
            Case 4
                GlobalVariables.UserRole = GlobalVariables.Role.Programmer
            Case 5
                GlobalVariables.UserRole = GlobalVariables.Role.Guest
            Case 6
                GlobalVariables.UserRole = GlobalVariables.Role.Approver
        End Select

    End Sub

    Friend Shared Function FormatThePhoneNumber(ByVal phoneText As String, ByVal separator As String) As String

        Dim formattedNumber As String = String.Empty

        If (IsNumeric(phoneText)) Then
            If (phoneText.Length = 7) Then
                formattedNumber = phoneText.Insert(3, separator)

            ElseIf (phoneText.Length = 10) Then
                formattedNumber = APCD.Tools.Constants.Character.LeftParenthesis _
                                & phoneText.Substring(0, 3) _
                                & APCD.Tools.Constants.Character.RightParenthesis _
                                & APCD.Tools.Constants.Character.Space _
                                & phoneText.Substring(3, 3) _
                                & separator _
                                & phoneText.Substring(6, 4)
            End If



        End If

        Return formattedNumber

    End Function

End Class
