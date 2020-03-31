Imports APCD.Common 'Applications, People
Imports System.IO

Public Class splashScreenForm

    Private Sub splashScreenForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If (Me.ApplicationIsRunning) Then
            MessageBox.Show("Application is already running", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End
        Else
#If DEBUG Then
            Call Me.StartUp_DebugMode()
#Else
            Call Me.StartUp_ReleaseMode()
#End If
            Me.Timer1.Start()
        End If

    End Sub

    Private Function ApplicationIsRunning() As Boolean

        Dim blnApplicationIsRunning As Boolean
        Dim strProcessName As String = Process.GetCurrentProcess().ProcessName
        Dim processes As Process() = Process.GetProcessesByName(strProcessName)

        If (processes.Length > 1) Then
            blnApplicationIsRunning = True
        Else
            blnApplicationIsRunning = False
        End If

        Return blnApplicationIsRunning

    End Function

    Private Sub StartUp_DebugMode()

        'is the user a programmer?
        '   yes - open choose screen
        '   no  - test mode

        Dim exitTheApplication As Boolean = False
        Me.Show()
        Application.DoEvents()

        Dim mainConnection As String = Me.GetMain
        'Dim devConfigFile As New FileInfo(GlobalVariables.ApplicationConfigurationDirectory & "\developer.cfg")

        'If (devConfigFile.Exists) Then 'user is a programmer
        '    If (mainConnection = String.Empty) Then
        '        'no main connection, work disconnected
        '        Call GlobalMethods.RunInDisconnectedMode(devConfigFile.FullName)
        '    Else
        '        'have main connection, let the developer choose
        '        Dim frm As New chooseApplicationModeForm(devConfigFile)
        '        frm.ShowDialog()
        '    End If

        'Else 'user is not a programmer
        If (mainConnection = String.Empty) Then
            MessageBox.Show("Main connection is unavailable. Please contact APCD IT for assistance.", My.Application.Info.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            exitTheApplication = True
        Else
            GlobalMethods.SetMainConnnection(mainConnection)
            GlobalVariables.DatabaseMode = GlobalVariables.DatabaseModeEnum.Test
            GlobalVariables.ConnectionStatusColor = Color.Yellow
            Call Me.Engage()
        End If

        'End If

        If (exitTheApplication = True) Then
            End
        End If

    End Sub

    Private Sub StartUp_ReleaseMode()

        Dim exitTheApplication As Boolean = False
        Me.Show()
        Application.DoEvents()

        Dim mainConnection As String = Me.GetMain
        If (mainConnection = String.Empty) Then
            MessageBox.Show("Main connection is unavailable. Please contact APCD IT for assistance.", My.Application.Info.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            exitTheApplication = True
        Else
            GlobalMethods.SetMainConnnection(mainConnection)
            GlobalVariables.DatabaseMode = GlobalVariables.DatabaseModeEnum.Production
            Call Me.Engage()
        End If

        If (exitTheApplication = True) Then
            End
        End If

    End Sub

    Private Function GetMain() As String

        Dim main As String = String.Empty
        Dim obj As APCD.DataProviderLibrary.MainService

        Try
            obj = CType(Activator.GetObject(Type.GetType("APCD.DataProviderLibrary.MainService, APCD.DataProviderLibrary"), _
                 "tcp://apcbar02tryand1:6000/MainServiceDataProvider"),  _
                  APCD.DataProviderLibrary.MainService)
            main = obj.GetMain
        Catch ex As Exception
            main = String.Empty
        Finally
            obj = Nothing
        End Try

        Return main

    End Function

#Region "----- methods to handle a connected session -----"

    Private Sub Engage()

        Dim userName As String = GlobalMethods.GetFriendlyUserName
        Dim exitTheApplication As Boolean = False

        Try
            Call GlobalMethods.InitializeApplicationObject()
            Call GlobalMethods.InitializeApplicationDatabases()
            Call GlobalMethods.ValidateTheUser(userName)
            Call Me.ValidateTheConnection()
            Call GlobalMethods.AssignUserRole()
            Call Me.SetLogDirectory()
            Applications.Utility.EmployeeApplicationUtility.Logon(GlobalVariables.EmployeeApplication.EmployeeApplicationID) 'flag the user as logged on
        Catch ex As Exception
            MessageBox.Show("An unexpected error occurred during application startup: " & ex.Message, "Application Aborted", MessageBoxButtons.OK, MessageBoxIcon.Error)
            exitTheApplication = True
        End Try

        If (exitTheApplication = True) Then
            End
        End If

    End Sub

    Private Sub ValidateTheConnection()

        Dim cnTesting123Testing As OleDb.OleDbConnection = Nothing

        Try
            cnTesting123Testing = New OleDb.OleDbConnection(PeopleLibrary.Globals.GlobalVariable.ConnectionString)
            cnTesting123Testing.Open()
        Catch ex As Exception
            Throw New ApplicationException("The application is currently offline.")
        Finally
            If (cnTesting123Testing.State = ConnectionState.Open) Then cnTesting123Testing.Close()
        End Try

    End Sub

    Private Sub SetLogDirectory()
        Select Case GlobalVariables.DatabaseMode
            Case GlobalVariables.DatabaseModeEnum.Test
                GlobalVariables.ApplicationLog = GlobalVariables.TestLogDirectory

            Case GlobalVariables.DatabaseModeEnum.Production
                GlobalVariables.ApplicationLog = GlobalVariables.ProductionLogDirectory
        End Select
    End Sub

#End Region '----- methods to handle a connected session -----

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.Close()
    End Sub

End Class