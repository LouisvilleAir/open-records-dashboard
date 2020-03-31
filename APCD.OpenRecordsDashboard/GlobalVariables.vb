Imports APCD
Imports System.Text
Imports System.IO

Public Class GlobalVariables

    'Load in MainForm.LoadLeaveCalendarEditorList()
    Private Shared m_LeaveCalendarEditors As New ArrayList
    Public Shared Property LeaveCalendarEditors() As ArrayList
        Get
            Return m_LeaveCalendarEditors
        End Get
        Set(ByVal value As ArrayList)
            m_LeaveCalendarEditors = value
        End Set
    End Property

    Private Shared m_TestGroupList As New ArrayList
    Public Shared Property TestGroupList() As ArrayList
        Get
            Return m_TestGroupList
        End Get
        Set(ByVal value As ArrayList)
            m_TestGroupList = value
        End Set
    End Property


    ''' <summary>
    ''' Specifies the appointment type for a calendar item.
    ''' </summary>
    ''' <remarks>The values here must match the AppointmentTypeID values in the AppointmentType.</remarks>
    Public Enum AppointmentTypeEnum
        AllDay = 1
        ArrivingLate = 2
        LeavingEarly = 3
        MidDayAppointment = 4
    End Enum

    Public Enum LeaveTypeEnum
        Off = 1
        Sick = 2
        Training = 3
        OutInField = 4
        WorkingFromHome = 5
        Inspection = 6
        Volunteering = 7
        Holiday = 8
        LeavingEarly = 9
        ArrivingLate = 10
        NA = -1
    End Enum


    Public Enum MainConnectionStatusEnum
        Available
        Unavailable
    End Enum

    Private Shared enumMainConnectionStatus As MainConnectionStatusEnum
    Public Shared Property MainConnectionStatus() As MainConnectionStatusEnum
        Get
            Return enumMainConnectionStatus
        End Get
        Set(ByVal value As MainConnectionStatusEnum)
            enumMainConnectionStatus = value
        End Set
    End Property

    'Friend Structure GlobalForms

    'End Structure

    Public Enum DatabaseModeEnum
        Development
        Test
        Production
    End Enum

    Public Enum Role
        User = 1
        PowerUser = 2
        Administrator = 3
        Programmer = 4
        Guest = 5
        Approver = 6
    End Enum

    Friend Structure DatabaseName
        Const Applications As String = "Applications"
        Const People As String = "People"
    End Structure

    Private Shared m_objApplication As Common.Applications.Business.Application
    Public Shared Property Application() As Common.Applications.Business.Application
        Get
            Return m_objApplication
        End Get
        Set(ByVal value As Common.Applications.Business.Application)
            m_objApplication = value
        End Set
    End Property

    Private Shared m_objEmployee As PeopleLibrary.Business.Employee
    Public Shared Property Employee() As PeopleLibrary.Business.Employee
        Get
            Return m_objEmployee
        End Get
        Set(ByVal value As PeopleLibrary.Business.Employee)
            m_objEmployee = value
        End Set
    End Property

    Private Shared m_objEmployeeApplication As Common.Applications.Business.EmployeeApplication
    Public Shared Property EmployeeApplication() As Common.Applications.Business.EmployeeApplication
        Get
            Return m_objEmployeeApplication
        End Get
        Set(ByVal value As Common.Applications.Business.EmployeeApplication)
            m_objEmployeeApplication = value
        End Set
    End Property

    Private Shared m_strDatabaseMode As DatabaseModeEnum
    Public Shared Property DatabaseMode() As DatabaseModeEnum
        Get
            Return m_strDatabaseMode
        End Get
        Set(ByVal value As DatabaseModeEnum)
            m_strDatabaseMode = value
        End Set
    End Property

    'todo 20120424: rename to something else? (DatabaseStatusColor, DatabaseModeColor)
    Private Shared m_colorConnectionStatusColor As Color
    Public Shared Property ConnectionStatusColor() As Color
        Get
            Return m_colorConnectionStatusColor
        End Get
        Set(ByVal value As Color)
            m_colorConnectionStatusColor = value
        End Set
    End Property

    Private Shared m_enumUserRole As GlobalVariables.Role
    Public Shared Property UserRole() As GlobalVariables.Role
        Get
            Return m_enumUserRole
        End Get
        Set(ByVal value As GlobalVariables.Role)
            m_enumUserRole = value
        End Set
    End Property

#Region "----- log and log locations -----"

    Public Shared ReadOnly Property APCDApplicationsRootDataDirectory() As String
        Get
            Return My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\APCD Applications"
        End Get
    End Property

    Public Shared ReadOnly Property ApplicationDataDirectory As String
        Get
            Return APCDApplicationsRootDataDirectory & "\Common"
        End Get
    End Property

    Public Shared ReadOnly Property ApplicationConfigurationDirectory As String
        Get
            Return ApplicationDataDirectory & "\Config"
        End Get
    End Property

    Public Shared ReadOnly Property TestLogDirectory() As String
        Get
            Return "G:\DATA\Test\Application Logs\Open Records Dashboard\" & Format(Now, "yyyyMMdd") & ".log"
        End Get
    End Property

    Public Shared ReadOnly Property ProductionLogDirectory() As String
        Get
            Return "G:\DATA\Production\Application Logs\Open Records Dashboard\" & Format(Now, "yyyyMMdd") & ".log"
        End Get
    End Property

    Private Shared m_strApplicationLog As String
    Public Shared Property ApplicationLog() As String
        Get
            Return m_strApplicationLog
        End Get
        Set(ByVal value As String)
            m_strApplicationLog = value
        End Set
    End Property

#End Region '----- log locations -----


End Class
