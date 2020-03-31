
Public Class MainForm

#Region "Constants"
    Private Structure HoverMessage
        Const ACTSSearch As String = "The search is case sensitive; enter part of the address using the correct capitalization. Keep pressing Enter to see more results. Press the Esc key to exit."
        Const ACTSPrint As String = "Enter the year, then the permit number within that year. The permit will print on the default printer and the program will close."
        Const ACTS7 As String = "Log on (need to have a logon to ACTS), then select year. Requires FoxPro runtime components."
        Const ArchivedPlantList As String = "Select the Archive tab. Sites are grouped by street name.You can search through the spreadsheet using Excel features."
        Const AsbestosArchive As String = "There is a text file for each year containing the full text of asbestos permits. You can search each file using Ctrl+F in Notepad."
        Const DocumentManager As String = "Select Tools, Search Documents.Select File Name as the search method. Check Search by Plant, then select the plant ID obtained from Hansen or Plants && Permits. Uncheck Include Draft (and leave Include Confidential unchecked), then click Search."
        Const Hansen8 As String = "You may want to start with the Contact and Property browser under Resources to see everything related to an address. Or you can search as follows:  Complaints:  Select Customer Service, Lookup Service Requests. For Responsiblity, select LAPC. Enter the address or other criteria, then click Search. Asbestos permits:  Select Building Permits, Lookup Applications. Select A/P type ASBESTOS, enter the address or just street name and click Search. Emissions-based permits and gasoline dispensing permits:  Select Project, Lookup Applications. Enter the address or just street name and click Search. Ignore results with application type SUBPMT. Results with application types TITLE V, FEDOOP, or MINOR can be ignored if there is an APCDOP application for the same facility. Starting in early 2016, enforcement cases and citations are under Code Enforcement case types ENFORCE and CITATE respectively."
        Const IMS As String = "To print a complaint: Press 1, 1, 2, 2 and enter the CAP report number or search for one. To print an APPI: 1, 4, 5. To print an incident:  2, 1, 4. Press 9 at the main menu to exit."
        Const IMSAccess As String = "Log on (contact APCD computer support if you need a login). For general reports, press 6, then select a report and enter appropriate criteria.  Or you can navigate for other functions; the menus roughly match IMS itself. Press 9 at the main menu to exit. This item only works if Access 2007 is installed on the computer."
        Const IMSSearch As String = "Double-click the appropriate query, then enter the search criteria. You can close the query results window and run another query. Close the program window when done. This item only works if Access 2007 is installed on the computer."
        Const OldViolations As String = "Enter the address or part of the address and press Enter. The search is not case sensitive."
        Const PlantsAndPermits As String = "Press 1, wait, press Enter, press 5, then press 1 to search by plant name or 2 to search by part of address. The search is not case sensitive."
        Const PlantReports As String = "Press 1 for general plant reports, then select the appropriate report and proceed to select criteria. When you have the correct output, return to the main menu and press 9 to select the network printer or a text file as output, then press 8 to re-run the report.  Press Enter at the main menu to exit."
        Const GasStationSystem As String = "Press 2, then 5, then press 1 to search by plant name or 2 to search by part of address. The search is not case sensitive."
        Const GasReports As String = "Press 2 for station reports, then select the appropriate report and proceed to select criteria. When you have the correct output, return to the main menu and press 9 to select the network printer or a text file as output, then press 8 to re-run the report.  Press Enter at the main menu to exit."
        Const OpenRecordsDatabase As String = "Start with the Open Records Request form."
    End Structure

#End Region

    Private Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Me.Visible = False
        'Dim frm As New splashScreenForm
        'frm.ShowDialog()

        'Me.Visible = True

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub OpenRecordsDatabaseButton_Click(sender As Object, e As EventArgs) Handles OpenRecordsDatabaseButton.Click
        Try
            System.Diagnostics.Process.Start("S:\Open Records\Open Records.accdb")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub HansenButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HansenButton.Click
        Try
            System.Diagnostics.Process.Start("http://mshanapp200.msd.louky.local/hansen8/")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub IMSButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IMSButton.Click
        Try
            System.Diagnostics.Process.Start("G:\DATA\IMS\IMS.BAT")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub IMSSearchButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IMSSearchButton.Click
        Try
            Dim program As String = "C:\Program Files\Microsoft Office\Office12\MSACCESS.EXE"
            Dim arguments As String = "G:\DATA\IMS2\IMSSearch.mdb"
            System.Diagnostics.Process.Start(program, arguments)
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub IMSAccessButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IMSAccessButton.Click
        Try
            Dim program As String = "C:\Program Files\Microsoft Office\Office12\MSACCESS.EXE"
            Dim arguments As String = "G:\DATA\IMS2\APCDIMS.mdb /WRKGRP G:\DATA\SYSTEM.MDW"
            System.Diagnostics.Process.Start(program, arguments)
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub PlantsAndPermitsButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlantsAndPermitsButton.Click
        Try
            System.Diagnostics.Process.Start("G:\FPS\OPR\PlantsAndPermits.bat")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub PlantReportsButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlantReportsButton.Click
        Try
            System.Diagnostics.Process.Start("G:\FPS\OPR\PlantReports.bat")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub GasStationSystemButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GasStationSystemButton.Click
        Try
            System.Diagnostics.Process.Start("G:\FPS\OPR\GasStationSystem.bat")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub GasStationReportsButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GasStationReportsButton.Click
        Try
            System.Diagnostics.Process.Start("G:\FPS\OPR\GasStationReports.bat")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub ACTSSearchButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ACTSSearchButton.Click
        Try
            System.Diagnostics.Process.Start("G:\DATA\ACTS6\PROG\search00.exe")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub ArchivedPlantListButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ArchivedPlantListButton.Click
        Try
            System.Diagnostics.Process.Start("S:\Engineering\Archives\ARC.XLS")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub DocumentManagerButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DocumentManagerButton.Click
        Try
            System.Diagnostics.Process.Start("http://apcbar02tryand1/InstallDocumentManager/APCD.DocumentManager.application")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub ACTSButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ACTSButton.Click
        Try
            System.Diagnostics.Process.Start("G:\DATA\Acts7\acts7.exe")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub OldNOVButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OldNOVButton.Click
        Try
            System.Diagnostics.Process.Start("G:\DATA\Archive\VIOL\ViolationSearch.bat")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub ACTSPrintButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ACTSPrintButton.Click
        Try
            System.Diagnostics.Process.Start("G:\DATA\ACTS6\PROG\prtdrf00.exe")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub AsbestosArchiveButton_Click(sender As Object, e As EventArgs) Handles AsbestosArchiveButton.Click
        Try
            System.Diagnostics.Process.Start("G:\DATA\Archive\ASBESTOS\")
        Catch ex As Exception
            Call GlobalMethods.HandleError(ex)
        End Try
    End Sub

    Private Sub HansenButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HansenButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.Hansen8
    End Sub

    Private Sub DocumentManagerButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DocumentManagerButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.DocumentManager
    End Sub

    Private Sub IMSSearchButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IMSSearchButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.IMSSearch
    End Sub

    Private Sub PlantsAndPermitsButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlantsAndPermitsButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.PlantsAndPermits
    End Sub

    Private Sub ACTSSearchButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ACTSSearchButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.ACTSSearch
    End Sub

    Private Sub ACTSPrintButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ACTSPrintButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.ACTSPrint
    End Sub

    Private Sub OldNOVButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OldNOVButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.OldViolations
    End Sub

    Private Sub HansenButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HansenButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub DocumentManagerButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DocumentManagerButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub IMSSearchButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IMSSearchButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub IMSAccessButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IMSAccessButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub IMSButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IMSButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub OldNOVButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OldNOVButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub PlantsAndPermitsButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlantsAndPermitsButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub PlantReportsButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlantReportsButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub GasStationSystemButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GasStationSystemButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.GasStationSystem
    End Sub

    Private Sub GasStationSystemButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GasStationSystemButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub GasStationReportsButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GasStationReportsButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub ArchivedPlantListButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ArchivedPlantListButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub ACTSSearchButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ACTSSearchButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub ACTSPrintButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ACTSPrintButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub ACTSButton_MouseLeave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ACTSButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub IMSAccessButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IMSAccessButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.IMSAccess
    End Sub

    Private Sub IMSButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IMSButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.IMS
    End Sub

    Private Sub PlantReportsButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlantReportsButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.PlantReports
    End Sub

    Private Sub GasStationReportsButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GasStationReportsButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.GasReports
    End Sub

    Private Sub ACTSButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ACTSButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.ACTS7
    End Sub

    Private Sub ArchivedPlantListButton_MouseHover(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ArchivedPlantListButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.ArchivedPlantList
    End Sub

    Private Sub AsbestosArchiveButton_MouseHover(sender As Object, e As EventArgs) Handles AsbestosArchiveButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.AsbestosArchive
    End Sub

    Private Sub AsbestosArchiveButton_MouseLeave(sender As Object, e As EventArgs) Handles AsbestosArchiveButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub

    Private Sub OpenRecordsDatabaseButton_MouseHover(sender As Object, e As EventArgs) Handles OpenRecordsDatabaseButton.MouseHover
        Me.InstructionsLabel.Text = MainForm.HoverMessage.OpenRecordsDatabase
    End Sub

    Private Sub OpenRecordsDatabaseButton_MouseLeave(sender As Object, e As EventArgs) Handles OpenRecordsDatabaseButton.MouseLeave
        Me.InstructionsLabel.Text = String.Empty
    End Sub
End Class

