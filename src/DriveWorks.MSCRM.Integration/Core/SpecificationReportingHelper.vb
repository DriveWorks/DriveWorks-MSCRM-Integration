Imports DriveWorks.Applications
Imports DriveWorks.Reporting

Public Class SpecificationReportingHelper
    Implements IApplicationEventService

    Public Event EventLogged(sender As Object, e As ApplicationEventEventArgs) Implements IApplicationEventService.EventLogged
    Public Event EventsCleared(sender As Object, e As EventArgs) Implements IApplicationEventService.EventsCleared

    Private ReadOnly mSpecReport As IReportWriter

    Public Sub New(specReport As IReportWriter)

        mSpecReport = specReport

    End Sub

    Public Sub AddEvent(type As ApplicationEventType, sourceInvariantName As String, sourceDisplayName As String, description As String, targetInvariantName As String, targetDisplayName As String, url As String) Implements IApplicationEventService.AddEvent

        Dim entryType As ReportEntryType

        ' The 2 reporting types have different enums which need normalizing.
        Select Case type
            Case ApplicationEventType.Error
                entryType = ReportEntryType.Error
            Case ApplicationEventType.Information
                entryType = ReportEntryType.Information
            Case ApplicationEventType.Warning
                entryType = ReportEntryType.Warning
        End Select

        ' Always write as minimal.
        mSpecReport.WriteEntry(ReportingLevel.Minimal, entryType, sourceDisplayName, targetDisplayName, description, url)

    End Sub

    Public Sub ClearEvents() Implements IApplicationEventService.ClearEvents

        ' We aren't expecting this to be called for our report.
        Throw New Exception()

    End Sub

    Public Function GetEvents() As IApplicationEvent() Implements IApplicationEventService.GetEvents

        ' We aren't expecting this to be called for our report.
        Throw New Exception()

    End Function
End Class