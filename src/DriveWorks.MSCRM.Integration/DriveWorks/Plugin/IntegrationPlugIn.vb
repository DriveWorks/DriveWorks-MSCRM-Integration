Imports System.Windows.Forms
Imports DriveWorks.Applications
Imports DriveWorks.Applications.Extensibility

<ApplicationPlugin( _
    "DriveWorks MSCRM Integration", _
    "DriveWorks MSCRM Integration",
    "An integration plugin for DriveWorks and Microsoft Dynamics CRM Online")>
Public NotInheritable Class IntegrationPlugIn
    Implements IApplicationPlugin
    Implements IHasConfiguration

    Friend Const SOURCE_INVARIANT_NAME As String = "urn://driveworks/plugins/core/MSCRM/integration"

    Private mEventService As IApplicationEventService
    Private mSettings As PluginSettings
    Private mBaseIntegrationCore As IntegrationCore

    Public Sub Initialize(application As IApplication) Implements IApplicationPlugin.Initialize

        ' Get hold of the connections class through the shared property
        mBaseIntegrationCore = New IntegrationCore()

        If mBaseIntegrationCore.DriveWorksLoggingService Is Nothing Then
            mBaseIntegrationCore.DriveWorksLoggingService = application.ServiceManager.GetService(Of IApplicationEventService)()
        End If

        ' Create a wrapper around the application's settings manager
        ' which we'll use to read/write our own settings
        mSettings = New PluginSettings(application.SettingsManager)

        ' Get the logging service
        mEventService = application.ServiceManager.GetService(Of IApplicationEventService)()

        ' Load settings
        Me.LoadSettings()
    End Sub

    Public Sub ShowConfiguration(owner As IWin32Window) Implements IHasConfiguration.ShowConfiguration

        ' Create a new settings form and make sure it is disposed when we're done
        Using configForm = New PlugInSettingsForm()

            ' Apply current settings
            configForm.UserName.Text = Me.Settings.UserName
            configForm.Password.Text = Me.Settings.Password
            configForm.ConnectionString.Text = Me.Settings.ConnectionString

            ' Show it, and if they cancel, exit straight away
            If Not configForm.ShowDialog(owner) = Windows.Forms.DialogResult.OK Then
                Return
            End If

            ' If we got here, time to save the settings
            Me.Settings.UserName = configForm.UserName.Text
            Me.Settings.Password = configForm.Password.Text
            Me.Settings.ConnectionString = configForm.ConnectionString.Text

            ' Load settings into our handlers
            Me.LoadSettings()
        End Using
    End Sub

#Region " Properties "

    ''' <summary>
    ''' Provides access to the plugin's settings.
    ''' </summary>
    Friend ReadOnly Property Settings() As PluginSettings
        Get
            Return mSettings
        End Get
    End Property

#End Region

#Region " Helpers "

    Private Sub LoadSettings()

        ' Load the settings onto the base integration core
        mBaseIntegrationCore.LoadSettings(Me.Settings)

        ' and onto the connections module
        Connections.ConnectionInstance.LoadSettings(Me.Settings)

        If mEventService IsNot Nothing Then
            mEventService.AddEvent(ApplicationEventType.Information, SOURCE_INVARIANT_NAME, "DriveWorks MSCRM Integration", String.Format("Setting loaded ({0}:{1})", "User name", mSettings.UserName), Nothing, Nothing, Nothing)
            mEventService.AddEvent(ApplicationEventType.Information, SOURCE_INVARIANT_NAME, "DriveWorks MSCRM Integration", String.Format("Setting loaded ({0}:{1})", "Connection String", mSettings.ConnectionString), Nothing, Nothing, Nothing)
        End If
    End Sub

#End Region
End Class