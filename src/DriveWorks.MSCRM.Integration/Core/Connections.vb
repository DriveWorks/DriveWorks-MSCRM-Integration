Friend Class Connections

    Private mSettings As PluginSettings
    Public Shared ReadOnly ConnectionInstance As New Connections()

    ' Make sure nothing else can create this.
    Private Sub New()

    End Sub

    ' We do need to store the settings here for passing through to the integrationcore.
    Friend Sub LoadSettings(ByVal settings As PluginSettings)

        ' Store the settings.
        mSettings = settings

    End Sub

    ' This will create a new connection manager if one doesn't exist already.
    Public Function GetConnection(project As Project) As ConnectionManager

        Dim connectionMan = project.GetSharedObject(Of ConnectionManager)()

        connectionMan.InitializeSettings(mSettings)
        connectionMan.InitializeProject(project)

        Return connectionMan

    End Function

End Class
