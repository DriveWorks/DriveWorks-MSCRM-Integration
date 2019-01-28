Friend Class ConnectionManager

    Private mIntegrationCore As IntegrationCore
    Private mSettings As PluginSettings
    Private mProject As Project

    ' The settings are required regardless of the entry point, so the connectionmanager needs to store them for when a connection is required.
    Public Sub InitializeSettings(settings As PluginSettings)

        If mSettings Is Nothing Then
            mSettings = settings
        End If

    End Sub

    ' If we are working with projects and specifications, we may also need one of those, mainly so that we can get hold of the specification report.
    Public Sub InitializeProject(project As Project)

        If mProject Is Nothing Then
            mProject = project
        End If

    End Sub

#Region " Retry code "

    ' This allows up to three attempts at each method or function.
    Friend Function RunWithRetry(Of TReturn)(ByVal action As Func(Of IntegrationCore, TReturn)) As TReturn

        Dim i As Integer = 3

        While True
            i -= 1

            Try
                ' Get hold of the existing connection if there is one, if there isn't, it will be created.
                Me.EnsureConnection()
                Return action(mIntegrationCore)
            Catch ex As Exception
                If i = 0 Then
                    Throw
                Else
                    ' It failed, throw away the connection.
                    Me.DestroyConnection()
                End If
            End Try
        End While

        ' Can't get here
        Throw New InvalidOperationException()
    End Function

    ' Make sure that we have a valid connection.
    Private Sub EnsureConnection()

        If mIntegrationCore Is Nothing Then
            mIntegrationCore = New IntegrationCore
            mIntegrationCore.LoadSettings(mSettings)

            ' Check that we are in a specification, and if we are pass the specification report into the integrationcore object.
            If mProject.SpecificationContext IsNot Nothing Then
                mIntegrationCore.DriveWorksLoggingService = New SpecificationReportingHelper(mProject.SpecificationContext.Report)
            End If
        End If

        mIntegrationCore.Connect()

    End Sub

    ' If something goes wrong, we need to destroy the connection and re-connect.
    Private Sub DestroyConnection()

        ' If we don't have an integrationcore, then we cannot disconnect.
        If mIntegrationCore Is Nothing Then
            Return
        End If

        mIntegrationCore.Disconnect()

    End Sub

#End Region

End Class