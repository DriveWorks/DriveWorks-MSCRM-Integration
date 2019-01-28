Imports DriveWorks.EventFlow
Imports DriveWorks.Specification

<Task("Microsoft CRM: Add a new activity to an existing account", "embedded://DriveWorks.MSCRM.Integration.Microsoft_Dynamics_Logo_16px.png", "Microsoft CRM")>
Public Class MSCRMAddActivityToAccount
    Inherits Specification.Task

    ' This task is bringing in all of the data as separate strings.
    ' It would be just as appropriate to pass the data in as a single array or named pairs, or by using variable prefix's or suffix's

    ' Register properties so DriveWorks can see them and build rules for them
    Private ReadOnly mSubject As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Subject", New FlowPropertyInfo("The Subject of the activity.", "Additional Information"))
    Private ReadOnly mType As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Type", New FlowPropertyInfo("The type of the activity (for instance email, fax, task.", "Additional Information"))
    Private ReadOnly mDescription As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Description", New FlowPropertyInfo("The Description of the activity.", "Additional Information"))
    Private ReadOnly mAccountID As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Account Id", New FlowPropertyInfo("This is the Account ID for the contact. The task will fail if this isn't a valid account GUID.", "Account Information"))
    Private ReadOnly mCustomFields As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Custom Data", New FlowPropertyInfo("Data for custom fields in the format CustField1=Custvalue1|CustField2=CustValue2. (Not Required).", "Additional Information"))
    Private ReadOnly mMarkCompleted As FlowProperty(Of Boolean) = Me.Properties.RegisterBooleanProperty("Mark as Completed", New FlowPropertyInfo("TRUE to mark the activity as completed, FALSE to leave open.", "Additional Information"))

    Private ReadOnly mOutputNodeActivityId As NodeOutput = Me.Outputs.Register("Activity ID", "todo", GetType(String))

    Protected Overrides Sub Execute(ByVal ctx As SpecificationContext)

        Dim result As String
        Dim connected As Boolean = False
        Dim accountID As String = mAccountID.Value
        Dim activityType As String = mType.Value
        Dim valuePairs As New Dictionary(Of String, String)
        Dim connectionManager = Connections.ConnectionInstance.GetConnection(ctx.Project)

        If String.IsNullOrEmpty(activityType) Then

            Me.SetState(NodeExecutionState.InputsMissing, "Activity Type Required")
            Return

        End If

        Try

            ' Collect the data into a dictionary.
            ' It doesn't have to be this way, any storage method is appropriate and should be created with regard to the data sending method (Some systems require a dictionary, some txt or xml, and others objects.).
            valuePairs("Description") = mDescription.Value
            valuePairs("Subject") = mSubject.Value
            valuePairs("activitytypecode") = mType.Value

            ' In this example, we are also have the ability to pass in custom field data.
            For Each customField As String In mCustomFields.Value.Split("|"c)

                If customField.Contains("=") Then

                    Dim fieldName = customField.Split("="c)(0)
                    Dim fieldValue = customField.Substring(fieldName.Length + 1)
                    valuePairs(fieldName) = fieldValue

                End If

            Next

            ' Send the data to the third party system.
            result = connectionManager.RunWithRetry(Function(connection)
                                                        Return connection.AddActivityToAccount(activityType, accountID, valuePairs, mMarkCompleted.Value)
                                                    End Function)

            mOutputNodeActivityId.Fulfill(result)
            Me.SetState(NodeExecutionState.Successful, result)

        Catch ex As Exception

            ' Report that we had a major problem.
            Dim errorMessage = "Error: " & ex.ToString()
            Me.SetState(NodeExecutionState.Failed, errorMessage)

        End Try

    End Sub

End Class