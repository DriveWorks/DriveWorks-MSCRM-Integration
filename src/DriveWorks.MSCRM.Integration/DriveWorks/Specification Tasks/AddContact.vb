Imports DriveWorks.EventFlow
Imports DriveWorks.Specification

<Task("Microsoft CRM: Add a new contact to an existing account", "embedded://DriveWorks.MSCRM.Integration.Microsoft_Dynamics_Logo_16px.png", "Microsoft CRM")>
Public Class MSCRMAddContact
    Inherits Specification.Task

    ' This task is bringing in all of the data as separate strings.
    ' It would be just as appropriate to pass the data in as a single array or named pairs, or by using variable prefix's or suffix's

    ' Register properties so DriveWorks can see them and build rules for them
    Private ReadOnly mFirstName As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("First Name", New FlowPropertyInfo("The First Name of the Contact.", "Names"))
    Private ReadOnly mMiddleName As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Middle Name", New FlowPropertyInfo("The Middle Name of the Contact (if known).", "Names"))
    Private ReadOnly mLastName As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Last Name", New FlowPropertyInfo("The Last Name of the Contact.", "Names"))
    Private ReadOnly mDescription As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Description", New FlowPropertyInfo("Notes that you may have about the contact.", "Additional Information"))
    Private ReadOnly mEMailAddress1 As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Email Address1", New FlowPropertyInfo("Primary Email Address (if more than one email address, use the custom fields).", "Names"))
    Private ReadOnly mJobTitle As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Job Title", New FlowPropertyInfo("The Job Title for the Contact (if known).", "Additional Information"))
    Private ReadOnly mMobilePhone As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Mobile Phone", New FlowPropertyInfo("Mobile phone number.", "Additional Information"))
    Private ReadOnly mSalutation As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Salutation", New FlowPropertyInfo("Salutation, for example Mr, Mrs, Ms.", "Additional Information"))
    Private ReadOnly mTelephone1 As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Business Phone", New FlowPropertyInfo("This is Primary phone number for the Contact.", "Additional Information"))
    Private ReadOnly mAccountID As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Account Id", New FlowPropertyInfo("This is the Account ID for the contact (leave blank if no account).", "Account Information"))
    Private ReadOnly mCustomFields As FlowProperty(Of String) = Me.Properties.RegisterStringProperty("Custom Data", New FlowPropertyInfo("Data for custom fields in the format CustField1=Custvalue1|CustField2=CustValue2. (Not Required).", "Additional Information"))

    Private ReadOnly mOutputNodeContactId As NodeOutput = Me.Outputs.Register("Contact ID", "The newly created contact's ID.", GetType(String))

    Protected Overrides Sub Execute(ByVal ctx As SpecificationContext)

        Dim result As String
        Dim connected As Boolean = False
        Dim valuePairs As New Dictionary(Of String, String)
        Dim accountID As String = mAccountID.Value
        Dim connectionManager = Connections.ConnectionInstance.GetConnection(ctx.Project)

        ' Collect the data into a dictionary.
        ' It doesn't have to be this way, any storage method is appropriate and should be created with regard to the data sending method (Some systems require a dictionary, some txt or xml, and others objects.)
        valuePairs("Description") = mDescription.Value
        valuePairs("EMailAddress1") = mEMailAddress1.Value
        valuePairs("FirstName") = mFirstName.Value
        valuePairs("JobTitle") = mJobTitle.Value
        valuePairs("LastName") = mLastName.Value
        valuePairs("MiddleName") = mMiddleName.Value
        valuePairs("MobilePhone") = mMobilePhone.Value
        valuePairs("Salutation") = mSalutation.Value
        valuePairs("Telephone1") = mTelephone1.Value

        Try

            ' In this example, we are also have the ability to pass in custom field data
            For Each CustomField As String In mCustomFields.Value.Split("|"c)

                If CustomField.Contains("=") Then

                    Dim fieldName = CustomField.Split("="c)(0)

                    Dim fieldValue As String = String.Empty

                    fieldValue = CustomField.Substring(fieldName.Length + 1)

                    valuePairs(fieldName) = fieldValue

                End If

            Next

            ' Send the data to the third party system.
            result = connectionManager.RunWithRetry(Function(connection)
                                                        Return connection.AddContact(accountID, valuePairs)
                                                    End Function)

            mOutputNodeContactId.Fulfill(result)
            Me.SetState(NodeExecutionState.Successful, result)

        Catch ex As Exception

            ' Report that we had a major problem.
            Dim errorMessage = "Error: " & ex.ToString()
            Me.SetState(EventFlow.NodeExecutionState.Failed, errorMessage)

        End Try

    End Sub

End Class