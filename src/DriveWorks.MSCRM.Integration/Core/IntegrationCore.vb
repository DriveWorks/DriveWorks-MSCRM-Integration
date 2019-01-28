Option Strict On

Imports DriveWorks.Applications
Imports Microsoft.Crm.Sdk.Messages
Imports Microsoft.Xrm.Sdk
Imports Microsoft.Xrm.Sdk.Query
Imports Microsoft.Xrm.Tooling.Connector

Public Class IntegrationCore

    Private mLoggingService As IApplicationEventService
    Private mCrmServiceClient As CrmServiceClient
    Private mSettings As PluginSettings

    Public Property DriveWorksLoggingService As IApplicationEventService
        Get
            Return mLoggingService
        End Get
        Set(value As IApplicationEventService)
            mLoggingService = value
        End Set
    End Property

    Public ReadOnly Property DriveWorksPluginSettings As PluginSettings
        Get
            Return mSettings
        End Get
    End Property

    Public ReadOnly Property IsClientConnected() As Boolean
        Get
            If mCrmServiceClient Is Nothing Then
                Return False
            Else
                Return mCrmServiceClient.IsReady
            End If
        End Get

    End Property

    Public ReadOnly Property ClientStatus As String
        Get
            If mCrmServiceClient Is Nothing OrElse Not mCrmServiceClient.IsReady Then
                Return "Not Connected"
            Else
                Return "Currently Connected"
            End If
        End Get
    End Property

    ''' <summary>
    ''' Attempts to connect the client if not already connected.
    ''' </summary>
    ''' <returns>If client is successfully connected.</returns>
    Public Function Connect() As Boolean

        Return Me.Connect(mSettings.UserName, mSettings.Password, mSettings.ConnectionString)

    End Function

    ''' <summary>
    ''' Disposes of the client connection.
    ''' </summary>
    Public Sub Disconnect()

        If mCrmServiceClient IsNot Nothing Then
            mCrmServiceClient.Dispose()
            mCrmServiceClient = Nothing
        End If

    End Sub

    ''' <summary>
    ''' Instantiates a new client connection.
    ''' </summary>
    ''' <param name="username"></param>
    ''' <param name="password"></param>
    ''' <param name="userSpecifiedConnectionString"></param>
    ''' <param name="forceCreateNewClient"></param>
    ''' <returns>Whether the client has successfully connected.</returns>
    Public Function Connect(username As String, password As String, userSpecifiedConnectionString As String, ByRef Optional forceCreateNewClient As Boolean = False) As Boolean

        ' Create a new client.
        If forceCreateNewClient Or mCrmServiceClient Is Nothing Then
            Dim fullConnectionString = String.Format("AuthType=Office365;Username={0};Password={1};{2}", username, password, userSpecifiedConnectionString)
            mCrmServiceClient = New CrmServiceClient(fullConnectionString)
        End If

        Return mCrmServiceClient.IsReady

    End Function

    ''' <summary>
    ''' Executes WhoAmIRequest with the client.
    ''' </summary>
    ''' <returns>System user name.</returns>
    Public Function ExecuteWhoAmI() As String

        If Not IsClientConnected Then
            Return "#MSCRM! Failure to connect to MSCRM."
        End If

        ' Obtain information about the logged on user from the web service.
        Dim userid As Guid = (CType(mCrmServiceClient.Execute(New WhoAmIRequest()), WhoAmIResponse)).UserId
        Dim systemUser As SystemUser = CType(mCrmServiceClient.Retrieve("systemuser", userid, New ColumnSet(New String() {"firstname", "lastname"})), SystemUser)

        Return systemUser.FirstName & " " & systemUser.LastName & "."

    End Function

    ''' <summary>
    ''' Executes a retrieval request to the client for a single accounts information.
    ''' </summary>
    ''' <param name="accountID"></param>
    ''' <param name="returnRelatedEntitiesByName"></param>
    ''' <param name="fieldNames"></param>
    ''' <returns>Collection of account information.</returns>
    Public Function GetAccount(accountID As String, returnRelatedEntitiesByName As Boolean, fieldNames As String) As Object(,)

        Dim accountGuid As New Guid(accountID)
        Dim fields = QueryController.ProcessFields(fieldNames, String.Empty)

        Return Me.Retrieve("account", New Guid(accountID), fields, returnRelatedEntitiesByName)

    End Function

    ''' <summary>
    ''' Executes a retrieval request to the client for information on multiple accounts.
    ''' </summary>
    ''' <param name="filterfield"></param>
    ''' <param name="filterValue"></param>
    ''' <param name="searchField"></param>
    ''' <param name="searchValue"></param>
    ''' <param name="fieldNames"></param>
    ''' <param name="returnRelatedEntitiesByName"></param>
    ''' <param name="activeOnly"></param>
    ''' <returns>Collection of information </returns>
    Public Function GetAccounts(filterfield As String, filterValue As Object, searchField As String, searchValue As String, fieldNames As String, returnRelatedEntitiesByName As Boolean, activeOnly As Boolean) As Object(,)

        Dim fields = QueryController.ProcessFields(fieldNames, "name|accountid")
        Dim thisFilter As New FilterExpression

        If searchField <> String.Empty Then
            Dim searchCondition As New ConditionExpression With {
                .AttributeName = searchField,
                .Operator = ConditionOperator.BeginsWith
            }
            searchCondition.Values.Add(searchValue)

            thisFilter.Conditions.Add(searchCondition)
        End If

        If filterfield <> String.Empty Then
            Dim filterCondition As New ConditionExpression With {
                .AttributeName = filterfield,
                .Operator = ConditionOperator.Equal
            }
            filterCondition.Values.Add(filterValue)

            thisFilter.Conditions.Add(filterCondition)
        End If

        Return Me.RetrieveMultiple("account", fields, thisFilter, Nothing, returnRelatedEntitiesByName, activeOnly)

    End Function

    ''' <summary>
    ''' Executes a retrieval request to the client for a single contacts information.
    ''' </summary>
    ''' <param name="contactID"></param>
    ''' <param name="returnRelatedEntitiesByName"></param>
    ''' <param name="fieldNames"></param>
    ''' <returns>Collection of contact fields.</returns>
    Public Function GetContact(contactID As String, returnRelatedEntitiesByName As Boolean, fieldNames As String) As Object(,)

        Dim fields = QueryController.ProcessFields(fieldNames, String.Empty)
        Return Me.Retrieve("contact", New Guid(contactID), fields, returnRelatedEntitiesByName)

    End Function

    ''' <summary>
    ''' Executes a retrieval request to the client for information on multiple contacts.
    ''' </summary>
    ''' <param name="searchfield"></param>
    ''' <param name="searchString"></param>
    ''' <param name="returnRelatedEntitiesByName"></param>
    ''' <param name="fieldNames"></param>
    ''' <param name="activeOnly"></param>
    ''' <returns>Collection of fields on multiple contacts.</returns>
    Public Function GetContacts(searchfield As String, searchString As String, returnRelatedEntitiesByName As Boolean, fieldNames As String, activeOnly As Boolean) As Object(,)

        Dim fields = QueryController.ProcessFields(fieldNames, String.Empty)
        Dim thisFilter As New FilterExpression

        Dim condition As New ConditionExpression With {
            .AttributeName = searchfield,
            .Operator = ConditionOperator.Equal
        }
        condition.Values.Add(searchString)

        thisFilter.Conditions.Add(condition)

        Return Me.RetrieveMultiple("contact", fields, thisFilter, Nothing, returnRelatedEntitiesByName, activeOnly)

    End Function

    ''' <summary>
    ''' Executes a retrieval request to the client for an entity by querying linked entity information.
    ''' </summary>
    ''' <param name="linkFromEntityName"></param>
    ''' <param name="linkFromAttributeName"></param>
    ''' <param name="searchfield"></param>
    ''' <param name="searchString"></param>
    ''' <param name="linkFromEntityFieldNames"></param>
    ''' <param name="linkToEntityName"></param>
    ''' <param name="linkToAttributeName"></param>
    ''' <param name="linktoFilterField"></param>
    ''' <param name="linktoFilterValue"></param>
    ''' <param name="linkToFieldNames"></param>
    ''' <param name="returnRelatedEntitiesByName"></param>
    ''' <param name="activeOnly"></param>
    ''' <returns></returns>
    Public Function SearchLinkedEntitiesFilter(linkFromEntityName As String, linkFromAttributeName As String, searchfield As String, searchString As String, linkFromEntityFieldNames As String, linkToEntityName As String, linkToAttributeName As String, linktoFilterField As String, linktoFilterValue As Object, linkToFieldNames As String, returnRelatedEntitiesByName As Boolean, activeOnly As Boolean) As Object(,)

        Dim linkFromFields = QueryController.ProcessFields(linkFromEntityFieldNames, String.Empty)
        Dim linkToFields = QueryController.ProcessFields(linkToFieldNames, String.Empty)
        Dim linkFromFilter As New FilterExpression

        Dim linkToEntity As New LinkEntity With {
            .LinkFromEntityName = linkFromEntityName,
            .LinkFromAttributeName = linkFromAttributeName,
            .LinkToEntityName = linkToEntityName,
            .LinkToAttributeName = linkToAttributeName,
            .JoinOperator = JoinOperator.Inner,
            .Columns = New ColumnSet(linkToFields),
            .EntityAlias = "EntityAlias"
        }

        If Not String.IsNullOrWhiteSpace(searchfield) Then
            Dim linkFromCondition As New ConditionExpression

            linkFromCondition.AttributeName = searchfield
            linkFromCondition.Operator = ConditionOperator.BeginsWith
            linkFromCondition.Values.Add(searchString)

            linkFromFilter.Conditions.Add(linkFromCondition)
        End If

        If Not String.IsNullOrWhiteSpace(linktoFilterField) Then
            Dim linkToCondition As New ConditionExpression

            linkToCondition.AttributeName = linktoFilterField
            linkToCondition.Operator = ConditionOperator.Equal
            linkToCondition.Values.Add(linktoFilterValue)

            Dim linkToFilter As New FilterExpression

            linkToFilter.Conditions.Add(linkToCondition)

            linkToEntity.LinkCriteria.Filters.Add(linkToFilter)
        End If

        Return Me.RetrieveMultiple(linkFromEntityName, linkFromFields, linkFromFilter, linkToEntity, returnRelatedEntitiesByName, activeOnly)

    End Function

    ''' <summary>
    ''' Executes a retrieval request to the client for contact entities using a linked account search.
    ''' </summary>
    ''' <param name="searchfield"></param>
    ''' <param name="searchString"></param>
    ''' <param name="accountSearchField"></param>
    ''' <param name="accountSearchValue"></param>
    ''' <param name="contactFieldNames"></param>
    ''' <param name="accountFieldNames"></param>
    ''' <param name="returnRelatedEntitiesByName"></param>
    ''' <param name="activeOnly"></param>
    ''' <returns></returns>
    Public Function SearchContactsWithAccountSearch(searchfield As String, searchString As String, accountSearchField As String, accountSearchValue As Object, contactFieldNames As String, accountFieldNames As String, returnRelatedEntitiesByName As Boolean, activeOnly As Boolean) As Object(,)

        Dim fields = QueryController.ProcessFields(contactFieldNames, String.Empty)
        Dim accountfields = QueryController.ProcessFields(accountFieldNames, String.Empty)
        Dim contactFilter As New FilterExpression

        Dim linkEntityAccount As New LinkEntity With {
            .LinkFromEntityName = "contact",
            .LinkFromAttributeName = "parentcustomerid",
            .LinkToEntityName = "account",
            .LinkToAttributeName = "accountid",
            .JoinOperator = JoinOperator.Inner,
            .Columns = New ColumnSet(accountfields),
            .EntityAlias = "aliasAccount"
        }

        If Not String.IsNullOrWhiteSpace(searchfield) Then
            Dim contactCondition As New ConditionExpression With {
                .AttributeName = searchfield,
                .Operator = ConditionOperator.BeginsWith
            }
            contactCondition.Values.Add(searchString)

            contactFilter.Conditions.Add(contactCondition)
        End If

        If Not String.IsNullOrWhiteSpace(accountSearchField) Then
            Dim accountCondition As New ConditionExpression With {
                .AttributeName = accountSearchField,
                .Operator = ConditionOperator.BeginsWith
            }
            accountCondition.Values.Add(accountSearchValue)

            Dim accountFilter As New FilterExpression
            accountFilter.Conditions.Add(accountCondition)

            linkEntityAccount.LinkCriteria.Filters.Add(accountFilter)
        End If

        Return Me.RetrieveMultiple("contact", fields, contactFilter, linkEntityAccount, returnRelatedEntitiesByName, activeOnly)

    End Function

    ''' <summary>
    ''' Executes a retrieval request to the client for contact entities.
    ''' </summary>
    ''' <param name="filterText"></param>
    ''' <param name="fieldNames"></param>
    ''' <param name="returnRelatedEntitiesByName"></param>
    ''' <param name="activeOnly"></param>
    ''' <returns></returns>
    Public Function GetContacts(filterText As String, fieldNames As String, returnRelatedEntitiesByName As Boolean, activeOnly As Boolean) As Object(,)

        Dim fields = QueryController.ProcessFields(fieldNames, "fullname|contactid")
        Dim thisFilter As New FilterExpression

        Dim condition As New ConditionExpression With {
            .AttributeName = "fullname",
            .Operator = ConditionOperator.BeginsWith
        }
        condition.Values.Add(filterText)

        thisFilter.Conditions.Add(condition)

        Return Me.RetrieveMultiple("contact", fields, thisFilter, Nothing, returnRelatedEntitiesByName, activeOnly)

    End Function

    ''' <summary>
    ''' Executes a retrieval request to the client for information on a single entity. 
    ''' </summary>
    ''' <param name="entityName"></param>
    ''' <param name="searchfield"></param>
    ''' <param name="searchString"></param>
    ''' <param name="returnRelatedEntitiesByName"></param>
    ''' <param name="fieldNames"></param>
    ''' <param name="activeOnly"></param>
    ''' <returns></returns>
    Public Function GetSpecificEntities(entityName As String, searchfield As String, searchString As String, returnRelatedEntitiesByName As Boolean, fieldNames As String, Optional activeOnly As Boolean = True) As Object(,)

        Dim fields = QueryController.ProcessFields(fieldNames, String.Empty)

        Dim condition As New ConditionExpression With {
            .AttributeName = searchfield,
            .Operator = ConditionOperator.Like
        }
        condition.Values.Add(searchString)

        Dim thisFilter As New FilterExpression
        thisFilter.Conditions.Add(condition)

        Return Me.RetrieveMultiple(entityName, fields, thisFilter, Nothing, returnRelatedEntitiesByName, activeOnly)

    End Function

    ''' <summary>
    ''' Allows a PluginSettings instance to be loaded into the controller. Required for client connection details.
    ''' </summary>
    ''' <param name="settings"></param>
    Public Sub LoadSettings(ByVal settings As PluginSettings)

        mSettings = settings

    End Sub

    ''' <summary>
    ''' Executes a create request to the client for creating a new contact.
    ''' </summary>
    ''' <param name="accountID"></param>
    ''' <param name="valuePairs"></param>
    ''' <returns></returns>
    Public Function AddContact(accountID As String, valuePairs As Dictionary(Of String, String)) As String

        Dim newContact As New Contact()

        ' Convert the new value pairs to data on the contract.
        For Each valuePair In valuePairs
            newContact(valuePair.Key.ToLower()) = valuePair.Value
        Next

        If Not (String.IsNullOrEmpty(accountID)) Then
            Dim accountname As String = Me.GetAccountNameFromID(accountID)

            If Not String.IsNullOrEmpty(accountname) Then
                ' We have both an account id and a account name.
                newContact.ParentCustomerId = New EntityReference("account", New Guid(accountID))
            Else
                ' The account name is not blank, and it doesn't exist, we will throw an exception.
                Return "#MSCRM! Contact not created because Account does not exist in CRM: " & accountID
            End If
        End If

        ' Create a new contact and return the new guid.
        If IsClientConnected Then
            Return mCrmServiceClient.Create(newContact).ToString()
        Else
            Return "#MSCRM! Failure to connect to MSCRM."
        End If

    End Function

    ''' <summary>
    ''' Executes a create request to the client for creating a new activity on an account.
    ''' </summary>
    ''' <param name="activityType"></param>
    ''' <param name="accountID"></param>
    ''' <param name="valuePairs"></param>
    ''' <param name="markCompleted"></param>
    ''' <returns></returns>
    Public Function AddActivityToAccount(activityType As String, accountID As String, valuePairs As Dictionary(Of String, String), markCompleted As Boolean) As String

        If Not IsClientConnected Then
            Return "#MSCRM! Failure to connect to MSCRM."
        End If

        Dim newActivity As New Entity(activityType)

        ' Convert the new value pairs to data on the contract.
        For Each valuePair In valuePairs
            newActivity(valuePair.Key.ToLower()) = valuePair.Value
        Next

        If Not (String.IsNullOrEmpty(accountID)) Then

            Dim accountname As String = Me.GetAccountNameFromID(accountID)

            If Not String.IsNullOrEmpty(accountname) Then
                ' We have both an account id and a account name.
                newActivity.Attributes("regardingobjectid") = New EntityReference("account", New Guid(accountID))
            Else
                ' The account name is not blank, and it doesn't exist, we will throw an exception.
                Return "#MSCRM! Activity not created because Account does not exist in CRM: " & accountID
            End If

        End If

        ' Create a new activity and return the new GUID.
        Dim activityGuid = mCrmServiceClient.Create(newActivity)

        If markCompleted Then
            Dim statusRequest As New SetStateRequest With {
                .EntityMoniker = New EntityReference(activityType, activityGuid),
                .State = New OptionSetValue(1),
                .Status = New OptionSetValue(2)
            }

            mCrmServiceClient.Execute(statusRequest)
        End If

        Return activityGuid.ToString()

    End Function

    Private Function GetAccountNameFromID(accountID As String) As String

        Dim fields = New String() {"name"}
        Dim result = Me.Retrieve("account", New Guid(accountID), fields, True)

        If result.GetUpperBound(0) > 0 Then

            Return result(1, 0).ToString()

        Else
            Return String.Empty
        End If

    End Function

    Private Function Retrieve(entityname As String, entityID As Guid, fieldNames As String(), returnRelatedEntitiesByName As Boolean) As Object(,)

        If Not IsClientConnected Then
            Return Nothing
        End If

        Dim queryColumns As New ColumnSet

        If fieldNames.Count = 0 Then
            queryColumns = New ColumnSet(True)
        Else
            queryColumns.Columns.AddRange(fieldNames)
            queryColumns.AllColumns = False
        End If

        Dim entityResult As Entity = mCrmServiceClient.Retrieve(entityname, entityID, queryColumns)
        Dim columnCount As Integer = entityResult.Attributes.Count
        Dim returnResults(1, columnCount - 1) As Object

        ' Iterate over all attributes.
        Dim index As Integer = 0
        For Each att In entityResult.Attributes

            returnResults(0, index) = att.Key
            returnResults(1, index) = QueryController.GetValueString(att.Value, returnRelatedEntitiesByName)

            index = index + 1

        Next

        Return returnResults

    End Function

    Private Function RetrieveMultiple(entityname As String, fieldNames As String(), filter As FilterExpression, link As LinkEntity, returnRelatedEntitiesByName As Boolean, activeOnly As Boolean) As Object(,)

        If Not IsClientConnected Then
            Return Nothing
        End If

        Dim queryExpression = QueryController.GetQueryExpression(entityname, fieldNames, filter, link, activeOnly)
        Dim resultsDictionary As New Dictionary(Of Integer, Dictionary(Of Integer, String))

        ' Initiate the header dictionary based on the incoming field names.
        Dim headerDictionary As New Dictionary(Of Integer, String)
        Dim attributeColumn As Integer = 0

        For Each headertext In fieldNames
            headerDictionary.Add(attributeColumn, headertext)
            attributeColumn = attributeColumn + 1
        Next

        Dim recordcount As Integer = 5000
        Dim pagenumber As Integer = 1

        While recordcount = 5000

            ' Call CRM web service.
            Dim entityResults = mCrmServiceClient.RetrieveMultiple(queryExpression)

            resultsDictionary = QueryController.LoadDictionary(entityResults, headerDictionary, returnRelatedEntitiesByName)

            queryExpression.PageInfo.PagingCookie = entityResults.PagingCookie
            queryExpression.PageInfo.PageNumber = pagenumber + 1

            recordcount = entityResults.Entities.Count
            pagenumber = pagenumber + 1

        End While

        ' We now have the results dictionary, lets add it to the results string array.
        Dim returnResults(resultsDictionary.Count, headerDictionary.Count - 1) As String

        ' Add on the headers.
        For Each header In headerDictionary
            returnResults(0, header.Key) = header.Value
        Next

        ' Now add on the data.
        For Each value In resultsDictionary

            For Each attribute In value.Value

                returnResults(value.Key, attribute.Key) = attribute.Value

            Next
        Next

        Return returnResults

    End Function

End Class