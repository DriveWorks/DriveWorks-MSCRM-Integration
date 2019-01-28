Imports DriveWorks.Extensibility
Imports Titan.Rules.Execution

Public Class DriveWorksFunctions
    Inherits SharedProjectExtender

    Private Const FUNCTION_CATEGORY As String = "DriveWorks MSCRM Integration"

    <Udf(True)>
    <FunctionInfo("Returns a String value that describes the current status of the connection to the third party system.", FUNCTION_CATEGORY)>
    Public Function MSCRMConnectionStatus(<ParamInfo("Trigger", "Change this value in order to re-evaluate the function.")> trigger As Object) As String

        Try
            Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                              Return connection.ClientStatus
                                                          End Function)
        Catch ex As Exception
            Return "#MSCRM! Unable to retrieve Connection status: " & ex.Message
        End Try

    End Function

    <Udf(True)>
    <FunctionInfo("Returns a Boolean value representing if the connection to the 3rd party system is open.", FUNCTION_CATEGORY)>
    Public Function MSCRMConnected(<ParamInfo("Trigger", "Change this value in order to re-evaluate the function.")> trigger As Object) As Object

        Try
            Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                              Return connection.IsClientConnected
                                                          End Function)
        Catch ex As Exception
            Return "#MSCRM! Unable to retrieve Connected status: " & ex.Message
        End Try

    End Function

    <Udf(True)>
    <FunctionInfo("Returns the 3rd Party connection string, based on the settings for this plugin.", FUNCTION_CATEGORY)>
    Public Function MSCRMConnectionString() As String

        Try
            Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                              Return connection.DriveWorksPluginSettings.ConnectionString
                                                          End Function)
        Catch ex As Exception
            Return "#MSCRM! Unable to retrieve Connection String: " & ex.Message
        End Try

    End Function

    <Udf(True)>
    <FunctionInfo("Returns the 3rd Party username, based on the settings for this plugin.", FUNCTION_CATEGORY)>
    Public Function MSCRMUserName() As String

        Try
            Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                              Return connection.DriveWorksPluginSettings.UserName
                                                          End Function)
        Catch ex As Exception
            Return "#MSCRM! Unable to retrieve user name: " & ex.Message
        End Try

    End Function

#Region " Contacts "

    <Udf(True)>
    <FunctionInfo("Returns the first name and last name of the connected user in Dynamics CRM.", FUNCTION_CATEGORY)>
    Public Function MSCRMExecuteWhoAmI(<ParamInfo("Connect to Dynamics CRM", "False to NOT connect, any other value to connect. Can be used as a way to refresh the data, by simply changing this value.")> connect As Object) As Object

        If IsFalse(connect) Then
            Return ErrorType.ValueException
        Else

            Try
                Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                                  Return connection.ExecuteWhoAmI()
                                                              End Function)
            Catch ex As Exception
                Return "#MSCRMCONNECT! Unable to get the data for that contact: " & ex.Message
            End Try
        End If

    End Function

    <Udf(True)>
    <FunctionInfo("Returns the data for a single contact from Dynamics CRM.", FUNCTION_CATEGORY)>
    Public Function MSCRMGetContactByID(<ParamInfo("Contact ID", "The contacts GUID.")> contactID As String,
                                        <ParamInfo("Field Names", "A pipebar (|) delimited list of field names. These names need to match those in CRM exactly.  Leave blank for all fields.")> FieldNames As String,
                                        <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntityAsID As Boolean,
                                        <ParamInfo("Connect to Dynamics CRM", "False to NOT connect, any other value to connect. Can be used as a way to refresh the data, by simply changing this value.")> connect As Object
                                        ) As Object

        If IsFalse(connect) Then
            Return ErrorType.ValueException
        Else

            Try
                Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                                  Return New StandardArrayValue(connection.GetContact(contactID, returnRelatedEntityAsID, FieldNames))
                                                              End Function)
            Catch ex As Exception
                Return "#MSCRMCONNECT! Unable to get the data for that contact: " & ex.Message
            End Try
        End If

    End Function

    <Udf(True)>
    <FunctionInfo("Returns the data for a single contact from Dynamics CRM.", FUNCTION_CATEGORY)>
    Public Function MSCRMGetContactByID(<ParamInfo("Contact ID", "The contacts GUID.")> contactID As String,
                                        <ParamInfo("Field Names", "A pipebar (|) delimited list of field names. These names need to match those in CRM exactly. Leave blank for all fields.")> FieldNames As String,
                                        <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntityAsID As Boolean
                                        ) As Object

        Return Me.MSCRMGetContactByID(contactID, FieldNames, returnRelatedEntityAsID, True)

    End Function

    <Udf(True)>
    <FunctionInfo("Returns contacts from Dynamics CRM. There may be a limit on the amount of data your CRM installation can send. You may need to filter this list or request less fields.", FUNCTION_CATEGORY)>
    Public Function MSCRMGetContacts(<ParamInfo("Filter Text", "The filter string. Any contact whose name contains this text will be returned.")> contactNameFilter As String,
                                     <ParamInfo("Field Names", "A pipebar (|) delimited list of field names. These names need to match those in CRM exactly. Leave blank for all fields.")> FieldNames As String,
                                     <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntitiesByName As Boolean,
                                     <ParamInfo("Connect to Dynamics CRM", "False to NOT connect, any other value to connect. Can be used as a way to refresh the data, by simply changing this value.")> connect As Object
                                     ) As Object

        If IsFalse(connect) Then
            Return ErrorType.ValueException
        Else

            Try
                Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                                  Return New StandardArrayValue(connection.GetContacts(contactNameFilter, FieldNames, returnRelatedEntitiesByName, True))
                                                              End Function)
            Catch ex As Exception
                Return "#MSCRMCONNECT! Unable to get list of contacts: " & ex.Message
            End Try
        End If
    End Function

    <Udf(True)>
    <FunctionInfo("Returns contacts from Dynamics CRM. There may be a limit on the amount of data your CRM installation can send. You may need to filter this list or request less fields.", FUNCTION_CATEGORY)>
    Public Function MSCRMSearchContactsWithAccountSearch(<ParamInfo("Search Field", "The name of the field that will be searched. Leave blank to show all results.")> searchField As String,
                                                         <ParamInfo("Search Value", "The value to search for. Results will be returned where the field value begins with the search value.")> searchString As String,
                                                         <ParamInfo("Account Filter Field", "The name of the field in the account entity that will be searched on. Leave blank to show all results.")> accountSearchField As String,
                                                         <ParamInfo("Account Filter Value", "The value in the account to searched on. Results will be returned where the field value in the account matches the search value.")> accountSearchValue As Object,
                                                         <ParamInfo("Contact Field Names", "A pipebar (|) delimited list of field names for the contact. These names need to match those in CRM exactly. Leave blank for all fields.")> contactFieldNames As String,
                                                         <ParamInfo("Account Field Names", "A pipebar (|) delimited list of field names for the account. These names need to match those in CRM exactly. Leave blank for all fields.")> accountFieldNames As String,
                                                         <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntitiesByName As Boolean,
                                                         <ParamInfo("Connect to Dynamics CRM", "False to NOT connect, any other value to connect. Can be used as a way to refresh the data, by simply changing this value.")> connect As Object
                                                         ) As Object

        If IsFalse(connect) Then
            Return ErrorType.ValueException
        Else

            Try
                Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                                  Return New StandardArrayValue(connection.SearchContactsWithAccountSearch(searchField, searchString, accountSearchField, accountSearchValue, contactFieldNames, accountFieldNames, returnRelatedEntitiesByName, True))
                                                              End Function)
            Catch ex As Exception
                Return "#MSCRMCONNECT! Unable to get list of contacts: " & ex.Message
            End Try
        End If
    End Function

    <Udf(True)>
    <FunctionInfo("Returns linked entity information from Dynamics CRM. There may be a limit on the amount of data your CRM installation can send. You may need to filter this list or request less fields.", FUNCTION_CATEGORY)>
    Public Function MSCRMSearchLinkedEntitiesWithFilter(<ParamInfo("Link From Entity Name", "The name of the main entity whose results will be returned.")> linkFromEntityName As String,
                                                        <ParamInfo("Link From Attribute Name", "The attribute name on the main entity that is used to link to the link entity.")> linkFromAttributeName As String,
                                                        <ParamInfo("Search Field", "The name of the field that will be searched. Leave blank to show all results.")> searchField As String,
                                                        <ParamInfo("Search Value", "The value to search for. Results will be returned where the field value begins with the search value.")> searchString As String,
                                                        <ParamInfo("Link From Field Names", "A pipebar (|) delimited list of field names for the Link From entity. These names need to match those in CRM exactly. Leave blank for all fields.")> linkFromFieldNames As String,
                                                        <ParamInfo("Link To Entity Name", "The name of the linked entity.")> linkToEntityName As String,
                                                        <ParamInfo("Link To Attribute Name", "The attribute name of the linked entity that corresponds to the Link From Attribute name on the Link From Entity.")> linkToAttributeName As String,
                                                        <ParamInfo("Link Filter Field", "The name of the field in the link to entity that will be filtered on. Leave blank to show all results")> linkToFilterField As String,
                                                        <ParamInfo("Link To Filter Value", "The value in the link from entity to filter on. Results will be returned where the field value in the Link From entity matches the search value.")> linkToFilterValue As Object,
                                                        <ParamInfo("Link To Field Names", "A pipebar (|) delimited list of field names for the Link To entity. These names need to match those in CRM exactly. Leave blank for all fields.")> linkToFieldNames As String,
                                                        <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntitiesByName As Boolean,
                                                        <ParamInfo("Connect to Dynamics CRM", "False to NOT connect, any other value to connect. Can be used as a way to refresh the data, by simply changing this value.")> connect As Object
                                                        ) As Object

        If IsFalse(connect) Then
            Return ErrorType.ValueException
        Else

            Try
                Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                                  Return New StandardArrayValue(connection.SearchLinkedEntitiesFilter(linkFromEntityName, linkFromAttributeName, searchField, searchString, linkFromFieldNames, linkToEntityName, linkToAttributeName, linkToFilterField, linkToFilterValue, linkToFieldNames, returnRelatedEntitiesByName, True))
                                                              End Function)
            Catch ex As Exception
                Return "#MSCRMCONNECT! Unable to get linked entity data: " & ex.Message
            End Try
        End If
    End Function

    <Udf(True)>
    <FunctionInfo("Returns contacts from Dynamics CRM. There may be a limit on the amount of data your CRM installation can send. You may need to filter this list or request less fields.", FUNCTION_CATEGORY)>
    Public Function MSCRMGetContacts(<ParamInfo("Filter Text", "The filter string. Any contact whose name contains this text will be returned.")> contactNameFilter As String,
                                     <ParamInfo("Field Names", "A pipebar (|) delimited list of field names. These names need to match those in CRM exactly. Leave blank for all fields.")> FieldNames As String,
                                     <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntitiesByName As Boolean
                                     ) As Object

        Return Me.MSCRMGetContacts(contactNameFilter, FieldNames, returnRelatedEntitiesByName, True)

    End Function

    <Udf(True)>
    <FunctionInfo("Returns contacts from Dynamics CRM. There may be a limit on the amount of data your CRM installation can send. You may need to filter this list or request less fields.", FUNCTION_CATEGORY)>
    Public Function MSCRMSearchContacts(<ParamInfo("Search Field", "The field that will be used to search.")> contactSearchField As String,
                                        <ParamInfo("Search String", "The values whose match must be found in the search field.")> contactSearchString As String,
                                        <ParamInfo("Field Names", "A pipebar (|) delimited list of field names. These names need to match those in CRM exactly. Leave blank for all fields.")> FieldNames As String,
                                        <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntitiesByName As Boolean,
                                        <ParamInfo("Connect to Dynamics CRM", "False to NOT connect, any other value to connect. Can be used as a way to refresh the data, by simply changing this value.")> connect As Object
                                        ) As Object

        If IsFalse(connect) Then
            Return ErrorType.ValueException
        Else

            Try
                Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                                  Return New StandardArrayValue(connection.GetContacts(contactSearchField, contactSearchString, returnRelatedEntitiesByName, FieldNames, True))
                                                              End Function)
            Catch ex As Exception
                Return "#MSCRMCONNECT! Unable to get list of contacts: " & ex.Message
            End Try
        End If
    End Function

    <Udf(True)>
    <FunctionInfo("Returns contacts from Dynamics CRM. There may be a limit on the amount of data your CRM installation can send. You may need to filter this list or request less fields.", FUNCTION_CATEGORY)>
    Public Function MSCRMSearchContacts(<ParamInfo("Search Field", "The field that will be used to search.")> contactSearchField As String,
                                        <ParamInfo("Search String", "The values whose match must be found in the search field.")> contactSearchString As String,
                                        <ParamInfo("Field Names", "A pipebar (|) delimited list of field names. These names need to match those in CRM exactly. Leave blank for all fields.")> FieldNames As String,
                                        <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntitiesByName As Boolean
                                        ) As Object

        Return Me.MSCRMSearchContacts(contactSearchField, contactSearchString, FieldNames, returnRelatedEntitiesByName, True)

    End Function

    <Udf(True)>
    <FunctionInfo("Returns Any Entity from Dynamics CRM. There may be a limit on the amount of data your CRM installation can send. You may need to filter this list or request less fields.", FUNCTION_CATEGORY)>
    Public Function MSCRMSearchEntities(<ParamInfo("Entity Name", "The specific entity name (Not the display Name).")> entityName As String,
                                        <ParamInfo("Search Field", "The field that will be used to search.")> contactSearchField As String,
                                        <ParamInfo("Search String", "The values whose match must be found in the search field.")> contactSearchString As String,
                                        <ParamInfo("Field Names", "A pipebar (|) delimited list of field names. These names need to match those in CRM exactly. Leave blank for all fields.")> fieldNames As String,
                                        <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntitiesByName As Boolean,
                                        <ParamInfo("Connect to Dynamics CRM", "False to NOT connect, any other value to connect. Can be used as a way to refresh the data, by simply changing this value.")> connect As Object,
                                        <ParamInfo("Include Inactive Entities", "True to include inactive entities in the search result, False to only return active entities")> includeInactiveEntities As Boolean
                                        ) As Object

        If IsFalse(connect) Then
            Return ErrorType.ValueException
        Else

            Try
                Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                                  Return New StandardArrayValue(connection.GetSpecificEntities(entityName, contactSearchField, contactSearchString, returnRelatedEntitiesByName, fieldNames, Not (includeInactiveEntities)))
                                                              End Function)
            Catch ex As Exception
                Return "#MSCRMCONNECT! Unable to get the specified entities: " & ex.Message
            End Try
        End If
    End Function

    <Udf(True)>
    <FunctionInfo("Returns Any Entity from Dynamics CRM. There may be a limit on the amount of data your CRM installation can send. You may need to filter this list or request less fields.", FUNCTION_CATEGORY)>
    Public Function MSCRMSearchEntities(<ParamInfo("Entity Name", "The specific entity name (Not the display Name).")> entityName As String,
                                        <ParamInfo("Search Field", "The field that will be used to search.")> contactSearchField As String,
                                        <ParamInfo("Search String", "The values whose match must be found in the search field.")> contactSearchString As String,
                                        <ParamInfo("Field Names", "A pipebar (|) delimited list of field names. These names need to match those in CRM exactly. Leave blank for all fields.")> FieldNames As String,
                                        <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntitiesByName As Boolean,
                                        <ParamInfo("Connect to Dynamics CRM", "False to NOT connect, any other value to connect. Can be used as a way to refresh the data, by simply changing this value.")> connect As Object
                                        ) As Object

        If IsFalse(connect) Then
            Return ErrorType.ValueException
        Else

            Try
                Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                                  Return New StandardArrayValue(connection.GetSpecificEntities(entityName, contactSearchField, contactSearchString, returnRelatedEntitiesByName, FieldNames))
                                                              End Function)
            Catch ex As Exception
                Return "#MSCRMCONNECT! Unable to get the specified entities: " & ex.Message
            End Try
        End If
    End Function

    <Udf(True)>
    <FunctionInfo("Returns Any Entity from Dynamics CRM. There may be a limit on the amount of data your CRM installation can send. You may need to filter this list or request less fields.", FUNCTION_CATEGORY)>
    Public Function MSCRMSearchEntities(<ParamInfo("Entity Name", "The specific entity name (not the display name).")> entityName As String,
                                        <ParamInfo("Search Field", "The field that will be used to search.")> contactSearchField As String,
                                        <ParamInfo("Search String", "The values whose match must be found in the search field.")> contactSearchString As String,
                                        <ParamInfo("Field Names", "A pipebar (|) delimited list of field names. These names need to match those in CRM exactly. Leave blank for all fields.")> FieldNames As String,
                                        <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntitiesByName As Boolean
                                        ) As Object

        Return Me.MSCRMSearchEntities(entityName, contactSearchField, contactSearchString, FieldNames, returnRelatedEntitiesByName, True)

    End Function

#End Region

#Region " Accounts "

    <Udf(True)>
    <FunctionInfo("Returns the data for a single account from Dynamics CRM.", FUNCTION_CATEGORY)>
    Public Function MSCRMGetAccountByID(<ParamInfo("Account ID", "The account's GUID.")> accountID As String,
                                        <ParamInfo("Field Names", "A pipebar (|) delimited list of field names. These names need to match those in CRM exactly. Leave blank for all fields.")> FieldNames As String,
                                        <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntitiesByName As Boolean,
                                        <ParamInfo("Connect to Dynamics CRM", "False to NOT connect, any other value to connect. Can be used as a way to refresh the data, by simply changing this value.")> connect As Object
                                        ) As Object

        If IsFalse(connect) Then
            Return ErrorType.ValueException
        Else

            Try
                Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                                  Return New StandardArrayValue(connection.GetAccount(accountID, returnRelatedEntitiesByName, FieldNames))
                                                              End Function)
            Catch ex As Exception
                Return "#MSCRMCONNECT! Unable to get the data for that account: " & ex.Message
            End Try
        End If

    End Function

    <Udf(True)>
    <FunctionInfo("Returns the data for a single account from Dynamics CRM.", FUNCTION_CATEGORY)>
    Public Function MSCRMGetAccountByID(<ParamInfo("Account ID", "The account's GUID.")> accountID As String,
                                        <ParamInfo("Field Names", "A pipebar (|) delimited list of field names. These names need to match those in CRM exactly. Leave blank for all fields.")> FieldNames As String,
                                        <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntitiesByName As Boolean
                                        ) As Object

        Return Me.MSCRMGetAccountByID(accountID, FieldNames, returnRelatedEntitiesByName, True)

    End Function

    <Udf(True)>
    <FunctionInfo("Returns accounts from Dynamics CRM. There may be a limit on the amount of data your CRM installation can send. You may need to filter this list or request less fields. Search is based on the account Name Field", FUNCTION_CATEGORY)>
    Public Function MSCRMGetAccounts(<ParamInfo("Filter Field", "The field name to filter on (leave blank if no filtering is required). Filters will look for an exact match.")> filterField As String,
                                     <ParamInfo("Filter value", "The value on which to base the filter. Filters will look for an exact match.")> filterValue As Object,
                                     <ParamInfo("Search Field", "The field name to search on.")> searchField As String,
                                     <ParamInfo("Search Value", "The search value.")> searchValue As String,
                                     <ParamInfo("Field Names", "A pipebar (|) delimited list of field names. These names need to match those in CRM exactly. Leave blank for all fields.")> FieldNames As String,
                                     <ParamInfo("Get Related Entities by Name", "True to return the name of related entities, False to return their ID (GUID).")> returnRelatedEntitiesByName As Boolean,
                                     <ParamInfo("Connect to Dynamics CRM", "False to NOT connect, any other value to connect. Can be used as a way to refresh the data, by simply changing this value.")> connect As Object
                                     ) As Object

        If IsFalse(connect) Then
            Return ErrorType.ValueException
        Else

            Try
                Return Me.GetConnectionManager().RunWithRetry(Function(connection)
                                                                  Return New StandardArrayValue(connection.GetAccounts(filterField, filterValue, searchField, searchValue, FieldNames, returnRelatedEntitiesByName, True))
                                                              End Function)
            Catch ex As Exception
                Return "#MSCRMCONNECT! Unable to get list of accounts: " & ex.Message
            End Try
        End If
    End Function

#End Region

    Private Shared Function IsFalse(valueToTest As Object) As Boolean

        If TypeOf valueToTest Is String Then
            If valueToTest.ToString().Equals("false", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        ElseIf TypeOf valueToTest Is Boolean Then
            If DirectCast(valueToTest, Boolean) = False Then
                Return True
            End If
        End If

        Return False

    End Function

    Private Function GetConnectionManager() As ConnectionManager
        Return Connections.ConnectionInstance.GetConnection(Me.Project)
    End Function

End Class