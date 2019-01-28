Imports Microsoft.Xrm.Sdk
Imports Microsoft.Xrm.Sdk.Query

Friend Module QueryController

    Function ProcessFields(specifiedFields As String, defaultFields As String) As String()

        ' Check if we have any specified fields, if we do, use them.
        If Not String.IsNullOrEmpty(specifiedFields) Then
            Return specifiedFields.Split(CChar("|"))
        Else
            ' We don't have any existing fields, check we have some default ones.
            If String.IsNullOrEmpty(defaultFields) Then
                Return New String() {}
            End If

            ' We do have some default ones, create a new list from them and return.
            Return defaultFields.Split(CChar("|"))
        End If

    End Function

    Function GetQueryExpression(entityname As String, fieldNames As String(), filter As FilterExpression, link As LinkEntity, activeOnly As Boolean) As QueryExpression

        Dim thisQuery As New QueryExpression
        Dim queryColumns As New ColumnSet
        Dim thisFilter As New FilterExpression

        thisQuery.EntityName = entityname

        If fieldNames.Count = 0 Then
            queryColumns = New ColumnSet(True)
        Else
            queryColumns.Columns.AddRange(fieldNames)
            queryColumns.AllColumns = False
        End If

        thisQuery.ColumnSet = queryColumns

        If Not filter Is Nothing Then
            If filter.Conditions.Count > 0 Then
                thisFilter.Filters.Add(filter)
            End If
        End If

        ' Now add a new condition if we are only showing active entities.
        If activeOnly Then
            Dim activeCondition As ConditionExpression

            activeCondition = New ConditionExpression With {
                .AttributeName = "statuscode",
                .Operator = ConditionOperator.Equal
            }
            activeCondition.Values.Add("1")
            thisFilter.Conditions.Add(activeCondition)
        End If

        thisQuery.Criteria = thisFilter

        If link IsNot Nothing Then
            thisQuery.LinkEntities.Add(link)
        End If

        Return thisQuery

    End Function

    Function GetValueString(attributeValue As Object, returnRelatedEntitiesByName As Boolean) As String

        Dim value As Object
        Dim normalisedValue As Object = Nothing
        Dim valueString As String = String.Empty

        If TypeOf attributeValue Is OptionSetValue Then
            Dim optionSet As OptionSetValue = DirectCast(attributeValue, OptionSetValue)

            value = optionSet.Value
        ElseIf TypeOf attributeValue Is EntityReference Then
            Dim entityRef As EntityReference = DirectCast(attributeValue, EntityReference)

            If Not returnRelatedEntitiesByName Then
                value = entityRef.Id
            Else
                value = entityRef.Name
            End If
        Else
            value = attributeValue
        End If

        If TypeOf value Is AliasedValue Then

            normalisedValue = DirectCast(value, AliasedValue).Value

            If TypeOf normalisedValue Is OptionSetValue Then
                Dim optionSet As OptionSetValue = DirectCast(normalisedValue, OptionSetValue)

                valueString = optionSet.Value.ToString()
            ElseIf TypeOf normalisedValue Is EntityReference Then
                Dim entityRef As EntityReference = DirectCast(normalisedValue, EntityReference)

                If Not returnRelatedEntitiesByName Then
                    valueString = entityRef.Id.ToString()
                Else
                    valueString = entityRef.Name
                End If
            Else
                valueString = normalisedValue.ToString()
            End If

        Else
            If Not value Is Nothing Then
                valueString = value.ToString()
            End If

        End If

        Return valueString

    End Function

    Function SetHeadersDictionaryFromAttributes(attributes As AttributeCollection) As Dictionary(Of Integer, String)

        Dim returnDic As New Dictionary(Of Integer, String)
        Dim i As Integer = 0

        For Each att In attributes
            returnDic.Add(i, att.Key)
            i = i + 1
        Next

        Return returnDic

    End Function

    Function LoadDictionary(entityResults As EntityCollection, ByRef headerDictionary As Dictionary(Of Integer, String), returnRelatedEntitiesByName As Boolean) As Dictionary(Of Integer, Dictionary(Of Integer, String))

        Dim resultsDictionary As New Dictionary(Of Integer, Dictionary(Of Integer, String))
        Dim attributeColumnIndex As Integer = -1
        Dim numberOfEntities As Integer = entityResults.Entities.Count
        Dim resultsRowNumber As Integer = 1

        ' Iterate through all entities.
        For Each entityResult In entityResults.Entities
            Dim valueDictionary As New Dictionary(Of Integer, String)

            ' Use attributes to build headers if non specified.
            If headerDictionary.Count = 0 Then
                headerDictionary = QueryController.SetHeadersDictionaryFromAttributes(entityResult.Attributes)
            End If

            ' Iterate through all attributes.
            For Each att In entityResult.Attributes
                attributeColumnIndex = -1

                For Each existingHeader In headerDictionary
                    If att.Key = existingHeader.Value Then
                        attributeColumnIndex = existingHeader.Key
                        Exit For
                    End If
                Next

                If attributeColumnIndex > -1 Then
                    ' Get the attribute value and add to collection.
                    Dim valueString = QueryController.GetValueString(att.Value, returnRelatedEntitiesByName)
                    valueDictionary.Add(attributeColumnIndex, valueString)
                End If

            Next

            resultsDictionary.Add(resultsRowNumber, valueDictionary)
            resultsRowNumber = resultsRowNumber + 1

        Next

        Return resultsDictionary

    End Function

End Module