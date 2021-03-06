﻿' Copyright (c) 2013-15 Arumsoft
' All rights reserved.
'
' Created:  21/01/2013.         Author:   Cloudcoder
'
' This source is subject to the Microsoft Permissive License. http://www.microsoft.com/en-us/openness/licenses.aspx
'
' THE CODE IS PROVIDED ON AN "AS IS" BASIS WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED. YOU EXPRESSLY AGREE 
' THAT THE USE OF THE SERVICE IS AT YOUR SOLE RISK. COMPANY DOES NOT WARRANT THAT THE SERVICE WILL BE UNINTERRUPTED 
' OR ERROR FREE, NOR DOES COMPANY MAKE ANY WARRANTY AS TO ANY RESULTS THAT MAY BE OBTAINED BY USE OF THE SERVICE. 
' COMPANY MAKES NO OTHER WARRANTIES, EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, ANY IMPLIED WARRANTIES OF 
' MERCHANTABILITY OR FITNESS FOR A PARTICULAR PURPOSE, IN RELATION TO THE SERVICE.

' TODO: should be Schema
Namespace Serialization
    ''' <summary>
    ''' Obtains information about the tables and columns of a database.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class DbSchema

#Region "Member variables"
        Private _tables As DbTableInfo()
        'Private _columns As DbColumnInfo()
#End Region

#Region "Constructor"
        ' <param name="columns">Column array of <see cref="DbColumnInfo">DbColumnInfo</see>.</param>

        ''' <summary>
        ''' Returns a new instance of the DbSchema class.
        ''' </summary>
        ''' <param name="tables">Table array of <see cref="DbTableInfo">DbTableInfo</see>.</param>
        ''' <remarks></remarks>
        Protected Friend Sub New(tables() As DbTableInfo)
            _tables = tables
            '_columns = columns
            ', columns() As DbColumnInfo
        End Sub
#End Region

#Region "Protected properties"
        Protected Friend ReadOnly Property Tables As DbTableInfo()
            Get
                Return _tables
            End Get
        End Property

        'Protected Friend ReadOnly Property Columns As DbColumnInfo()
        '    Get
        '        Return _columns
        '    End Get
        'End Property
#End Region

#Region "Protected methods"
        ''' <summary>
        ''' Returns the schema from the reader.
        ''' </summary>
        ''' <param name="reader">IDataReader to return schema from.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Shared Function GetSchema(reader As IDataReader) As DbSchema
            Dim schemaTable As DataTable = reader.GetSchemaTable

            ' define the temp. lists
            Dim schemaTables As Dictionary(Of String, DbTableInfo) = New Dictionary(Of String, DbTableInfo)
            ' Dim schemaColumns As List(Of DbColumnInfo) = New List(Of DbColumnInfo)

            ' loop through the schema and retrieve the information of table and columns
            For Each schemaRow As DataRow In schemaTable.Rows
                Dim dbColumnIndex As Integer = schemaRow("ColumnOrdinal")
                Dim dbTableName As String = schemaRow("BaseTableName")

                Dim currentTable As DbTableInfo = Nothing
                If Not schemaTables.ContainsKey(dbTableName.ToLower) Then
                    currentTable = New DbTableInfo With {.TableName = dbTableName}
                    schemaTables.Add(dbTableName.ToLower, currentTable)
                Else
                    currentTable = schemaTables(dbTableName.ToLower)
                End If

                Dim d1 As Object = schemaRow("ProviderType")
                Dim d2 As SqlDbType = CType(d1, SqlDbType)

                currentTable.Columns.Add(New DbColumnInfo() With {
                                  .TableName = dbTableName,
                                  .FieldIndex = dbColumnIndex,
                                  .FieldName = schemaRow("BaseColumnName"),
                                  .FieldValue = Nothing,
                                  .DataTypeName = d2.ToString,
                                  .RuntimeType = schemaRow("DataType")
                    })

                ' try to determine the primary key of the table
                If schemaRow("IsKey") Then
                    currentTable.KeyColumnIndex = schemaRow("ColumnOrdinal")
                End If

            Next

            Return New DbSchema(schemaTables.ValuesToArray)
        End Function
#End Region
    End Class
End Namespace
