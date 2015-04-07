' Copyright (c) 2013-15 Arumsoft
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

Imports System.Reflection
Imports System.Xml.Serialization
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters

Namespace Serialization.Formatters
    ''' <summary>
    ''' Provides a serialization formatter for the data source.
    ''' </summary>
    ''' <typeparam name="T">Type of object to serialize or deserialize.</typeparam>
    ''' <remarks></remarks>
    Public NotInheritable Class DataSourceFormatter(Of T)
        Implements IDataReaderFormatter(Of T)

        ''' <summary>
        ''' Deserialize a data row into an object of type T, navigating the properties and related tables from the reader.
        ''' </summary>
        ''' <param name="reader">DataReader to return row from.</param>
        ''' <param name="schema">Schema of the data reader.</param>
        ''' <returns>Returns object instance of T.</returns>
        ''' <remarks></remarks>
        Public Function DeserializeNested(reader As IDataReader, schema As DbSchema) As T Implements IDataReaderFormatter(Of T).DeserializeNested
            ' get the type list nested
            Dim members As List(Of PropertyInfo) = Nothing
            DataSourceFormatterServices.GetSerializableProperties(GetType(T), True, members)

            ' instanciate a new version
            Dim instance As Object = FormatterServices.GetSafeUninitializedObject(GetType(T))
            instance = instance


            For Each t As DbTableInfo In schema.Tables
                t = t
            Next


            Return instance


            'Dim members() As PropertyInfo = DataSourceFormatterServices.GetSerializableProperties(GetType(T))
            ''Dim u = members(0).PropertyPath
            'u = u

            For Each m As PropertyInfo In members
                Console.WriteLine(m.PropertyPath)
            Next

            Throw New NotImplementedException

            ' retrieve property info array for the type and navigate any child properties
            ' that is not of IsSystemType
            '    Dim nestedMembers() As PropertyInfoList

            '     Dim schemaDt = dataReader.GetSchemaTable()


        End Function

        ''' <summary>
        ''' Deserialize a data row from the reader.
        ''' </summary>
        ''' <param name="reader">DataReader to return row from.</param>
        ''' <returns>Object of type T, deserialized from the provided DataReader.</returns>
        ''' <remarks></remarks>
        Public Function Deserialize(reader As IDataReader) As T Implements IDataReaderFormatter(Of T).Deserialize
            ' retrieve property info array for the type
            Dim members() As PropertyInfo = DataSourceFormatterServices.GetSerializableProperties(GetType(T))

            ' instanciate a new version
            Dim instance As Object = FormatterServices.GetSafeUninitializedObject(GetType(T)) ' .GetUninitializedObject(GetType(T))

            Try
                ' get database fields
                Dim fields() As DbColumnInfo = DataSourceFormatterServices.GetFieldMembers(reader)

                ' maps properties and fields and return data
                Dim data() As Object = MapAndDeserializePrimities(members, fields)

                ' populate the instance, properties and data
                DataSourceFormatterServices.PopulateObjectMembers(instance, members, data)
            Catch ex As Exception
                Throw New SerializationException("Deserialize threw an exception.", ex)
            End Try

            Return instance
        End Function

        ''' <summary>
        ''' Maps fields and properties in the class.
        ''' </summary>
        ''' <param name="members">Properties in the class.</param>
        ''' <param name="fields">Data fields from the data source.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function MapAndDeserializePrimities(members() As MemberInfo, fields() As DbColumnInfo) As Object()
            Dim retVal As Object() = New Object(members.Length - 1) {}

            For i As Integer = 0 To members.Length - 1
                'members(i).Name 
                For Each dataField As DbColumnInfo In fields
                    If String.Compare(dataField.FieldName, members(i).Name, True) = 0 Then
                        retVal(i) = DeserializePrimitive(dataField)
                        Exit For
                    End If
                Next
            Next

            Return retVal
        End Function

        ''' <summary>
        ''' Deserializes a data field into a class.
        ''' </summary>
        ''' <param name="field">Field to deserialize.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function DeserializePrimitive(field As DbColumnInfo) As Object
            Dim fieldType As Type = field.RuntimeType
            Select Case Type.GetTypeCode(fieldType)
                Case TypeCode.Char
                    Return field.ReadChar
                Case TypeCode.Single
                    Return field.ReadSingle
                Case TypeCode.Double
                    Return field.ReadDouble
                Case TypeCode.Decimal
                    Return field.ReadDecimal
                Case TypeCode.Int64, TypeCode.UInt64
                    Return field.ReadLong
                Case TypeCode.Int16, TypeCode.Int32, TypeCode.UInt16, TypeCode.UInt32
                    Return field.ReadInt
                Case TypeCode.Boolean
                    Return field.ReadBool
                Case TypeCode.String
                    Return field.ReadString()
                Case TypeCode.DateTime
                    Return field.ReadDateTime
                Case TypeCode.Empty
                    Throw New ArgumentNullException
                Case TypeCode.DBNull
                    Throw New ArgumentNullException
                Case TypeCode.Object
                    If fieldType Is GetType(Guid) Then
                        Return field.ReadGuid
                    ElseIf fieldType.IsEnum Then
                        Throw New NotImplementedException
                    Else
                        Return field.FieldValue
                    End If
                Case Else
                    Return field.FieldValue
            End Select
        End Function

        ''' <summary>
        ''' Serialize a datareader. This method is not implemented.
        ''' </summary>
        ''' <param name="reader">DataReader to serialize from.</param>
        ''' <param name="graph">Graph object to serialize.</param>
        ''' <remarks></remarks>
        Public Sub Serialize(reader As IDataReader, graph As T) Implements IDataReaderFormatter(Of T).Serialize
            Throw New NotImplementedException
        End Sub
    End Class
End Namespace
