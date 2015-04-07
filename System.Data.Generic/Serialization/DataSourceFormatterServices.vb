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
    ''' Provides static methods to aid with the implementation of any <see cref="IDataReaderFormatter">IDataReaderFormatter</see> for serialization. 
    ''' </summary>
    ''' <remarks></remarks>
    Friend NotInheritable Class DataSourceFormatterServices
        ''' <summary>
        ''' Populares members to an object.
        ''' </summary>
        ''' <param name="obj">Object to populate.</param>
        ''' <param name="members">Properties to populare.</param>
        ''' <param name="data">Data array to pass to the properties.</param>
        ''' <remarks>Both array must have the same index.</remarks>
        Friend Shared Sub PopulateObjectMembers(ByRef obj As Object, members As PropertyInfo(), data As Object())
            For i As Integer = 0 To members.Length - 1
                Try
                    Dim p As PropertyInfo = obj.GetType().GetProperty(members(i).Name)
                    If p.CanWrite Then
                        p.SetValue(obj, data(i), Nothing)
                    End If
                Catch ex As Exception
                    Throw ex
                End Try
            Next
        End Sub

        ''' <summary>
        ''' Validates the property and determine if the property should be ignored.
        ''' </summary>
        ''' <param name="p">Property to validate.</param>
        ''' <returns></returns>
        ''' <remarks>Non-writable properties are ignored. And users may specify either the <see cref="Xml.Serialization.XmlIgnoreAttribute">XmlIgnoreAttribute</see> or <see cref="IgnorePropertyAttribute">IgnorePropertyAttribute</see> attribute to force exclusion.</remarks>
        Private Shared Function IgnoreProperty(p As PropertyInfo) As Boolean
            ' Exclude read-only
            If Not p.CanWrite Then
                Return True
            End If
            ' Check for XmlIgnore on the public field
            If (p.GetCustomAttributes(GetType(XmlIgnoreAttribute), True).Length > 0) Then
                Return True
            End If
            If (p.GetCustomAttributes(GetType(IgnorePropertyAttribute), True).Length > 0) Then
                Return True
            End If

            Return False
        End Function

        <Obsolete()>
        Friend Shared Function GetSafeInitializedObject(t As Type) As Object
            Dim instance As Object = FormatterServices.GetSafeUninitializedObject(t)

            Return instance
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="type"></param>
        ''' <param name="navigateProperties">Recruisive navigation through collection properties.</param>
        ''' <param name="list"></param>
        ''' <remarks></remarks>
        Friend Shared Sub GetSerializableProperties(type As Type, navigateProperties As Boolean, ByRef list As List(Of PropertyInfo))
            If list Is Nothing Then
                list = New List(Of PropertyInfo)
            End If

            For Each p As PropertyInfo In DataSourceFormatterServices.GetSerializableProperties(type)
                list.Add(p)
                If p.IsCollection() AndAlso navigateProperties Then
                    GetSerializableProperties(p.CollectionItemType, navigateProperties, list)
                End If
            Next
        End Sub

        ''' <summary>
        ''' Returns an array of properties based on the specified type.
        ''' </summary>
        ''' <param name="type">Object type to get properties from.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Friend Shared Function GetSerializableProperties(type As Type) As PropertyInfo()
            Dim props As PropertyInfo() = type.GetProperties(BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
            Dim countProper As Integer = 0

            For i As Integer = 0 To props.Length - 1
                Dim prop As PropertyInfo = props(i)
                If IgnoreProperty(prop) Then
                    Continue For
                End If
                countProper += 1
            Next

            If (countProper <> props.Length) Then
                ' render the reduced list
                Dim properFields As PropertyInfo() = New PropertyInfo(countProper - 1) {}
                countProper = 0
                For i As Integer = 0 To props.Length - 1
                    If IgnoreProperty(props(i)) Then
                        Continue For
                    End If
                    properFields(countProper) = props(i)
                    countProper += 1
                Next

                Return properFields
            Else
                Return props
            End If
        End Function

        ''' <summary>
        ''' Get all the serializable members for a class of the specified System.Type.
        ''' </summary>
        ''' <param name="type">The type being serialized.</param>
        ''' <returns></returns>
        ''' <remarks>This is an extended version of the <see cref="FormatterServices.GetSerializableMembers">FormatterServices.GetSerializableMembers</see> method. This implementation does not required the class to be marked as Serialize, rather - supports the <see cref="XmlIgnoreAttribute">XmlIgnoreAttribute</see> for public fields to exclude.</remarks>
        Friend Shared Function GetSerializableMembers(type As Type) As FieldInfo()
            Dim fields As FieldInfo() = type.GetFields(BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance)
            Dim countProper As Integer = 0

            For i As Integer = 0 To fields.Length - 1
                Dim field As FieldInfo = fields(i)
                ' Check for NonSerializable on the member field
                If (field.Attributes And FieldAttributes.NotSerialized) = FieldAttributes.NotSerialized Then
                    Continue For
                End If
                ' Check for XmlIgnore on the public field
                If (field.GetCustomAttributes(GetType(XmlIgnoreAttribute), True).Length > 0) Then
                    Continue For
                End If

                countProper += 1
            Next

            If (countProper <> fields.Length) Then
                ' render the reduced list
                Dim properFields As FieldInfo() = New FieldInfo(countProper - 1) {}
                countProper = 0
                For i As Integer = 0 To fields.Length - 1
                    If (fields(i).Attributes And FieldAttributes.NotSerialized) = FieldAttributes.NotSerialized Then
                        Continue For
                    End If
                    If (fields(i).GetCustomAttributes(GetType(XmlIgnoreAttribute), True).Length > 0) Then
                        Continue For
                    End If
                    properFields(countProper) = fields(i)
                    countProper += 1
                Next

                Return properFields
            Else
                Return fields
            End If
        End Function

        ''' <summary>
        ''' Returns an array of data field information.
        ''' </summary>
        ''' <param name="reader">DataReader to read from.</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetFieldMembers(reader As IDataReader) As DbColumnInfo()
            Dim fields As DbColumnInfo() = New DbColumnInfo(reader.FieldCount - 1) {}

            For i As Integer = 0 To reader.FieldCount - 1
                '    .FieldIndex = reader.GetOrdinal(reader.GetName(i)),
                fields(i) = New DbColumnInfo With {
                    .FieldName = reader.GetName(i),
                    .FieldIndex = reader.GetOrdinal(.FieldName),
                    .FieldValue = reader(i),
                    .DataTypeName = reader.GetDataTypeName(i),
                    .RuntimeType = reader.GetFieldType(i)
                }
            Next

            Return fields
        End Function
    End Class
End Namespace
