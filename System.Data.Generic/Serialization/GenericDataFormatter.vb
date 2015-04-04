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

'Imports System
'Imports System.Collections.Generic
'Imports System.Text
'Imports System.Xml.Serialization
'Imports System.Runtime.Serialization
'Imports System.Runtime.Serialization.Formatters
'Imports System.Runtime.Serialization.Formatters.Binary
'Imports System.Reflection
'Imports System.IO
'Imports System.Collections.Specialized

' ''' <summary>
' ''' Generic data formatter, helps to serialize and deserialize objects.
' ''' </summary>
' ''' <typeparam name="T">Type of object to format.</typeparam>
' ''' <remarks></remarks>
'Friend NotInheritable Class GenericDataFormatter(Of T)

'#Region "Member variables"
'    Private _reflectedType As Type
'#End Region

'#Region "Constructor"
'    ''' <summary>
'    ''' Returns a new instance of the GenericDataFormatter class.
'    ''' </summary>
'    ''' <remarks></remarks>
'    Public Sub New()

'    End Sub

'    ''' <summary>
'    ''' Returns a new instance of the GenericDataFormatter class.
'    ''' </summary>
'    ''' <param name="enforceStrictModel"></param>
'    ''' <remarks></remarks>
'    Public Sub New(ByVal enforceStrictModel As Boolean)
'        ' enforce requires that field in class and database is the same
'        ' not implemented yet
'    End Sub
'#End Region

'#Region "Protected properties"
'    ''' <summary>
'    ''' Returns the type of the reflected object.
'    ''' </summary>
'    ''' <value></value>
'    ''' <returns></returns>
'    ''' <remarks></remarks>
'    Protected Friend ReadOnly Property ReflectedType As Type
'        Get
'            If _reflectedType Is Nothing Then
'                _reflectedType = GetType(T)
'            End If

'            Return _reflectedType
'        End Get
'    End Property
'#End Region

'#Region "Public methods"
'    ''' <summary>
'    ''' Not implmented.
'    ''' </summary>
'    ''' <param name="reader"></param>
'    ''' <param name="value"></param>
'    ''' <remarks></remarks>
'    Public Sub Serialize(ByVal reader As IDataReader, value As T)
'        Throw New NotImplementedException
'    End Sub

'    ''' <summary>
'    ''' 
'    ''' </summary>
'    ''' <param name="reader"></param>
'    ''' <returns></returns>
'    ''' <remarks></remarks>
'    <Obsolete("Please use the DataSourceFormatter")>
'    Public Function DeSerialize(ByVal reader As IDataReader) As T

'        Dim fields() As InnerFieldInfo = ExtendedFormatterServices.GetSerializableMembers(ReflectedType)

'        ' set new instance
'        Dim instance As Object = FormatterServices.GetUninitializedObject(ReflectedType)

'        ' populare data values
'        For Each field As InnerFieldInfo In fields
'            ' might cause an error if XmlAttribute is set to wrong field name
'            Try
'                If ContainsField(reader, field.DatabaseFieldName) Then
'                    ' value of the database
'                    Dim dataValue As Object = reader(field.DatabaseFieldName)

'                    ' check conversion
'                    If Not String.IsNullOrEmpty(field.TargetType) Then
'                        Dim t1 As Type = Type.GetType(field.TargetType)
'                        Dim t2 As Type = dataValue.GetType

'                        ' convert value from the field, if the two types are not equal
'                        If t1 IsNot t2 Then
'                            ConvertValue(dataValue, field.TargetType)
'                        End If
'                    End If

'                    ' populate value and field
'                    ExtendedFormatterServices.PopulareFieldValue(instance,
'                                                                     field.FieldName, dataValue)
'                End If
'            Catch ex As FormatException
'                ' thrown on ConvertValue
'                Throw ex
'            Catch ex As Exception
'                Throw ex
'            End Try
'        Next

'        Return instance
'    End Function
'#End Region

'#Region "Private methods"
'    ''' <summary>
'    ''' 
'    ''' </summary>
'    ''' <param name="reader"></param>
'    ''' <param name="fieldName"></param>
'    ''' <returns></returns>
'    ''' <remarks></remarks>
'    Private Function ContainsField(ByVal reader As IDataReader, ByVal fieldName As String) As Boolean
'        For i As Integer = 0 To reader.FieldCount - 1
'            Dim n As String = reader.GetName(i)
'            Dim c As Integer = String.Compare(fieldName, n, True)
'            If String.Compare(fieldName, n, True) = 0 Then
'                Return True
'            End If
'        Next

'        Return False
'    End Function

'    ''' <summary>
'    ''' Converts the type from the datasource to the target object type (CLR).
'    ''' </summary>
'    ''' <param name="value">Value from the SQL datasource.</param>
'    ''' <param name="targetType">Target CLR type.</param>
'    ''' <remarks>The method handles nullable fields, defaults and values.</remarks>
'    Private Sub ConvertValue(ByRef value As Object, targetType As String)
'        Select Case targetType            
'            Case GetType(String).ToString
'                value = Convert.ToString(value)
'            Case GetType(Guid).ToString
'                If IsDBNull(value) Then
'                    value = Guid.Empty
'                Else
'                    value = New Guid(value.ToString)
'                End If
'            Case GetType(Long).ToString
'                value = CLng(value)
'            Case GetType(Char).ToString
'                value = CChar(value)
'            Case GetType(Byte).ToString
'                value = CByte(value)
'            Case GetType(Single).ToString
'                value = CSng(value)
'            Case GetType(Short).ToString
'                value = CShort(value)
'            Case GetType(Integer).ToString
'                value = CInt(value)
'            Case GetType(Decimal).ToString
'                value = CDec(value)
'            Case GetType(Double).ToString
'                value = CDbl(value)
'            Case GetType(DateTime).ToString
'                If IsDBNull(value) Then
'                    value = DateTime.MinValue
'                Else
'                    value = CType(value, DateTime)
'                End If
'            Case Else
'                ' do default
'                value = value.ToString
'        End Select
'    End Sub
'#End Region

'#Region "Friend Classes"
'    ''' <summary>
'    ''' Data class for field information.
'    ''' </summary>
'    ''' <remarks></remarks>
'    Friend Class InnerFieldInfo
'        'Public Property Ignore As Boolean
'        Public Property DatabaseFieldName As String
'        Public Property FieldName As String
'        Public Property TargetType As String
'    End Class

'    ''' <summary>
'    ''' 
'    ''' </summary>
'    ''' <remarks></remarks>
'    Friend NotInheritable Class ExtendedFormatterServices
'        ''' <summary>
'        ''' 
'        ''' </summary>
'        ''' <param name="t"></param>
'        ''' <returns></returns>
'        ''' <remarks></remarks>
'        Protected Friend Shared Function GetSerializableMembers(ByVal t As Type) As InnerFieldInfo()
'            If t Is Nothing Then
'                Throw New ArgumentNullException("type")
'            End If

'            Dim fields As ArrayList = New ArrayList
'            GetFields(t, fields)

'            Dim result As InnerFieldInfo() = New InnerFieldInfo(fields.Count - 1) {}
'            fields.CopyTo(result)

'            Return result
'        End Function

'        ''' <summary>
'        ''' 
'        ''' </summary>
'        ''' <param name="t"></param>
'        ''' <param name="fields"></param>
'        ''' <remarks></remarks>
'        Private Shared Sub GetFields(ByVal t As Type, ByRef fields As ArrayList)
'            For Each prop As PropertyInfo In t.GetProperties()
'                If prop.CanWrite AndAlso Not IgnoreField(prop) Then

'                    'fields.Add(New InnerFieldInfo With {
'                    '    .DatabaseFieldName = GetAttributeFieldName(prop),
'                    '    .FieldName = prop.Name})

'                    fields.Add(GetInnerFieldInfo(prop))
'                End If
'            Next
'        End Sub

'        ''' <summary>
'        ''' 
'        ''' </summary>
'        ''' <param name="obj"></param>
'        ''' <param name="fieldName"></param>
'        ''' <param name="value"></param>
'        ''' <remarks></remarks>
'        Protected Friend Shared Sub PopulareFieldValue(ByRef obj As Object, ByVal fieldName As String, ByVal value As Object)
'            obj.GetType().GetProperty(fieldName).SetValue(obj, value, Nothing)
'        End Sub

'        ''' <summary>
'        ''' 
'        ''' </summary>
'        ''' <param name="p"></param>
'        ''' <returns></returns>
'        ''' <remarks></remarks>
'        Private Shared Function IgnoreField(ByVal p As PropertyInfo) As Boolean
'            Return p.GetCustomAttributes(GetType(XmlIgnoreAttribute), True).Length > 0
'        End Function

'        ''' <summary>
'        ''' 
'        ''' </summary>
'        ''' <param name="prop"></param>
'        ''' <returns></returns>
'        ''' <remarks></remarks>
'        Private Shared Function GetInnerFieldInfo(ByVal prop As PropertyInfo) As InnerFieldInfo            
'            Dim returnValue As InnerFieldInfo = New InnerFieldInfo() With {.FieldName = prop.Name}
'            PopulareInnerFieldInfo(returnValue, prop)

'            'Return New InnerFieldInfo With {
'            '            .DatabaseFieldName = GetAttributeFieldName(prop),
'            '            .FieldName = prop.Name}

'            Return returnValue
'        End Function

'        Private Shared Sub PopulareInnerFieldInfo(ByRef innerField As InnerFieldInfo, ByVal p As PropertyInfo)
'            Dim att() As XmlAttributeAttribute = p.GetCustomAttributes(GetType(XmlAttributeAttribute), True)
'            If att.Length > 0 Then
'                'If att(0).Type IsNot Nothing Then
'                '    Dim h As Type = p.PropertyType
'                '    innerField.TargetType = att(0).Type.ToString
'                'End If
'                If Not String.IsNullOrEmpty(att(0).AttributeName) Then
'                    innerField.DatabaseFieldName = att(0).AttributeName
'                    Exit Sub
'                End If
'            End If

'            innerField.DatabaseFieldName = p.Name
'            If p.PropertyType.IsEnum Then
'                Dim t As Type = System.Enum.GetUnderlyingType(p.PropertyType)
'                innerField.TargetType = t.ToString
'            Else
'                innerField.TargetType = p.PropertyType.ToString
'            End If
'        End Sub


'        'Private Shared Function GetAttributeFieldName(ByVal p As PropertyInfo) As String
'        '    Dim att() As XmlAttributeAttribute = p.GetCustomAttributes(GetType(XmlAttributeAttribute), True)
'        '    If att.Length > 0 Then
'        '        If att(0).Type IsNot Nothing Then

'        '        End If
'        '        Return att(0).AttributeName
'        '    Else
'        '        Return p.Name
'        '    End If
'        'End Function
'    End Class
'#End Region

'End Class