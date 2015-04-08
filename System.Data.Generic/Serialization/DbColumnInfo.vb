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

Namespace Serialization
    ''' <summary>
    ''' Obtains information about the values of a data column from the database.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class DbColumnInfo

#Region "Nested implementation"
        ''' <summary>
        ''' Gets or sets the table name of the column.
        ''' </summary>
        ''' <remarks></remarks>
        Public TableName As String = String.Empty
        ''' <summary>
        ''' Gets or sets the column index.
        ''' </summary>
        ''' <remarks></remarks>
        Public FieldIndex As Integer = -1
        ''' <summary>
        ''' Returns the path of the column value.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property FieldPath As String
            Get
                If String.IsNullOrEmpty(TableName) Then
                    Return FieldName
                Else
                    Return TableName + "." + FieldName
                End If
            End Get
        End Property
#End Region

#Region "Properties"        
        ''' <summary>
        ''' Name of the datasource field.
        ''' </summary>
        ''' <remarks></remarks>
        Public FieldName As String
        ''' <summary>
        ''' Value of the datasource field.
        ''' </summary>
        ''' <remarks></remarks>
        Public FieldValue As Object
        ''' <summary>
        ''' Name of the data source field type.
        ''' </summary>
        ''' <remarks></remarks>
        Public DataTypeName As String
        ''' <summary>
        ''' System.Type of the field, reflected by CLR.
        ''' </summary>
        ''' <remarks></remarks>
        Public RuntimeType As Type
#End Region

        Public Overrides Function ToString() As String
            Return Me.FieldName + " (" + Me.FieldPath + ")"
        End Function
#Region "Read Methods"
        ''' <summary>
        ''' Returns a long from the field value.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Function ReadLong() As Long
            Return CType(FieldValue, Long)
        End Function
        ''' <summary>
        ''' Returns a integer from the field value.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Function ReadInt() As Integer
            Return CType(FieldValue, Integer)
        End Function
        ''' <summary>
        ''' Returns a boolean from the field value.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Function ReadBool() As Boolean
            Return CBool(FieldValue)
        End Function
        ''' <summary>
        ''' Returns a guid from the field value.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Function ReadGuid() As Guid
            If IsDBNull(FieldValue) Then
                FieldValue = Guid.Empty
            End If
            Return New Guid(FieldValue.ToString)
        End Function
        ''' <summary>
        ''' Returns a string from the field value.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Function ReadString() As String
            Return FieldValue.ToString
        End Function
        ''' <summary>
        ''' Returns a datetime from the field value.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Function ReadDateTime() As DateTime
            If IsDBNull(FieldValue) Then
                FieldValue = DateTime.MinValue
            End If
            Return CType(FieldValue, DateTime)
        End Function
        ''' <summary>
        ''' Returns a decimal from the field value.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Function ReadDecimal() As Decimal
            Return CDec(FieldValue)
        End Function
        ''' <summary>
        ''' Returns a char from the field value.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Function ReadChar() As Char
            Return CChar(FieldValue)
        End Function
        ''' <summary>
        ''' Returns a single from the field value.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Function ReadSingle() As Single
            Return CType(FieldValue, Single)
        End Function
        ''' <summary>
        ''' Returns a decimal from the field value.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Function ReadDouble() As Double
            Return CDbl(FieldValue)
        End Function
        ''' <summary>
        ''' Returns a short from the field value.
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Friend Function ReadInt16() As Int16
            Return CShort(FieldValue)
        End Function
#End Region
    End Class
End Namespace