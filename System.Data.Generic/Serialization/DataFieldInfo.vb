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
    ''' Obtains information about the values of a data field.
    ''' </summary>
    ''' <remarks></remarks>
    Friend NotInheritable Class DataFieldInfo

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

#Region "Read Methods"
        Public Function ReadLong() As Long
            Return CType(FieldValue, Long)
        End Function
        Public Function ReadInt() As Integer
            Return CType(FieldValue, Integer)
        End Function
        Public Function ReadBool() As Boolean
            Return CBool(FieldValue)
        End Function
        Public Function ReadGuid() As Guid
            If IsDBNull(FieldValue) Then
                FieldValue = Guid.Empty
            End If
            Return New Guid(FieldValue.ToString)
        End Function
        Public Function ReadString() As String
            Return FieldValue.ToString
        End Function
        Public Function ReadDateTime() As DateTime
            If IsDBNull(FieldValue) Then
                FieldValue = DateTime.MinValue
            End If
            Return CType(FieldValue, DateTime)
        End Function
        Public Function ReadDecimal() As Object
            Return CDec(FieldValue)
        End Function
        Public Function ReadChar() As Object
            Return CChar(FieldValue)
        End Function
        Public Function ReadSingle() As Object
            Return CType(FieldValue, Single)
        End Function
        Public Function ReadDouble() As Object
            Return CDbl(FieldValue)
        End Function
#End Region
    End Class
End Namespace