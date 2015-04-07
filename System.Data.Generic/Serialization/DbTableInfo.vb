' Copyright (c) 2013-15 Arumsoft
' All rights reserved.
'
' Created:  07/04/2015.         Author:   Cloudcoder
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
    ''' Obtains information about the tables of a database.
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class DbTableInfo
#Region "Public properties"
        ''' <summary>
        ''' Name of the data table.
        ''' </summary>
        ''' <remarks></remarks>
        Public TableName As String
        ''' <summary>
        ''' Index of the key column.
        ''' </summary>
        ''' <remarks></remarks>
        Public KeyColumnIndex As Integer
#End Region

#Region "Public methods"
        ''' <summary>
        ''' Returns a string representation of the DbTableInfo class.
        ''' </summary>
        ''' <returns>String representation.</returns>
        ''' <remarks></remarks>
        Public Overrides Function ToString() As String
            Return Me.TableName + ", key column " + KeyColumnIndex.ToString
        End Function
#End Region
    End Class
End Namespace
