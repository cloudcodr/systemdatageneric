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

Imports System.Xml

''' <summary>
''' Not implemented
''' </summary>
''' <remarks></remarks>
Friend Interface IDataSource
    Inherits IDisposable

    Function ExecuteScalar(Of T)(ByVal SQL As String, ByVal defaultValue As T) As T
    Function ExecuteScalar(Of T)(ByVal SQL As String) As T
    Function ExecuteScalar(ByVal SQL As String) As Object

    Sub ExecuteNoReturn(ByVal SQL As String)

    Function ExecuteDataSet(ByVal SQL As String) As DataSet

    Function ExecuteXmlReader(ByVal SQL As String) As XmlReader

    Function ExecuteReader(ByVal SQL As String) As IDataReader

    Function ExecuteDictionary(ByVal SQL As String) As Dictionary(Of String, Object)

    Function ExecuteDataTable(ByVal SQL As String) As DataTable

    Function ExecuteObjects(Of T)(ByVal SQL As String) As List(Of T)

    Function ExecuteObject(Of T)(ByVal SQL As String) As T
End Interface
