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

''' <summary>
''' Not implemented
''' </summary>
''' <remarks></remarks>
Friend NotInheritable Class SqlCeDataSource
    Implements IDataSource

    Public Function ExecuteDataSet(SQL As String) As DataSet Implements IDataSource.ExecuteDataSet
        Throw New NotImplementedException
    End Function

    Public Function ExecuteDataTable(SQL As String) As DataTable Implements IDataSource.ExecuteDataTable
        Throw New NotImplementedException
    End Function

    Public Function ExecuteDictionary(SQL As String) As Dictionary(Of String, Object) Implements IDataSource.ExecuteDictionary
        Throw New NotImplementedException
    End Function

    Public Sub ExecuteNoReturn(SQL As String) Implements IDataSource.ExecuteNoReturn
        Throw New NotImplementedException
    End Sub

    Public Function ExecuteObject(Of T)(SQL As String) As T Implements IDataSource.ExecuteObject
        Throw New NotImplementedException
    End Function

    Public Function ExecuteObjects(Of T)(SQL As String) As List(Of T) Implements IDataSource.ExecuteObjects
        Throw New NotImplementedException
    End Function

    Public Function ExecuteReader(SQL As String) As IDataReader Implements IDataSource.ExecuteReader
        Throw New NotImplementedException
    End Function

    Public Function ExecuteScalar(SQL As String) As Object Implements IDataSource.ExecuteScalar
        Throw New NotImplementedException
    End Function

    Public Function ExecuteScalar1(Of T)(SQL As String) As T Implements IDataSource.ExecuteScalar
        Throw New NotImplementedException
    End Function

    Public Function ExecuteScalar1(Of T)(SQL As String, defaultValue As T) As T Implements IDataSource.ExecuteScalar
        Throw New NotImplementedException
    End Function

    Public Function ExecuteXmlReader(SQL As String) As Xml.XmlReader Implements IDataSource.ExecuteXmlReader
        Throw New NotImplementedException
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
