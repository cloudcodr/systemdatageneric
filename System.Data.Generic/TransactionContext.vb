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
''' Provides context of any transaction.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class TransactionContext
    Implements IDisposable

#Region "Member variables"
    Private _innerTransaction As Common.DbTransaction
    Private _name As String
    Private _disposedValue As Boolean ' To detect redundant calls
#End Region

#Region "Constructor"
    ''' <summary>
    ''' Returns a new instnace of the TransactionContext class.
    ''' </summary>
    ''' <param name="transaction">Inner transaction of the context.</param>
    ''' <remarks></remarks>
    Protected Friend Sub New(transaction As Common.DbTransaction)
        Me.New(transaction, String.Empty)
    End Sub

    ''' <summary>
    ''' Returns a new instnace of the TransactionContext class.
    ''' </summary>
    ''' <param name="transaction">Inner transaction of the context.</param>
    ''' <param name="name">Name of the transaction.</param>
    ''' <remarks></remarks>
    Protected Friend Sub New(transaction As Common.DbTransaction, name As String)
        _innerTransaction = transaction
        _name = name
    End Sub
#End Region

#Region "Public properties"
    ''' <summary>
    ''' Gets the name of the transaction.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Name As String
        Get
            Return _name
        End Get
    End Property

    ''' <summary>
    ''' Holds the inner transaction.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Friend ReadOnly Property InnerTransaction As System.Data.Common.DbTransaction
        Get
            Return _innerTransaction
        End Get
    End Property
#End Region

#Region "Public methods"
    ''' <summary>
    ''' Commits the transction.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Commit()
        Try
            _innerTransaction.Commit()
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Sub

    ''' <summary>
    ''' Rolls back the transaction.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RollBack()
        Try
            _innerTransaction.Rollback()
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Sub
#End Region

#Region "IDisposable Support"
    ' IDisposable
    Protected Sub Dispose(disposing As Boolean)
        If Not Me._disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                _innerTransaction.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me._disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
