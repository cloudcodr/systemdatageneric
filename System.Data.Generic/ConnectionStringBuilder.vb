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

Imports System.Net

''' <summary>
''' Builds a connection to a SQL server.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class ConnectionStringBuilder
#Region "Member Variables"
    Private _server As String
    Private _database As String
#End Region

#Region "Constructor"
    ''' <summary>
    ''' Returns a new instance of the ConnectionStringBuilder class.
    ''' </summary>
    ''' <param name="server">Server address of the connection.</param>
    ''' <param name="database">Database of the connection.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal server As String, ByVal database As String)
        _server = server
        _database = database
    End Sub

    ''' <summary>
    ''' Returns a new instance of the ConnectionStringBuilder class, specifying credentials.
    ''' </summary>
    ''' <param name="server">Server address of the connection.</param>
    ''' <param name="database">Database of the connection.</param>
    ''' <param name="credentials">Username and password for the connection.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal server As String, ByVal database As String, ByVal credentials As NetworkCredential)
        _server = server
        _database = database

        Me.Credentials = credentials
    End Sub
#End Region

#Region "Shared Methods"
    ''' <summary>
    ''' Returns a connectionstring for a SQL Express.
    ''' </summary>
    ''' <param name="database">Name of database.</param>
    ''' <returns>A new ConnectionStringBuilder class for SQL Express.</returns>
    ''' <remarks></remarks>
    Public Shared Function GetSqlExpress(ByVal database As String) As ConnectionStringBuilder
        Return New ConnectionStringBuilder(".\SQLExpress", database)
    End Function
#End Region

#Region "Public properties"
    ''' <summary>
    ''' Gets the value of the server.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Server As String
        Get
            Return _server
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets the value of the database.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Database As String
        Get
            Return _database
        End Get
        Set(value As String)
            _database = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets if the connection use trusted connection capabilities.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TrustedConnection As Boolean = True

    ''' <summary>
    ''' Gets or sets if the connection use integrated security capabilities.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IntegratedSecurity As Boolean = False

    ''' <summary>
    ''' Gets or sets if the connection enables multiple active resultsets.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MultipleActiveResultSets As Boolean = False

    ''' <summary>
    ''' Gets or sets the username and password of the connection.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Credentials As NetworkCredential
#End Region

#Region "Public Methods"
    ''' <summary>
    ''' Returns the connection string.
    ''' </summary>
    ''' <returns>String. Connection string for the SQL server.</returns>
    ''' <remarks></remarks>
    Public Overrides Function ToString() As String
        Dim returnString As New Text.StringBuilder
        returnString.Append("Server=" & Me.Server)
        returnString.Append(";Database=" & Me.Database)
        returnString.Append(";Trusted_Connection=" & IIf(Me.TrustedConnection, "Yes", "No"))
        If MultipleActiveResultSets Then
            returnString.Append(";MultipleActiveResultSets=True")
        End If
        If IntegratedSecurity Then
            If TrustedConnection Then
                Throw New ArgumentException("Cannot use trusted connection with integrated security.")
            End If
            returnString.Append(";Integrated Security=True")
        End If
        If Me.Credentials IsNot Nothing Then
            If TrustedConnection Or IntegratedSecurity Then
                Throw New ArgumentException("Cannot use credentials with integrated security or trusted connection.")
            End If
            returnString.Append(String.Format(";User Id={0};Password={1}", Me.Credentials.UserName, Me.Credentials.Password))
        End If

        Return returnString.ToString
    End Function
#End Region
End Class
