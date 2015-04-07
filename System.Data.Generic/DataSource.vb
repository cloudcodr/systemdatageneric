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
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Collections
Imports System.Collections.Generic
Imports System.Collections.ObjectModel
Imports System.Collections.Specialized
Imports System.Data.Generic.Serialization.Formatters

''' <summary>
''' Enables easy data access to SQL server.
''' </summary>
''' <remarks>
''' Enables global data access by editing web.config, app.config or the Azure configuration file, and adding GlobalDataSource to the appSettings configuration elements.
''' </remarks>
''' <example>
''' <code>
''' <appSettings><add key="GlobalDataSource" value="Server=.\SQLExpress;Database=somedatabase;" /></appSettings>
''' </code>
''' Usage example 1 (get instance into variable):
''' <code>
''' DataSource source = DataSource.Current();
''' </code>
''' Usage example 2 (use static method):
''' <code>
''' DataSource.Current.ExecuteNoReturn("select * from table");
''' </code>
''' Usage example 3 (create an instance):
''' <code>
''' DataSource source = new DataSource("Server=.\SQLExpress;Database=somedatabase;");
''' source.ExecuteNoReturn("select * from table");
''' </code>
''' </example>
Public NotInheritable Class DataSource
    Implements IDisposable

#Region "Member variables"
    Private _connectionString As String

    Private Shared _instance As DataSource = Nothing
    Private Shared _serializationSettings As SerializationSettings = Nothing
#End Region

#Region "Constructor"
    ''' <summary>
    ''' Returns a new instance of the DataSource class, using a connectionstring builder.
    ''' </summary>
    ''' <param name="connection">ConnectionStringBuilder class for the connection.</param>
    ''' <param name="delayOpening">Boolean. Indicates if the data connection should be opened upon intialization of the DataSource class.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal connection As ConnectionStringBuilder, ByVal delayOpening As Boolean)
        Me.New(connection.ToString, delayOpening)
    End Sub

    ''' <summary>
    ''' Returns a new instance of the DataSource class.
    ''' </summary>
    ''' <param name="connectionString">Connection string to open.</param>
    ''' <param name="delayOpening">Boolean. Indicates if the data connection should be opened upon intialization of the DataSource class.</param>
    ''' <remarks>Use <see cref="DataSource.Open">Open</see> method to open later.</remarks>
    Public Sub New(ByVal connectionString As String, ByVal delayOpening As Boolean)
        _connectionString = connectionString

        'If Not delayOpening Then
        '    Open()
        'End If
    End Sub

    ''' <summary>
    ''' Returns a new instance of the DataSource class, and opens the connection immediately.
    ''' </summary>
    ''' <param name="connectionString">Connection string to open.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal connectionString As String)
        Me.New(connectionString, False)
    End Sub
#End Region

#Region "Public properties"
    ''' <summary>
    ''' Returns a new instance of the DataSource class using the connection string, specific in the app.config.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Add GlobalDataSource to the app.config connectionstring elements.
    ''' <example>
    ''' <code>
    ''' <ConnectionStrings>
    '''        <add name="GlobalDataSource" connectionString="Server=.\SQLExpress;Database=somedatabase;" />
    ''' </ConnectionStrings>
    ''' </code>
    ''' </example>
    ''' </remarks>
    Public Shared ReadOnly Property Current As DataSource
        Get
            If _instance Is Nothing Then
                Dim connectionString As String = ConfigurationHelper.GetSetting("GlobalDataSource")

                If String.IsNullOrEmpty(connectionString) Then
                    Throw New InvalidConstraintException("[GlobalDataSource] connectionstring expected in app.config.")
                End If

                _instance = New DataSource(connectionString)
            End If

            Return _instance
        End Get
    End Property

    ''' <summary>
    ''' Formats a SQL statement based on the data type of the <paramref name="parameters">parameters</paramref> parameters.
    ''' </summary>
    ''' <param name="sql">SQL statement.</param>
    ''' <param name="parameters">Parameters.</param>
    ''' <returns>Formattet SQL statement.</returns>
    ''' <remarks>Behaviour of the SQL statement formatting is based on the <see cref="SerializationSettings.ThreatMinValuesAsNull">ThreatMinValuesAsNull</see> property on <see cref="DataSource.SerializationSettings">SerializationSettings</see> property.</remarks>
    Public Shared Function PrepareSql(ByVal sql As String, ByVal ParamArray parameters() As Object) As String
        Return HelperTools.PrepareSqlExtended(sql, parameters)
    End Function

    ''' <summary>
    ''' Escapes standard SQL characters.
    ''' </summary>
    ''' <param name="sql">SQL statement to escape.</param>
    ''' <returns>Escaped SQL string.</returns>
    ''' <remarks>Replaces ' with ''.</remarks>
    Public Shared Function EscapeSql(sql As String) As String
        Return HelperTools.EscapeSql(sql)
    End Function

    ''' <summary>
    ''' Escapes standard SQL characters. Can be used as String.Format.
    ''' </summary>
    ''' <param name="sql">SQL statement to escape.</param>
    ''' <param name="args">Arguments to format.</param>
    ''' <returns>Escaped SQL string.</returns>
    ''' <remarks>Replaces ' with ''.</remarks>
    <Obsolete("EscapeSql is obsolete and embedded with PrepareSql.")>
    Public Shared Function EscapeSql(sql As String, ParamArray args() As Object) As String
        Dim arg2() As Object = New Object(args.Length - 1) {}
        For i As Integer = 0 To args.Length - 1 ' TODO: is this length correct?
            arg2(i) = args(i).ToString.Replace("'", "''")
        Next

        sql = String.Format(sql, arg2)
        Return sql
    End Function

    ''' <summary>
    ''' Returns the connection state of the data connection.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Obsolete("You can remove Open. Does not be used anymore.")>
    Public ReadOnly Property ConnectionState As ConnectionState
        Get
            Return Data.ConnectionState.Connecting
        End Get
    End Property

    ''' <summary>
    ''' Returns the connection string of the current data connection.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ConnectionString As String
        Get
            Return _connectionString
        End Get
    End Property

    ''' <summary>
    ''' Returns the settings for serialization.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Settings class for the supporting serializer.</returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property SerializationSettings As SerializationSettings
        Get
            If _serializationSettings Is Nothing Then
                _serializationSettings = New SerializationSettings()
            End If

            Return _serializationSettings
        End Get
    End Property
#End Region

#Region "Open/Close methods"
    ''' <summary>
    ''' Open the data connection.
    ''' </summary>
    ''' <remarks></remarks> 
    <Obsolete("You can remove Open. Does not be used anymore.")>
    Public Sub Open()
        'Try
        '    _dataConnection = New SqlConnection(Me.ConnectionString)
        '    _dataConnection.Open()

        'Catch ex As SqlException
        '    Throw ex
        'Catch ex As Exception
        '    Throw ex
        'Finally

        'End Try
    End Sub

    ''' <summary>
    ''' Close the data connection.
    ''' </summary>
    ''' <remarks></remarks>
    <Obsolete("You can remove Open. Does not be used anymore.")>
    Public Sub Close()
        'If _dataConnection IsNot Nothing Then
        '    _dataConnection.Close()
        'End If
    End Sub
#End Region

#Region "Connection Pool methods"
    ''' <summary>
    ''' Empties the connection pool associated with the current connection.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClearPool()
        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            SqlConnection.ClearPool(dataConnection)
        End Using
    End Sub

    ''' <summary>
    ''' Empties all connection pool(s).
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub ClearPools()
        SqlConnection.ClearAllPools()
    End Sub
#End Region

#Region "ExecuteScalar"
    ''' <summary>
    ''' Executes the query, and returns the first column of the first row in the result set returned by the query. Supports transactions.
    ''' </summary>
    ''' <typeparam name="T">Type to return.</typeparam>
    ''' <param name="SQL">Transact-SQL statement to execute against the connection.</param>
    ''' <param name="transaction">Context of the transaction.</param>
    ''' <param name="defaultValue">Default value to return in case of no return.</param>
    ''' <returns>The first column of the first row in the result set, or a null reference (Nothing in Visual Basic) if the result set is empty. Returns a maximum of 2033 characters.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentNullException">In case the SQL is null.</exception>
    ''' <exception cref="SqlException">In case the underlying connection cause an exception.</exception>
    Public Function ExecuteScalar(Of T)(ByVal SQL As String, transaction As TransactionContext, ByVal defaultValue As T) As T
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        ' set defaults
        Dim dataValue As Object = Nothing

        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            dataConnection.Open()

            ' build command
            Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection, transaction.InnerTransaction)
            dataCommand.CommandType = CommandType.Text

            Try
                dataValue = dataCommand.ExecuteScalar()

            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                dataCommand.Dispose()
                ' close and return connection to pool
                If dataConnection.State = Data.ConnectionState.Open Then
                    dataConnection.Close()
                End If
            End Try
        End Using

        If dataValue IsNot Nothing AndAlso Not IsDBNull(dataValue) Then
            'Return TryCastDefault(Of T)(dataValue, defaultValue)
            'Return TryCast(dataValue, T)
            Return DirectCast(dataValue, T)
        Else
            Return defaultValue
        End If
    End Function

    ''' <summary>
    ''' Executes the query, and returns the first column of the first row in the result set returned by the query.
    ''' </summary>
    ''' <typeparam name="T">Type to return.</typeparam>
    ''' <param name="SQL">Transact-SQL statement to execute against the connection.</param>
    ''' <param name="defaultValue">Default value to return in case of no return.</param>
    ''' <returns>The first column of the first row in the result set, or a null reference (Nothing in Visual Basic) if the result set is empty. Returns a maximum of 2033 characters.</returns>
    ''' <remarks></remarks>
    Public Function ExecuteScalar(Of T)(ByVal SQL As String, ByVal defaultValue As T) As T
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        ' set defaults
        Dim dataValue As Object = Nothing

        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            dataConnection.Open()

            ' build command
            Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection)
            dataCommand.CommandType = CommandType.Text

            Try
                dataValue = dataCommand.ExecuteScalar()

            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                dataCommand.Dispose()
                ' close and return connection to pool
                If dataConnection.State = Data.ConnectionState.Open Then
                    dataConnection.Close()
                End If
            End Try
        End Using

        If dataValue IsNot Nothing AndAlso Not IsDBNull(dataValue) Then
            'Return TryCastDefault(Of T)(dataValue, defaultValue)
            'Return TryCast(dataValue, T)
            Return DirectCast(dataValue, T)
        Else
            Return defaultValue
        End If
    End Function

    ''' <summary>
    ''' Executes the query, and returns the first column of the first row in the result set returned by the query.
    ''' </summary>
    ''' <typeparam name="T">Type to return.</typeparam>
    ''' <param name="SQL">Transact-SQL statement to execute against the connection.</param>
    ''' <returns>The first column of the first row in the result set, or a null reference (Nothing in Visual Basic) if the result set is empty. Returns a maximum of 2033 characters.</returns>
    ''' <remarks></remarks>
    Public Function ExecuteScalar(Of T)(ByVal SQL As String) As T
        Return ExecuteScalar(Of T)(SQL, Nothing)
    End Function

    ''' <summary>
    ''' Executes the query, and returns the first column of the first row in the result set returned by the query.
    ''' </summary>
    ''' <param name="SQL">Transact-SQL statement to execute against the connection.</param>
    ''' <returns>The first column of the first row in the result set, or a null reference (Nothing in Visual Basic) if the result set is empty. Returns a maximum of 2033 characters.</returns>
    ''' <remarks></remarks>
    Public Function ExecuteScalar(ByVal SQL As String) As Object
        Return ExecuteScalar(Of Object)(SQL, Nothing)
    End Function
#End Region

#Region "ExecuteNoReturn"
    ''' <summary>
    ''' Executes the Query against the datasource without returning any value.
    ''' </summary>
    ''' <param name="SQL">SQL statement to execute.</param>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentNullException">In case the SQL is null.</exception>
    ''' <exception cref="SqlException">In case the underlying connection cause an exception.</exception>
    Public Sub ExecuteNoReturn(ByVal SQL As String)
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            dataConnection.Open()

            ' build command
            Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection)
            dataCommand.CommandType = CommandType.Text

            Try
                dataCommand.ExecuteNonQuery()

            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                dataCommand.Dispose()
                ' close connection and return to pool
                If dataConnection.State = Data.ConnectionState.Open Then
                    dataConnection.Close()
                End If
            End Try
        End Using
    End Sub

    ''' <summary>
    ''' Executes the Query against the datasource without returning any value. Supports transactions.
    ''' </summary>
    ''' <param name="SQL">SQL statement to execute.</param>
    ''' <param name="transaction">Context of the transaction.</param>
    ''' <remarks>The connection will remain open, until the TransactionContext is disposed or RollBack/Commit is called.</remarks>
    Public Sub ExecuteNoReturn(SQL As String, transaction As TransactionContext)
        If transaction Is Nothing Then
            Throw New NullReferenceException("Cannot specify a transaction of NULL.")
        End If
        If transaction.InnerTransaction Is Nothing Then
            Throw New NullReferenceException("Cannot specify a InnerTransaction of NULL.")
        End If
        If transaction.InnerTransaction.Connection Is Nothing Then
            Throw New NullReferenceException("Cannot specify a Connection of NULL.")
        End If

        ' build command
        Dim dataCommand As SqlCommand = New SqlCommand(SQL, transaction.InnerTransaction.Connection, transaction.InnerTransaction)
        dataCommand.CommandType = CommandType.Text

        Try
            dataCommand.ExecuteNonQuery()

        Catch ex As SqlException
            Throw ex
        Catch ex As Exception
            Throw ex
        Finally
            ' leave the connection open
            dataCommand.Dispose()
        End Try
    End Sub
#End Region

#Region "ExecuteDataSet"
    ''' <summary>
    ''' Executes the SQL query and returns a DataSet.
    ''' </summary>
    ''' <param name="SQL">SQL query to execute.</param>
    ''' <returns>DataSet of the data.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentNullException">In case the SQL is null.</exception>
    ''' <exception cref="SqlException">In case the underlying connection cause an exception.</exception>
    Public Function ExecuteDataSet(ByVal SQL As String) As DataSet
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        Dim dataReturnSet As DataSet = New DataSet

        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            dataConnection.Open()

            ' build command
            Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection)
            dataCommand.CommandType = CommandType.Text

            Try
                Dim dataAdapter As SqlDataAdapter = New SqlDataAdapter(dataCommand)
                dataAdapter.Fill(dataReturnSet)

            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                dataCommand.Dispose()
                ' close connection and return to pool
                If dataConnection.State = Data.ConnectionState.Open Then
                    dataConnection.Close()
                End If
            End Try
        End Using

        Return dataReturnSet
    End Function
#End Region

#Region "ExecuteXmlReader"
    ''' <summary>
    ''' Executes the Query against the datasource and returns a XML reader.
    ''' </summary>
    ''' <param name="SQL">SQL query to execute.</param>
    ''' <returns>XmlReader.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentNullException">In case the SQL is null.</exception>
    ''' <exception cref="SqlException">In case the underlying connection cause an exception.</exception>
    Public Function ExecuteXmlReader(ByVal SQL As String) As XmlReader
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        Dim dataXmlReader As XmlReader = Nothing

        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            dataConnection.Open()

            ' build command
            Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection)
            dataCommand.CommandType = CommandType.Text

            Try
                dataXmlReader = dataCommand.ExecuteXmlReader()

            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                dataCommand.Dispose()
                ' close connection and return to pool
                If dataConnection.State = Data.ConnectionState.Open Then
                    dataConnection.Close()
                End If
            End Try
        End Using

        Return dataXmlReader
    End Function
#End Region

#Region "ExecuteReader"
    ''' <summary>
    ''' Sends the SQL Query to the Connection and builds a SqlDataReader.
    ''' </summary>
    ''' <param name="SQL">SQL Query to execute.</param>
    ''' <returns>SqlDataReader.</returns>
    ''' <remarks>The ExecuteReader method keeps the connection open. Dispose of the <see cref="System.Data.Common.DbDataReader">DbDataReader</see> will close the connection.</remarks>
    ''' <exception cref="ArgumentNullException">In case the SQL is null.</exception>
    ''' <exception cref="SqlException">In case the underlying connection cause an exception.</exception>
    Public Function ExecuteReader(ByVal SQL As String) As IDataReader
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        Dim dataReader As SqlDataReader = Nothing

        Dim dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
        dataConnection.Open()

        ' build command
        Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection)
        dataCommand.CommandType = CommandType.Text

        Try
            dataReader = dataCommand.ExecuteReader()

        Catch ex As SqlException
            Throw ex
        Catch ex As Exception
            Throw ex
        Finally
            ' leave connection open for reader to operate.
        End Try

        Return dataReader
    End Function
#End Region

#Region "ExecuteDictionary"
    ''' <summary>
    ''' Executes the SQL statement and returns a dictionary of the first row. Keys in the dictionary are field names and values associated.
    ''' </summary>
    ''' <param name="SQL">SQL statement to execute.</param>
    ''' <returns>Dictionary class.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentNullException">In case the SQL is null.</exception>
    ''' <exception cref="SqlException">In case the underlying connection cause an exception.</exception>
    Public Function ExecuteDictionary(ByVal SQL As String) As Dictionary(Of String, Object)
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        Dim dataDictionary As Dictionary(Of String, Object) = New Dictionary(Of String, Object)

        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            dataConnection.Open()

            ' build command
            Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection)
            dataCommand.CommandType = CommandType.Text

            Try
                Using dataReader As SqlDataReader = dataCommand.ExecuteReader()
                    ' read first row
                    While dataReader.Read()
                        ' read each field
                        PopulateDictionary(dataDictionary, dataReader)

                        ' only return one
                        Exit While
                    End While

                    dataReader.Close()
                End Using

            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                dataCommand.Dispose()
                ' close connection and return to pool
                If dataConnection.State = Data.ConnectionState.Open Then
                    dataConnection.Close()
                End If
            End Try
        End Using

        Return dataDictionary
    End Function
#End Region

#Region "ExecuteDataTable"
    ''' <summary>
    ''' Sends the SQL Query to the Connection and returns a DataTable.
    ''' </summary>
    ''' <param name="SQL">SQL statement to execute.</param>
    ''' <returns>DataTable.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentNullException">In case the SQL is null.</exception>
    ''' <exception cref="SqlException">In case the underlying connection cause an exception.</exception>
    Public Function ExecuteDataTable(ByVal SQL As String) As DataTable
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        Dim dataReturnSet As DataSet = New DataSet

        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            dataConnection.Open()

            ' build command
            Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection)
            dataCommand.CommandType = CommandType.Text

            Try
                Dim dataAdapter As SqlDataAdapter = New SqlDataAdapter(dataCommand)
                dataAdapter.Fill(dataReturnSet)

            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                dataCommand.Dispose()
                ' close connection and return to pool
                If dataConnection.State = Data.ConnectionState.Open Then
                    dataConnection.Close()
                End If
            End Try
        End Using

        Return dataReturnSet.Tables(0)
    End Function
#End Region

#Region "ExecuteCallBack"
    ''' <summary>
    ''' Delegated method for call back event methods. See also <see cref="ExecuteCallBack">ExecuteCallBack</see>.
    ''' </summary>
    ''' <param name="item">Current data item as a Dictionary(Of String, Object).</param>
    ''' <param name="rownum">Row number of the item.</param>
    ''' <param name="cancel">Gets or sets a value indicating if the read should continue.</param>
    ''' <remarks></remarks>
    Public Delegate Sub ItemCallBackDelegate(item As Dictionary(Of String, Object), rownum As Integer, ByRef cancel As Boolean)

    ''' <summary>
    ''' Executes the query, and raises a callback of <see cref="ItemCallBackDelegate">ItemCallBackDelegate</see> for each item it iterates.
    ''' </summary>
    ''' <param name="SQL">SQL statement to execute.</param>
    ''' <param name="callback"><see cref="ItemCallBackDelegate">ItemCallBackDelegate</see> method.</param>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentNullException">In case the SQL is null.</exception>
    ''' <exception cref="SqlException">In case the underlying connection cause an exception.</exception>
    Public Sub ExecuteCallBack(ByVal SQL As String, ByVal callback As ItemCallBackDelegate)
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            dataConnection.Open()

            ' build command
            Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection)
            dataCommand.CommandType = CommandType.Text

            Try
                Using dataReader As SqlDataReader = dataCommand.ExecuteReader(CommandBehavior.CloseConnection)
                    Dim cancelRead As Boolean = False
                    Dim rowCounter As Integer = 0

                    ' read first row
                    While dataReader.Read()
                        rowCounter += 1

                        ' read current field into dictionary
                        Dim dataDictionary As Dictionary(Of String, Object) = New Dictionary(Of String, Object)
                        PopulateDictionary(dataDictionary, dataReader)

                        ' raise the call back
                        callback(dataDictionary, rowCounter, cancelRead)
                        ' check if we should continue
                        If cancelRead Then
                            Exit While
                        End If
                    End While
                    dataReader.Close()
                End Using

            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                dataCommand.Dispose()
                ' close and return connection to pool
                If dataConnection.State = Data.ConnectionState.Open Then
                    dataConnection.Close()
                End If
            End Try
        End Using
    End Sub
#End Region

#Region "ExecuteList"
    ''' <summary>
    ''' Executes the query, and returns a collection of the first column.
    ''' </summary>
    ''' <typeparam name="T">Type of the first column.</typeparam>
    ''' <param name="SQL">SQL statement to execute.</param>
    ''' <returns>List of first column</returns>
    ''' <remarks>Specifying a SQL statement with multiple columns will not change the output. Only the first column will be returned.</remarks>
    ''' <exception cref="ArgumentNullException">In case the SQL is null.</exception>
    ''' <exception cref="SqlException">In case the underlying connection cause an exception.</exception>
    Public Function ExecuteList(Of T)(ByVal SQL As String) As List(Of T)
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        Dim dataReturnList As List(Of T) = New List(Of T)

        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            dataConnection.Open()

            ' build command
            Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection)
            dataCommand.CommandType = CommandType.Text

            Try
                Using dataReader As SqlDataReader = dataCommand.ExecuteReader(CommandBehavior.CloseConnection)
                    ' read first row
                    While dataReader.Read()
                        dataReturnList.Add(CType(dataReader(0), T))
                    End While
                    dataReader.Close()
                End Using

            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                dataCommand.Dispose()
                ' close and return connection to pool
                If dataConnection.State = Data.ConnectionState.Open Then
                    dataConnection.Close()
                End If
            End Try
        End Using

        Return dataReturnList
    End Function
#End Region

#Region "ExecuteNObjects"
    ''' <summary>
    ''' Executes the query, and returns a collection of object by the type T; hieratically compounded by the provided SQL JOIN and realted classes.
    ''' </summary>
    ''' <typeparam name="T">Type to return in list. This must be a managed object.</typeparam>
    ''' <param name="SQL">SQL statement to execute. SQL statement must have an 'ORDER BY' statement. Should be a JOIN statement, otherwise please use ExecuteObjects for performance.</param>
    ''' <returns>Collections of object, type T - populated based on class references and provided tables.</returns>
    ''' <remarks></remarks>
    ''' <exception cref="ArgumentNullException">In case the SQL is null.</exception>
    ''' <exception cref="SqlException">In case the underlying connection cause an exception.</exception>
    Public Function ExecuteNObjects(Of T As Class)(SQL As String) As List(Of T)
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        ' check for order by.
        If Not SQL.Contains("order by", StringComparison.OrdinalIgnoreCase) Then
            Throw New ArgumentException("SQL must have define ORDER BY to map objects in the right order.")
        End If

        ' setup return list
        Dim dataReturnList As List(Of T) = New List(Of T)

        ' initialize the data connection an dopen
        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            dataConnection.Open()

            ' build command and apply command text (SQL)
            Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection)
            dataCommand.CommandType = CommandType.Text

            Try
                ' setup formatter of type T
                Dim sqlFormatter As New Serialization.Formatters.DataSourceFormatter(Of T)

                ' important - open the connection with close and keyinfo!!
                Using dataReader As SqlDataReader = dataCommand.ExecuteReader(CommandBehavior.CloseConnection Or CommandBehavior.KeyInfo)
                    ' get schema, and pass to the deserializenested
                    Dim dataSchema As Serialization.DbSchema = Serialization.DbSchema.GetSchema(dataReader)

                    While dataReader.Read()
                        ' deserialize the object and reader as a nested hieraticy.
                        Dim dataReturn As T = sqlFormatter.DeserializeNested(dataReader, dataSchema)

                        dataReturnList.Add(dataReturn)
                    End While
                    dataReader.Close()
                End Using
            Catch ex As System.Runtime.Serialization.SerializationException
                Throw ex
            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                dataCommand.Dispose()
                ' close and return connection to pool
                If dataConnection.State = Data.ConnectionState.Open Then
                    dataConnection.Close()
                End If
            End Try
        End Using

        Return dataReturnList
    End Function
#End Region

#Region "ExecuteObjects"
    ''' <summary>
    ''' Executes the query, and returns a collection of object by the type T.
    ''' </summary>
    ''' <typeparam name="T">Type to return in list. This must be a managed object.</typeparam>
    ''' <param name="SQL">SQL statement to execute.</param>
    ''' <returns>Collections of object, type T.</returns>
    ''' <remarks>ExecuteObjects does not support standardized types, as reflection is reversed. This methods tries to match properties of the T type to the data source. Use <see cref="DataSource.ExecuteList">ExecuteList</see> or <see cref="DataSource.ExecuteDictionary">ExecuteDictionary</see>.</remarks>
    ''' <exception cref="ArgumentNullException">In case the SQL is null.</exception>
    ''' <exception cref="SqlException">In case the underlying connection cause an exception.</exception>
    Public Function ExecuteObjects(Of T)(ByVal SQL As String) As List(Of T)
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        ' setup return value
        Dim dataReturnList As List(Of T) = New List(Of T)

        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            dataConnection.Open()

            ' build command
            Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection)
            dataCommand.CommandType = CommandType.Text

            Try
                Dim sqlFormatter As New Serialization.Formatters.DataSourceFormatter(Of T) ' GenericDataFormatter(Of T) = New GenericDataFormatter(Of T)()

                Using dataReader As SqlDataReader = dataCommand.ExecuteReader(CommandBehavior.CloseConnection)
                    ' read first row
                    While dataReader.Read()

                        Dim dataReturn As T = sqlFormatter.Deserialize(dataReader)
                        dataReturnList.Add(dataReturn)

                    End While
                    dataReader.Close()
                End Using
            Catch ex As System.Runtime.Serialization.SerializationException
                Throw ex
            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                dataCommand.Dispose()
                ' close and return connection to pool
                If dataConnection.State = Data.ConnectionState.Open Then
                    dataConnection.Close()
                End If
            End Try
        End Using

        Return dataReturnList
    End Function
#End Region

#Region "ExecuteObject"
    ''' <summary>
    ''' Executes the query, and returns a object by the type T.
    ''' </summary>
    ''' <typeparam name="T">Type of object to return.</typeparam>
    ''' <param name="SQL">SQL statement to execute.</param>
    ''' <returns>Object of type T.</returns>
    ''' <exception cref="ArgumentNullException">In case the SQL is null.</exception>
    ''' <exception cref="SqlException">In case the underlying connection cause an exception.</exception>
    ''' <remarks></remarks>
    Public Function ExecuteObject(Of T)(ByVal SQL As String) As T
        ' validate SQL input
        If String.IsNullOrEmpty(SQL) Then
            Throw New ArgumentNullException("SQL must be explicitly stated.")
        End If

        Dim dataReturn As T = Nothing

        Using dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
            dataConnection.Open()

            ' build command
            Dim dataCommand As SqlCommand = New SqlCommand(SQL, dataConnection)
            dataCommand.CommandType = CommandType.Text

            Try
                ' setup serializer
                Dim dataFormatter As New DataSourceFormatter(Of T) ' GenericDataFormatter(Of T) = New GenericDataFormatter(Of T)()

                Using dataReader As SqlDataReader = dataCommand.ExecuteReader(CommandBehavior.CloseConnection)
                    ' read first row
                    While dataReader.Read()

                        dataReturn = dataFormatter.Deserialize(dataReader)
                        ' only return one
                        Exit While
                    End While
                    dataReader.Close()
                End Using
            Catch ex As SqlException
                Throw ex
            Catch ex As Exception
                Throw ex
            Finally
                dataCommand.Dispose()
                ' close and return connection to pool
                If dataConnection.State = Data.ConnectionState.Open Then
                    dataConnection.Close()
                End If
            End Try
        End Using

        Return dataReturn
    End Function
#End Region

#Region "Transaction"
    ''' <summary>
    ''' Commits a transaction from the specified context.
    ''' </summary>
    ''' <param name="context">Context of the transaction.</param>
    ''' <remarks></remarks>
    <Obsolete("Please use Commit() method on the TransactionContext")>
    Public Sub CommitTransaction(context As TransactionContext)
        Throw New NotSupportedException("Obsolete. Use methods on TransactionContext.")
    End Sub

    ''' <summary>
    ''' Rolls back a transaction from the specified context.
    ''' </summary>
    ''' <param name="context">Context of the transaction.</param>
    ''' <remarks></remarks>
    <Obsolete("Please use Rollback() method on the TransactionContext")>
    Public Sub RollbackTransaction(context As TransactionContext)
        Throw New NotSupportedException("Obsolete. Use methods on TransactionContext.")
    End Sub

    ''' <summary>
    ''' Create a new transaction context.
    ''' </summary>
    ''' <returns>Context of the transaction to passe to subsequent requests.</returns>
    ''' <remarks></remarks>
    Public Function CreateTransaction() As TransactionContext
        Return Me.CreateTransaction(String.Empty)
    End Function

    ''' <summary>
    ''' Create a new transaction context and leaves the connection open.
    ''' </summary>
    ''' <param name="name">Name of the transaction</param>
    ''' <returns>Context of the transaction to passe to subsequent requests.</returns>
    ''' <remarks></remarks>
    Public Function CreateTransaction(name As String) As TransactionContext
        Dim dataConnection As SqlConnection = New SqlConnection(Me.ConnectionString)
        dataConnection.Open()

        Dim dataTransaction As SqlTransaction = dataConnection.BeginTransaction()

        Return New TransactionContext(dataTransaction, name)
    End Function
#End Region

#Region "Private Methods"
    ''' <summary>
    ''' Populates a dicrionary from a datareader.
    ''' </summary>
    ''' <param name="dataDictionary">Dictionary to populate.</param>
    ''' <param name="dataReader">Datareader to ready from.</param>
    ''' <remarks></remarks>
    Private Sub PopulateDictionary(dataDictionary As Dictionary(Of String, Object), dataReader As SqlDataReader)
        For i As Integer = 0 To dataReader.FieldCount - 1
            Dim fieldName As String = dataReader.GetName(i)
            Dim fieldObject As Object = dataReader(dataReader.GetName(i))

            dataDictionary.Add(fieldName, fieldObject)
        Next
    End Sub
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ''' <summary>
    ''' Releases all resources used by the Component.
    ''' </summary>
    ''' <param name="disposing"></param>
    ''' <remarks></remarks>
    Protected Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                ' Me.Close()
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

    ''' <summary>
    ''' Releases all resources used by the Component.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
