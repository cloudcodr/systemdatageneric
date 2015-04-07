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
''' Helper methods for Data manipulation.
''' </summary>
''' <remarks></remarks>
Friend Module HelperTools
#Region "Shared Methods"
    ''' <summary>
    ''' Cast object into Type of T. If cast cannot be obtained, <paramref name="defaultValue">defaultValue</paramref> parameters is returned.
    ''' </summary>
    ''' <typeparam name="T">Type of object to cast to.</typeparam>
    ''' <param name="o">Object to cast.</param>
    ''' <param name="defaultValue">Default value to return if failure.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function TryCastDefault(Of T)(o As Object, ByVal defaultValue As Object) As T
        Try
            Dim returnValue As Object = CType(o, T)
            Return returnValue
        Catch ex As Exception
            Return defaultValue
        End Try
    End Function

    ''' <summary>
    ''' Cast object into Type of T. If cast cannot be obtained, the passed object <paramref name="o">o</paramref> is returned.
    ''' </summary>
    ''' <typeparam name="T">Type of object to cast to.</typeparam>
    ''' <param name="o">Object to cast.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function TryCastDefault(Of T)(o As Object) As T
        Return TryCastDefault(Of T)(o, o)
    End Function

    ''' <summary>
    ''' Escapes standard SQL characters.
    ''' </summary>
    ''' <param name="sql">SQL statement to escape.</param>
    ''' <returns></returns>
    ''' <remarks>Replaces ' with ''.</remarks>
    Friend Function EscapeSql(sql As String) As String
        sql = sql.Replace("'", "''")
        Return sql
    End Function

    ''' <summary>
    ''' Formats a SQL string with the <paramref name="parameters">parameters</paramref>, assigned the right usage of quotas.
    ''' </summary>
    ''' <param name="sql">SQL statement to format. Use String formatting {0}, {1} etc.</param>
    ''' <param name="parameters">Array of parameters.</param>
    ''' <returns>Formatted SQL string.</returns>
    ''' <remarks>The PrepareSqlExtended method inspects each parameters by datatype and assigns correct T-SQL quota usage. 
    ''' This might impact performance or expected output. You can escape the proper formatting using the \ escape char first and last. 
    ''' The method will add NULL to the T-SQL if the value is nothing, empty or minvalue.
    ''' </remarks>
    Friend Function PrepareSqlExtended(ByVal sql As String, ByVal ParamArray parameters() As Object) As String
        Dim stringParameters As List(Of String) = New List(Of String)

        ' loop each parameter and assign qupta usage
        For Each param As Object In parameters
            ' default empty parameters to empty string (which will be NULL)
            If param Is Nothing Then param = ""

            If param.GetType Is GetType(String) Then
                If String.IsNullOrEmpty(param) AndAlso DataSource.SerializationSettings.ThreatMinValuesAsNull Then
                    stringParameters.Add("NULL")
                Else
                    Dim paramValue As String = param.ToString
                    If paramValue.StartsWith("\") AndAlso paramValue.EndsWith("\") Then
                        ' string is escaped using format \something\
                        ' do not add ' ' around. example is using ORDER BY {0} in string.format. No use of ORDER BY 'DESC'
                        stringParameters.Add(EscapeSql(paramValue.Substring(1, paramValue.Length - 2)))
                    Else
                        stringParameters.Add("'" + EscapeSql(paramValue) + "'")
                    End If
                End If
            ElseIf param.GetType Is GetType(Guid) Then
                If DirectCast(param, Guid) = Guid.Empty AndAlso DataSource.SerializationSettings.ThreatMinValuesAsNull Then
                    stringParameters.Add("NULL")
                Else
                    stringParameters.Add("'" + param.ToString + "'")
                End If
            ElseIf param.GetType Is GetType(Boolean) Then
                stringParameters.Add(CInt(param))
            ElseIf param.GetType Is GetType(Date) Then
                If DirectCast(param, Date) = Date.MinValue AndAlso DataSource.SerializationSettings.ThreatMinValuesAsNull Then
                    stringParameters.Add("NULL")
                Else
                    stringParameters.Add("'" + DirectCast(param, Date).ToString("yyyy-MM-dd") + "'")
                End If
            ElseIf param.GetType Is GetType(DateTime) Then
                If DirectCast(param, DateTime) = DateTime.MinValue AndAlso DataSource.SerializationSettings.ThreatMinValuesAsNull Then
                    stringParameters.Add("NULL")
                Else
                    stringParameters.Add("'" + DirectCast(param, DateTime).ToString("yyyy-MM-dd HH:mm:ss") + "'")
                End If
            Else
                ' format as integer, long, float, decimal
                stringParameters.Add(param)
            End If
        Next

        Return String.Format(sql, stringParameters.ToArray)
        'Return HelperTools.PrepareSql(sql, stringParameters.ToArray)
    End Function
#End Region
End Module
