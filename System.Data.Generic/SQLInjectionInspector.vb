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
''' Injection inspector class to enable safe SQL statements.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class SQLInjectionInspector
#Region "Member Variables"
    ''' <summary>
    ''' Returns an array of unsafe injections.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Friend Shared Injections As String() = New String() {"DROP DATABASE", "DROP TABLE", "DROP COLUMN", "ALTER TABLE", "ALTER COLUMN"}
#End Region

#Region "Public Methods"
    ''' <summary>
    ''' Validate the SQL string and thrown an exception if harmfull injections are found.
    ''' </summary>
    ''' <param name="SQL">SQL statement to validate.</param>
    ''' <returns>Safe SQL statement.</returns>
    ''' <remarks>This class does not ensure unsafe SQL. Please use all means to secure your SQL access.</remarks>
    Public Shared Function Validate(ByVal SQL As String) As String
        Return SQLInjectionInspector.Validate(SQL, Injections)
    End Function

    ''' <summary>
    ''' Validate the SQL string and thrown an exception if harmfull injections are found.
    ''' </summary>
    ''' <param name="SQL">SQL statement to validate.</param>
    ''' <param name="injectionArray">Array of unsafe statements.</param>
    ''' <returns>Safe SQL statement.</returns>
    ''' <remarks>This class does not ensure unsafe SQL. Please use all means to secure your SQL access.</remarks>
    Public Shared Function Validate(ByVal SQL As String, ByVal injectionArray As String()) As String
        Dim UpperSQL As String = SQL.ToUpper
        UpperSQL = UpperSQL.Replace("'", "''")

        For Each injection As String In injectionArray
            If UpperSQL.Contains(injection) Then
                Throw New InvalidConstraintException(String.Format("SQL contains harmfull injection '{0}'. Aborting.", injection))
            End If
        Next

        Return UpperSQL
    End Function
#End Region
End Class
