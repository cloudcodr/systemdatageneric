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

Imports System.Runtime
Imports System.Reflection
Imports System.Data.Generic.Settings

''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Friend NotInheritable Class ConfigurationHelper

#Region "Member variables"
    Private Shared _appSettings As AzureApplicationSettings
    Private Shared lock As Object = New Object()
#End Region

#Region "Public methods"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="key"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSetting(ByVal key As String) As String
        If key Is Nothing Then
            Throw New ArgumentNullException("Key")
        End If

        If key.Length = 0 Then
            Throw New ArgumentException("Key is not specified")
        End If

        Return AppSettings.GetSetting(key)
    End Function
#End Region

#Region "Private methods"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared ReadOnly Property AppSettings As AzureApplicationSettings
        Get
            If _appSettings Is Nothing Then
                SyncLock lock
                    If _appSettings Is Nothing Then
                        _appSettings = New AzureApplicationSettings()
                    End If
                End SyncLock
            End If

            Return _appSettings
        End Get
    End Property
#End Region
End Class
