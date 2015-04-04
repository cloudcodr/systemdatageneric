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

Imports System.Configuration
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System

Namespace Settings
    ''' <summary>
    ''' Windows Azure settings.
    ''' </summary>
    Friend NotInheritable Class AzureApplicationSettings

#Region "Private const"
        Private Const DebugOutputEnabled As Boolean = False ' TODO: enable to output
        Private Const RoleEnvironmentTypeName As String = "Microsoft.WindowsAzure.ServiceRuntime.RoleEnvironment"
        Private Const RoleEnvironmentExceptionTypeName As String = "Microsoft.WindowsAzure.ServiceRuntime.RoleEnvironmentException"
        Private Const IsAvailablePropertyName As String = "IsAvailable"
        Private Const GetSettingValueMethodName As String = "GetConfigurationSettingValue"
#End Region

#Region "Member variables"
        ' Keep this array sorted by the version in the descendant order.
        Private ReadOnly knownAssemblyNames As String() = New String() {"Microsoft.WindowsAzure.ServiceRuntime, Culture=neutral, PublicKeyToken=31bf3856ad364e35, ProcessorArchitecture=MSIL"}

        Private _roleEnvironmentExceptionType As Type
        ' Exception thrown for missing settings.
        Private _getServiceSettingMethod As MethodInfo
        ' Method for getting values from the service configuration.
#End Region

#Region "Debug Methods"
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="msg"></param>
        ''' <remarks></remarks>
        <Conditional("DEBUG")>
        Private Shared Sub TraceWriteLine(msg As String)
            If DebugOutputEnabled Then
                Trace.WriteLine(msg)
            End If
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="condition"></param>
        ''' <remarks></remarks>
        <Conditional("DEBUG")>
        Private Shared Sub DebugAssert(condition As Boolean)
            If DebugOutputEnabled Then
                Debug.Assert(condition)
            End If
        End Sub
#End Region

#Region "Constructor"
        ''' <summary>
        ''' Initializes the settings.
        ''' </summary>
        Friend Sub New()
            ' Find out if the code is running in the cloud service context.
            Dim assembly As Assembly = GetServiceRuntimeAssembly()
            If assembly IsNot Nothing Then
                Dim roleEnvironmentType As Type = assembly.[GetType](RoleEnvironmentTypeName, False)
                _roleEnvironmentExceptionType = assembly.[GetType](RoleEnvironmentExceptionTypeName, False)
                If roleEnvironmentType IsNot Nothing Then
                    Dim isAvailableProperty As PropertyInfo = roleEnvironmentType.GetProperty(IsAvailablePropertyName)
                    Dim isAvailable As Boolean


                    Try
                        isAvailable = isAvailableProperty IsNot Nothing AndAlso CBool(isAvailableProperty.GetValue(Nothing, New Object() {}))
                        Dim message As String = String.Format(CultureInfo.InvariantCulture, "Loaded ""{0}""", assembly.FullName)
                        TraceWriteLine(message)
                    Catch e As TargetInvocationException
                        ' Running service runtime code from an application targeting .Net 4.0 results
                        ' in a type initialization exception unless application's configuration file
                        ' explicitly enables v2 runtime activation policy. In this case we should fall
                        ' back to the web.config/app.config file.
                        If Not (TypeOf e.InnerException Is TypeInitializationException) Then
                            Throw
                        End If
                        isAvailable = False
                    End Try


                    If isAvailable Then
                        _getServiceSettingMethod = roleEnvironmentType.GetMethod(GetSettingValueMethodName, BindingFlags.[Public] Or BindingFlags.[Static] Or BindingFlags.InvokeMethod)
                    End If
                End If
            End If
        End Sub
#End Region

#Region "Public methods"
        ''' <summary>
        ''' Gets a setting with the given name.
        ''' </summary>
        ''' <param name="name">Setting name.</param>
        ''' <returns>Setting value or null if such setting does not exist.</returns>
        Friend Function GetSetting(name As String) As String
            DebugAssert(Not String.IsNullOrEmpty(name))
            Dim value As String = Nothing

            value = GetValue("ServiceRuntime", name, AddressOf GetServiceRuntimeSetting)
            If value Is Nothing Then
                value = GetValue("ConfigurationManager", name, Function(n) ConfigurationManager.AppSettings(n))
            End If

            Return value
        End Function
#End Region

#Region "Private Methods"
        ''' <summary>
        ''' Checks whether the given exception represents an exception throws
        ''' for a missing setting.
        ''' </summary>
        ''' <param name="e">Exception</param>
        ''' <returns>True for the missing setting exception.</returns>
        Private Function IsMissingSettingException(e As Exception) As Boolean
            If e Is Nothing Then
                Return False
            End If
            Dim type As Type = e.[GetType]()


            Return Object.ReferenceEquals(type, _roleEnvironmentExceptionType) OrElse type.IsSubclassOf(_roleEnvironmentExceptionType)
        End Function

        ''' <summary>
        ''' Gets setting's value from the given provider.
        ''' </summary>
        ''' <param name="providerName">Provider name.</param>
        ''' <param name="settingName">Setting name</param>
        ''' <param name="getValue1">Method to obtain given setting.</param>
        ''' <returns>Setting value, or null if not found.</returns>
        Private Shared Function GetValue(providerName As String, settingName As String, getValue1 As Func(Of String, String)) As String
            Dim value As String = getValue1(settingName)
            Dim message As String

            If value IsNot Nothing Then
                message = String.Format(CultureInfo.InvariantCulture, "PASS ({0})", value)
            Else
                message = "FAIL"
            End If

            message = String.Format(CultureInfo.InvariantCulture, "Getting ""{0}"" from {1}: {2}.", settingName, providerName, message)
            TraceWriteLine(message)
            Return value
        End Function

        ''' <summary>
        ''' Gets a configuration setting from the service runtime.
        ''' </summary>
        ''' <param name="name">Setting name.</param>
        ''' <returns>Setting value or null if not found.</returns>
        Private Function GetServiceRuntimeSetting(name As String) As String
            DebugAssert(Not String.IsNullOrEmpty(name))

            Dim value As String = Nothing

            If _getServiceSettingMethod IsNot Nothing Then
                Try
                    value = DirectCast(_getServiceSettingMethod.Invoke(Nothing, New Object() {name}), String)
                Catch e As TargetInvocationException
                    If Not IsMissingSettingException(e.InnerException) Then
                        Throw
                    End If
                End Try
            End If
            Return value
        End Function

        ''' <summary>
        ''' Loads and returns the latest available version of the service 
        ''' runtime assembly.
        ''' </summary>
        ''' <returns>Loaded assembly, if any.</returns>
        Private Function GetServiceRuntimeAssembly() As Assembly
            Dim assembly__1 As Assembly = Nothing

            For Each assemblyName As String In knownAssemblyNames
                Dim assemblyPath As String = Utilities.GetAssemblyPath(assemblyName)

                Try
                    If Not String.IsNullOrEmpty(assemblyPath) Then
                        assembly__1 = Assembly.LoadFrom(assemblyPath)
                    End If
                Catch e As Exception
                    ' The following exceptions are ignored for enabling configuration manager to proceed
                    ' and load the configuration from application settings instead of using ServiceRuntime.
                    If Not (TypeOf e Is FileNotFoundException OrElse TypeOf e Is FileLoadException OrElse TypeOf e Is BadImageFormatException) Then
                        Throw
                    End If
                End Try
            Next

            Return assembly__1
        End Function

        ''' <summary>
        ''' Gets the setting defined in the Windows Azure configuration file.
        ''' </summary>
        ''' <param name="name">Setting name.</param>
        ''' <returns>Setting value.</returns>
        Private Function GetServiceSetting(name As String) As String
            If _getServiceSettingMethod IsNot Nothing Then
                Return DirectCast(_getServiceSettingMethod.Invoke(Nothing, New Object() {name}), String)
            End If

            Return Nothing
        End Function
#End Region
    End Class
End Namespace
