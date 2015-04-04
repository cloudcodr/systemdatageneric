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

Imports System.Runtime.InteropServices

Namespace Settings
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Class Utilities

        Private Const AssemblyPathMax As Integer = 1024

        ''' <summary>
        ''' Interface for Azure assembly cache.
        ''' </summary>
        ''' <remarks></remarks>
        <ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("e707dcde-d1cd-11d2-bab9-00c04f8eceae")> _
        Private Interface IAssemblyCache
            Sub Reserved0()

            <PreserveSig> _
            Function QueryAssemblyInfo(flags As Integer, <MarshalAs(UnmanagedType.LPWStr)> assemblyName As String, ByRef assemblyInfo As AssemblyInfo) As Integer
        End Interface

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks></remarks>
        <StructLayout(LayoutKind.Sequential)> _
        Private Structure AssemblyInfo
            Public cbAssemblyInfo As Integer
            Public assemblyFlags As Integer
            Public assemblySizeInKB As Long
            <MarshalAs(UnmanagedType.LPWStr)> _
            Public currentAssemblyPath As String
            Public cchBuffer As Integer
            ' size of path buffer.
        End Structure

        <DllImport("fusion.dll")> _
        Private Shared Function CreateAssemblyCache(ByRef ppAsmCache As IAssemblyCache, reserved As Integer) As Integer
        End Function

        ''' <summary>
        ''' Gets an assembly path from the GAC given a partial name.
        ''' </summary>
        ''' <param name="name">An assembly partial name. May not be null.</param>
        ''' <returns>
        ''' The assembly path if found; otherwise null;
        ''' </returns>
        Public Shared Function GetAssemblyPath(name As String) As String
            If name Is Nothing Then
                Throw New ArgumentNullException("name")
            End If

            Dim finalName As String = name
            Dim aInfo As New AssemblyInfo()
            aInfo.cchBuffer = AssemblyPathMax
            aInfo.currentAssemblyPath = New [String](ControlChars.NullChar, aInfo.cchBuffer)

            Dim ac As IAssemblyCache = Nothing
            Dim hr As Integer = CreateAssemblyCache(ac, 0)
            If hr >= 0 Then
                hr = ac.QueryAssemblyInfo(0, finalName, aInfo)
                If hr < 0 Then
                    Return Nothing
                End If
            End If

            Return aInfo.currentAssemblyPath
        End Function
    End Class
End Namespace
