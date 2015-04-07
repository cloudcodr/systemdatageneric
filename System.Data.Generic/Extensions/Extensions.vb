' Copyright (c) 2013-15 Arumsoft
' All rights reserved.
'
' Created:  03/04/2015.         Author:   Cloudcoder
'
' This source is subject to the Microsoft Permissive License. http://www.microsoft.com/en-us/openness/licenses.aspx
'
' THE CODE IS PROVIDED ON AN "AS IS" BASIS WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED. YOU EXPRESSLY AGREE 
' THAT THE USE OF THE SERVICE IS AT YOUR SOLE RISK. COMPANY DOES NOT WARRANT THAT THE SERVICE WILL BE UNINTERRUPTED 
' OR ERROR FREE, NOR DOES COMPANY MAKE ANY WARRANTY AS TO ANY RESULTS THAT MAY BE OBTAINED BY USE OF THE SERVICE. 
' COMPANY MAKES NO OTHER WARRANTIES, EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, ANY IMPLIED WARRANTIES OF 
' MERCHANTABILITY OR FITNESS FOR A PARTICULAR PURPOSE, IN RELATION TO THE SERVICE.

Imports System.Runtime.CompilerServices
Imports System.Reflection

''' <summary>
''' Provides the extensions methods.
''' </summary>
''' <remarks></remarks>
Module Extensions
#Region "String extensions"
    ''' <summary>
    ''' Returns a value indicating whether specified System.String value occurs in the string, using the comparision provided.
    ''' </summary>
    ''' <param name="source">Source string to check again.</param>
    ''' <param name="value">Value string to check of.</param>
    ''' <param name="comp">Comparison method to compare using.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Friend Function Contains(source As String, value As String, comp As StringComparison) As Boolean
        Return source.IndexOf(value, comp) >= 0
    End Function
#End Region

#Region "PropertyInfo extensions"
    ''' <summary>
    ''' Returns the full name of the property, as class.property.
    ''' </summary>
    ''' <param name="source">PropertyInfo class.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Friend Function PropertyPath(source As PropertyInfo) As String
        Dim propertyType As Type = source.PropertyType
        Return source.DeclaringType.Name + "." + propertyType.Name
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="source">PropertyInfo class.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Friend Function IsSystemType(source As PropertyInfo) As Boolean
        Dim propertyType As Type = source.PropertyType
        Return Not propertyType.IsClass OrElse propertyType.Equals(GetType(String))
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="source">PropertyInfo class.</param>
    ''' <param name="collectionType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Friend Function IsCollection(source As PropertyInfo, ByRef collectionType As Type) As Boolean
        Dim propertyType As Type = source.PropertyType

        ' TODO: how about simple "HashTable, List, Dictionary etc.

        ' An Object() array
        If propertyType.IsArray Then
            collectionType = propertyType.GetElementType
            Return True
        End If
        ' An Enum...(of T), List(of T) etc.
        If propertyType.IsGenericType Then
            collectionType = propertyType.GetGenericArguments(0)
            Return True
        End If

        collectionType = Nothing
        Return False
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="source">PropertyInfo class.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Friend Function IsCollection(source As PropertyInfo) As Boolean
        Dim propertyType As Type = source.PropertyType

        ' TODO: how about simple "HashTable, List, Dictionary etc.

        ' An Object() array
        If propertyType.IsArray Then
            Return True
        End If
        ' An Enum...(of T), List(of T) etc.
        If propertyType.IsGenericType Then
            Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="source">PropertyInfo class.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Friend Function CollectionItemType(source As PropertyInfo) As Type
        Dim collectionType As Type = Nothing

        IsCollection(source, collectionType)
        Return collectionType
    End Function
#End Region

#Region "Dictionary extensions"
    ''' <summary>
    ''' Converts the value collection of a dictionary to an array.
    ''' </summary>
    ''' <typeparam name="T">Type of items.</typeparam>
    ''' <typeparam name="K">Type of key.</typeparam>
    ''' <param name="source"></param>
    ''' <returns>Array of type T.</returns>
    ''' <remarks></remarks>
    <Extension()>
    Friend Function ValuesToArray(Of K, T)(source As Dictionary(Of K, T)) As T()
        Dim values As T() = New T(source.Count - 1) {}
        source.Values.CopyTo(values, 0)

        Return values
    End Function
#End Region
End Module
