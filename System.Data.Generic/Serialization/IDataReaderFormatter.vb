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

Namespace Serialization.Formatters
    ''' <summary>
    ''' Provides functionality for formattering DataReader and objects.
    ''' </summary>
    ''' <typeparam name="T">Type of object.</typeparam>
    ''' <remarks></remarks>
    Public Interface IDataReaderFormatter(Of T)
        ''' <summary>
        ''' Serializes an object, or graph as objects with the given root to the provided reader.
        ''' </summary>
        ''' <param name="reader">The DataReader where the formatter puts the serialized data.</param>
        ''' <param name="graph">The object, or root of the object graph, to serialize.</param>
        ''' <remarks></remarks>
        Sub Serialize(reader As IDataReader, graph As T)
        ''' <summary>
        ''' Deserializes the data on the provided DataReader and reconstitutes the graph of objects.
        ''' </summary>
        ''' <param name="reader">The DataReader that contains the data to deserialize.</param>
        ''' <returns>Deserialized object of type T from the reader.</returns>
        ''' <remarks></remarks>
        Function Deserialize(reader As IDataReader) As T        
    End Interface
End Namespace
