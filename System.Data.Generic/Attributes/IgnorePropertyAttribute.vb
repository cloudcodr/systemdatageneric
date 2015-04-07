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
''' Instructs the <see cref="Serialization.Formatters.DataSourceFormatter">DataSourceFormatter</see> method not to serialize/deserialize the public field.
''' </summary>
''' <remarks>For interoperability you can also use the <see cref="Xml.Serialization.XmlIgnoreAttribute">XmlIgnoreAttribute</see>.</remarks>
<AttributeUsage(AttributeTargets.Field Or AttributeTargets.Property, AllowMultiple:=False)>
Public Class IgnorePropertyAttribute
    Inherits Attribute

End Class

