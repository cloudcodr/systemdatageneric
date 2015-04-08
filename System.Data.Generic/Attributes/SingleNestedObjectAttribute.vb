Public Enum SingleNestedObjectEnum
    SetFirst
    SetLast
End Enum

<AttributeUsage(AttributeTargets.Property, AllowMultiple:=False)>
Public Class SingleNestedObjectAttribute
    Inherits Attribute

    Private _approach As SingleNestedObjectEnum
    Public Sub New(approach As SingleNestedObjectEnum)
        _approach = approach
    End Sub

    Protected Friend ReadOnly Property Approach As SingleNestedObjectEnum
        Get
            Return _approach
        End Get
    End Property
End Class
