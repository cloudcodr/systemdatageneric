<AttributeUsage(AttributeTargets.Property Or AttributeTargets.Field, AllowMultiple:=False)>
Public Class GenericMapAttribute
    Inherits Attribute

    Private _dbColumn As String

    Public Sub New(dataBaseColumn As String)
        _dbColumn = dataBaseColumn
    End Sub

    Public ReadOnly Property DatabaseColumn As String
        Get
            Return _dbColumn
        End Get
    End Property
End Class
