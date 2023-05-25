Friend MustInherit Class Person
    Protected _id As Integer
    Protected _name As String

    Public Property ID As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Public Property Name As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    Public Sub New()
        _name = String.Empty
        _id = 0
    End Sub

    Public Sub New(ByVal name As String, ByVal id As Integer)
        Me._id = id
        Me._name = name
    End Sub
End Class

