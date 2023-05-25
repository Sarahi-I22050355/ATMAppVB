Friend Class User
    Inherits Person
    Private _nip As String
    Private _balance As Integer

    Public Property NIP As String
        Get
            Return _nip
        End Get
        Set(value As String)
            _nip = value
        End Set
    End Property

    Public Property Balance As Integer
        Get
            Return _balance
        End Get
        Set(value As Integer)
            _balance = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
        _nip = String.Empty
        _balance = 0
    End Sub

    Public Sub New(nip As String, balance As Integer, id As Integer, name As String)
        MyBase.New(name, id)
        Me._nip = nip
        Me._balance = balance
    End Sub
End Class
