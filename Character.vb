Imports System.ComponentModel

Public Class Character

    Public Enum GenderType
        <Description("Male")> Male
        <Description("Female")> Female
        <Description("Agender")> Agender
        <Description("Neuter")> Neuter
        <Description("Nonbinary")> Nonbinary
        <Description("polygender")> polygender
        <Description("genderfluid")> genderfluid
        <Description("genderqueer")> genderqueer
        <Description("androgynous")> androgynous
        <Description("bigender")> bigender
        <Description("transgender")> transgender
        <Description("intersex")> intersex
        <Description("Other")> Other
        <Description("Unknown")> Unknown
    End Enum

    Private _Name As String
    Public Property Name As String
        Get
            Return _Name
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then
                Throw New ArgumentException("Name cannot be empty.")
            End If
            _Name = value.Trim()
        End Set
    End Property

    Private _Pseudonyms As List(Of String)
    Public Property Pseudonyms As List(Of String)
        Get
            Return _Pseudonyms
        End Get
        Set(value As List(Of String))
            _Pseudonyms = value
        End Set
    End Property

    Private _Age As Integer
    Public Property Age As Integer
        Get
            Return _Age
        End Get
        Set(value As Integer)
            If value < 0 Then
                Throw New ArgumentException("Age cannot be negative.")
            End If
            _Age = value
        End Set
    End Property

    Private _Gender As GenderType
    Public Property Gender As GenderType
        Get
            Return _Gender
        End Get
        Set(value As GenderType)
            _Gender = value
        End Set
    End Property

    Private _VoiceName As String
    Public Property VoiceName As String
        Get
            Return _VoiceName
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then
                Throw New ArgumentException("VoiceName cannot be empty.")
            End If
            _VoiceName = value.Trim()
        End Set
    End Property

    Sub New(_Name As String, _age As Integer, _Gender As GenderType, _VoiceName As String, Optional _Pseudonyms As List(Of String) = Nothing)
        Name = _Name
        Age = _age
        Gender = _Gender
        VoiceName = _VoiceName
        Pseudonyms = _Pseudonyms
    End Sub

    Sub New(_Name As String)
        Name = _Name
    End Sub

    Sub New()
    End Sub

    Public Overrides Function ToString() As String
        Return $"Name: {Name}, Age: {Age}, Gender: {Gender}, Voice: {VoiceName}, Pseudonym: {Pseudonyms}"
    End Function

End Class
