Public Class Smtp
    Private Shared _SMTPServer As String
    Private Shared _SMTPPort As Integer
    Private Shared _SMTPSSL As Boolean
    Private Shared _SMTPCuenta As String
    Private Shared _SMTPMail As String
    Private Shared _SMTPUsuario As String
    Private Shared _SMTPClave As String

    Public Shared Property SMTPServer() As String
        Get
            Return _SMTPServer
        End Get
        Set(ByVal value As String)
            _SMTPServer = value
        End Set
    End Property

    Public Shared Property SMTPPort() As Integer
        Get
            Return _SMTPPort
        End Get
        Set(ByVal value As Integer)
            _SMTPPort = value
        End Set
    End Property

    Public Shared Property SMTPSSL() As Boolean
        Get
            Return _SMTPSSL
        End Get
        Set(ByVal value As Boolean)
            _SMTPSSL = value
        End Set
    End Property

    Public Shared Property SMTPCuenta() As String
        Get
            Return _SMTPCuenta
        End Get
        Set(ByVal value As String)
            _SMTPCuenta = value
        End Set
    End Property

    Public Shared Property SMTPMail() As String
        Get
            Return _SMTPMail
        End Get
        Set(ByVal value As String)
            _SMTPMail = value
        End Set
    End Property

    Public Shared Property SMTPUsuario() As String
        Get
            Return _SMTPUsuario
        End Get
        Set(ByVal value As String)
            _SMTPUsuario = value
        End Set
    End Property

    Public Shared Property SMTPClave() As String
        Get
            Return _SMTPClave
        End Get
        Set(ByVal value As String)
            _SMTPClave = value
        End Set
    End Property

End Class
