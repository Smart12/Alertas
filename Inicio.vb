Public Class Inicio

    Private Sub Inicio_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim db As New Dao
        Dim obj As DataDao.param = db.params.First()

        Smtp.SMTPServer = obj.PAR_SMTP_SERVER
        Smtp.SMTPCuenta = obj.PAR_SMTP_CUENTA
        Smtp.SMTPMail = obj.PAR_SMTP_MAIL
        Smtp.SMTPUsuario = obj.PAR_SMTP_USUARIO
        Smtp.SMTPClave = obj.PAR_SMTP_CLAVE
        Smtp.SMTPPort = obj.par_smtp_port
        Smtp.SMTPSSL = obj.par_smtp_ssl
    End Sub

    Private Sub Inicio_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Dim fecha As Date = Now
        'fecha = New Date(2014, 11, 28)

        Dim mailbox As String = "noreply@etertin.com"
        Dim cCCO = "esteban@etertin.com;diego@etertin.com;pcansado@etertin.com;" + _
                "wschulkin@etertin.com;jfrias@etertin.com;gonzalo@etertin.com;" + _
                "npignataro@etertin.com;arenzo@etertin.com;gonzalo_giazitzian@hotmail.com;" + _
                "josta@etertin.com;eugenioalonso@xoft.com.ar;isarachaga@etertin.com"

        'cCCO = "eugenioalonso@xoft.com.ar"

        lblStatus.Text = "Estado:  Ejecutando Alerta FACTURADO DIARIO ETERTIN (1/5)"
        lblStatus.Refresh()
        Util.Workflow.FacturacionEtertin(mailbox, cCCO, fecha)

        cCCO = "esteban@etertin.com;diego@etertin.com;pcansado@etertin.com;wschulkin@etertin.com;" + _
                "jfrias@etertin.com;gonzalo@etertin.com;npignataro@etertin.com;arenzo@etertin.com;" + _
                "gonzalo_giazitzian@hotmail.com;josta@etertin.com;" + _
                "mgonzalez@etertin.com;efrezza@etertin.com;sgoenaga@etertin.com;isarachaga@etertin.com"

        lblStatus.Text = "Estado:  Ejecutando Alerta STOCK VALORIZADO (2/5)"
        lblStatus.Refresh()
        Util.Workflow.StockValorizado(mailbox, cCCO, fecha)

        cCCO = "esteban@etertin.com;diego@etertin.com;wschulkin@etertin.com;npignataro@etertin.com;" + _
                "arenzo@etertin.com;isarachaga@etertin.com"
        lblStatus.Text = "Estado:  Ejecutando Alerta GP Diario (3/5)"
        lblStatus.Refresh()
        Util.Workflow.GpDiario(mailbox, cCCO, fecha)

        cCCO = "esteban@etertin.com;diego@etertin.com;wschulkin@etertin.com;gonzalo@etertin.com;josta@etertin.com"
        lblStatus.Text = "Estado:  Ejecutando Alerta MPS (4/5)"
        lblStatus.Refresh()
        Util.Workflow.Mps(mailbox, cCCO, fecha)

        cCCO = "eugenioalonso@xoft.com.ar;pcansado@etertin.com;diego@etertin.com;esteban@etertin.com;npignataro@etertin.com;gonzalo_giazitzian@hotmail.com;mrizzolo@etertin.com"
        'cCCO = "eugenioalonso@xoft.com.ar"
        lblStatus.Text = "Estado: Ejecutando Alerta Información Diaria (5/5)"
        lblStatus.Refresh()
        Util.Workflow.InfoDiaria(mailbox, cCCO, fecha)


        Me.Close()
        End
    End Sub

End Class