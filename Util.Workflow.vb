Imports System.Net.Mail

Namespace Util
    Public Class Workflow

        Public Shared id_total As Decimal

        Public Shared Sub EnviarMail(ByVal cFrom As String, ByVal cTo As String, ByVal cCC As String, ByVal cCCO As String, ByVal cSubject As String, ByVal cHtmlBody As String, ByVal cFile As String)
            If cTo.Trim = "" Then
                Return
            End If

            If cFrom = "" Then
                cFrom = Smtp.SMTPMail
            End If

            ''''''''' envio mail
            Dim message As MailMessage = New MailMessage()
            message.From = New MailAddress(cFrom, Smtp.SMTPCuenta)

            Dim vTo() As String = cTo.Split(";")
            For i As Integer = 0 To vTo.Count() - 1
                message.To.Add(New MailAddress(vTo(i)))
            Next

            If cCC.Trim <> "" Then
                Dim vCC() As String = cCC.Split(";")
                For i As Integer = 0 To vCC.Count() - 1
                    message.CC.Add(New MailAddress(vCC(i)))
                Next
            End If

            If cCCO.Trim <> "" Then
                Dim vCCO() As String = cCCO.Split(";")
                For i As Integer = 0 To vCCO.Count() - 1
                    message.Bcc.Add(New MailAddress(vCCO(i)))
                Next
            End If

            message.Subject = cSubject
            message.IsBodyHtml = True
            message.Body = cHtmlBody

            If cFile <> "" Then
                Dim attach As New System.Net.Mail.Attachment(cFile)
                message.Attachments.Add(attach)
            End If

            Dim client As SmtpClient = New SmtpClient()

            client.Credentials = New System.Net.NetworkCredential(Smtp.SMTPUsuario, Smtp.SMTPClave)

            ''----------- Esto es para que vaya a través de SSL que es obligatorio con GMail
            client.Port = Smtp.SMTPPort
            client.EnableSsl = Smtp.SMTPSSL
            ''--------------------------------------------
            client.Host = Smtp.SMTPServer
            client.Send(message)
        End Sub

        Private Shared Function Eml2Str_OR(ByVal pEML As String, ByVal Inc_custom4 As String, ByVal Req_nombre As String) As String

            pEML = pEML.Replace("#INC_CUSTOM4#", Inc_custom4)
            pEML = pEML.Replace("#REQ_NOMBRE#", Req_nombre)

            Return pEML
        End Function

        Public Shared Sub FacturacionEtertin(ByVal mail As String, ByVal CCO As String, ByVal fecha As Date)
            Dim AuxTit, AuxMen As String

            Dim lcSql As String = "select total = isnull(sum(case when fac_tipmov <> 'NC' then fac_neto1+fac_neto2 else 0 end),0),  " + _
                    " total_nc = isnull(sum(case when fac_tipmov = 'NC' then fac_neto1+fac_neto2 else 0 end),0) " + _
                    " from facvecab " + _
                    " where convert(varchar,fac_fecha,112) = '" + fecha.ToString("yyyyMMdd") + "' and fac_financiera=0"

            Dim lcSqlAcum As String = "select total = isnull(sum(case when fac_tipmov <> 'NC' then fac_neto1+fac_neto2 else -(fac_neto1+fac_neto2) end),0)  " + _
                    " from facvecab " + _
                    " where Month(fac_fecha) = " + Month(fecha).ToString() + _
                    " and year(fac_fecha) = " + Year(fecha).ToString() + _
                    " and fac_financiera=0 "

            Dim db As New Dao

            Dim importe = db.ExecuteQuery(Of vm_VtaRubro)(lcSql + " and cli_codigo not in (15711,15885) ").First()
            Dim importeMPS = db.ExecuteQuery(Of vm_VtaRubro)(lcSql + " and cli_codigo = 15711").First()
            Dim importeIntermaco = db.ExecuteQuery(Of vm_VtaRubro)(lcSql + " and cli_codigo = 15885").First()

            Dim acum = db.ExecuteQuery(Of Decimal)(lcSqlAcum + " and cli_codigo not in (15711,15885)").First()
            Dim acumMPS = db.ExecuteQuery(Of Decimal)(lcSqlAcum + " and cli_codigo = 15711").First()
            Dim acumIntermaco = db.ExecuteQuery(Of Decimal)(lcSqlAcum + " and cli_codigo = 15885").First()

            Dim lcSqlCanAcum As String = "select isnull(count(*),0) as cant from (select distinct cli_codigo from facvecab " + _
        " where Month(fac_fecha) = " + Month(fecha).ToString() + _
            " and year(fac_fecha) = " + Year(fecha).ToString() + _
            " and fac_financiera=0 ) a "

            Dim cancli_acum = db.ExecuteQuery(Of Integer)(lcSqlCanAcum).First()



            AuxTit = "Alerta - Facturación del día: " + fecha.ToShortDateString()
            AuxMen = "<h2>Facturacion del dia: " + fecha.ToShortDateString() + " (sin IVA)</h2>" + _
                "<table border=""1""><tr><th>Empresa </th><th>Total Facturado $ </th><th>Total NC $ </th><th>Total $ </th><th>Acumulado Mensual $</th></tr>" + _
                "<tr><td>Mayorista</td><td align=""right"">" + importe.total.ToString("N") + "</td><td align=""right"">" + importe.total_nc.ToString("N") + "</td><td align=""right"">" + (importe.total - importe.total_nc).ToString("N") + "</td><td align=""right"">" + (acum).ToString("N") + "</td></tr>" + _
                "<tr><td>Distribuidores</td><td align=""right"">" + importeIntermaco.total.ToString("N") + "</td><td align=""right"">" + importeIntermaco.total_nc.ToString("N") + "</td><td align=""right"">" + (importeIntermaco.total - importeIntermaco.total_nc).ToString("N") + "</td><td align=""right"">" + (acumIntermaco).ToString("N") + "</td></tr>" + _
                "<tr><td>MPS</td><td align=""right"">" + importeMPS.total.ToString("N") + "</td><td align=""right"">" + importeMPS.total_nc.ToString("N") + "</td><td align=""right"">" + (importeMPS.total - importeMPS.total_nc).ToString("N") + "</td><td align=""right"">" + (acumMPS).ToString("N") + "</td></tr>" + _
                "</table>"

            AuxMen = AuxMen + "<h2>Facturación por Familia (sin IVA)</h2>"

            lcSql = "select rub.rub_codigo, rub.rub_Descri, " + _
 " acumulado = sum(case when fac_tipmov <> 'NC' then rr.fac_cantid*rr.art_pu else -(rr.fac_cantid*rr.art_pu) end), " + _
 " total = (select isnull(sum(case when f.fac_tipmov <> 'NC' then r.fac_cantid*r.art_pu else 0 end),0) " + _
                    " from facvecab f, facveren r, articulos a " + _
                     " where convert(varchar,f.fac_fecha,112) = '" + fecha.ToString("yyyyMMdd") + "'  " + _
                     " and f.cli_codigo  not in (15711,15885)  " + _
                     " and f.fac_financiera=0  " + _
                     " and f.fac_codigo = r.fac_codigo   " + _
                     " and r.art_codigo = a.art_codigo  " + _
                     " and a.rub_codigo = rub.rub_codigo ), " + _
 " total_nc = (select isnull(sum(case when f.fac_tipmov = 'NC' then r.fac_cantid*r.art_pu else 0 end),0) " + _
                    " from facvecab f, facveren r, articulos a " + _
                     " where convert(varchar,f.fac_fecha,112) = '" + fecha.ToString("yyyyMMdd") + "'  " + _
                     " and f.cli_codigo  not in (15711,15885)  " + _
                     " and f.fac_financiera=0  " + _
                     " and f.fac_codigo = r.fac_codigo   " + _
                     " and r.art_codigo = a.art_codigo  " + _
                     " and a.rub_codigo = rub.rub_codigo ),  " + _
  " cant_cli = (select count(distinct cli_codigo)" + _
                     " from facvecab f, facveren r, articulos a " + _
            " where Month(f.fac_fecha) = " + Month(fecha).ToString() + _
        " And Year(f.fac_fecha) = " + Year(fecha).ToString() + _
                    " and f.cli_codigo  not in (15711,15885)  " + _
                    " and f.fac_financiera=0  " + _
                    " and f.fac_codigo = r.fac_codigo   " + _
                    " and r.art_codigo = a.art_codigo  " + _
                    " and a.rub_codigo = rub.rub_codigo ) " + _
   " from facvecab ff, facveren rr, articulos aa , rubros rub " + _
            " where Month(ff.fac_fecha) = " + Month(fecha).ToString() + _
  "  And Year(ff.fac_fecha) = " + Year(fecha).ToString() + _
  "  and ff.cli_codigo  not in (15711,15885) " + _
  "  and ff.fac_financiera=0  " + _
  "  and ff.fac_codigo = rr.fac_codigo   " + _
  "  and rr.art_codigo = aa.art_codigo   " + _
  "  and aa.rub_codigo = rub.rub_codigo  " + _
" group by rub.rub_codigo, rub.rub_Descri " + _
" order by 3 desc "

            Dim list = db.ExecuteQuery(Of vm_VtaRubro)(lcSql)

            AuxMen = AuxMen + "<table border=""1""><tr><th>Familia</th><th>Total Facturado $ </th><th>Total NC $</th><th>Total $</th><th>Acumulado Mensual $</th><th>Cant.Clientes</th></tr>"

            Dim tot As Decimal = 0.0
            Dim totnc As Decimal = 0.0
            Dim totacum As Decimal = 0.0

            For Each obj In list
                AuxMen = AuxMen + "<tr><td>" + obj.rub_descri + "</td><td align=""right"">" + obj.total.ToString("N") + "</td><td align=""right"">" + obj.total_nc.ToString("N") + "</td><td align=""right"">" + (obj.total - obj.total_nc).ToString("N") + "</td><td align=""right"">" + obj.acumulado.ToString("N") + "</td><td align=""right"">" + obj.cant_cli.ToString() + "</td></tr>"

                tot = tot + obj.total
                totnc = totnc + obj.total_nc
                totacum = totacum + obj.acumulado
            Next
            AuxMen = AuxMen + "<tr><td>Totales</td><td align=""right"">" + tot.ToString("N") + "</td><td align=""right"">" + totnc.ToString("N") + "</td><td align=""right"">" + (tot - totnc).ToString("N") + "</td><td align=""right"">" + totacum.ToString("N") + "</td><td  align=""right"">" + cancli_acum.ToString() + "</td></tr>"
            AuxMen = AuxMen + "</table>"

            EnviarMail("", mail, "", CCO, AuxTit, AuxMen, "")
        End Sub

        Public Shared Sub StockValorizado(ByVal mail As String, ByVal CCO As String, ByVal fecha As Date)
            Dim AuxTit, AuxMen As String
            Dim lcSql As String

            Dim db As New Dao

            AuxTit = "Alerta - Stock Valorizado al día: " + fecha.ToShortDateString()
            AuxMen = "<h1>Stock Valorizado al dia: " + fecha.ToShortDateString() + "</h1>"

            Dim lisdep = db.depositos.Where(Function(x) x.dep_activo = True).ToList()
            Dim xx As Integer
            Dim auxRub As String
            Dim totdp, totdd As Decimal
            Dim totp, totd As Decimal

            Dim otr_p, otr_d As Decimal

            Dim dolar As Decimal
            Try
                dolar = db.dolars.Single(Function(x) x.dol_fecha = Now.Date).dol_venta
            Catch ex As Exception
                dolar = 0
            End Try

            If dolar = 0 Then
                Return
            End If

            totp = 0
            totd = 0

            Dim otros As Boolean

            For Each dep In lisdep

                lcSql = "select r.mar_codigo, r.mar_descri, sum(s.art_stock)  as stock, SUM(a.art_costo * s.art_stock) as pesos, SUM(a.art_costo * s.art_stock / " + dolar.ToString().Replace(",", ".") + ") as dolares " + _
                        " from stock s, articulos a, marcas r " + _
                        " where s.dep_codigo = " + dep.dep_codigo.ToString() + _
                        " and s.art_codigo = a.art_codigo " + _
                        " and a.mar_codigo = r.mar_codigo " + _
                        " and a.art_moneda = 'P' " + _
                        " group by r.mar_codigo,mar_descri having sum(s.art_stock * a.art_costo) > 0 " + _
                        " union " + _
                        " select r.mar_codigo, r.mar_descri, sum(s.art_stock)  as stock, SUM(a.art_costo * s.art_stock * " + dolar.ToString().Replace(",", ".") + ") as pesos, SUM(a.art_costo * s.art_stock) as dolares " + _
                        " from stock s, articulos a, marcas r " + _
                        " where s.dep_codigo = " + dep.dep_codigo.ToString() + _
                        " and s.art_codigo = a.art_codigo " + _
                        " and a.mar_codigo = r.mar_codigo  " + _
                        " and a.art_moneda = 'D' " + _
                        " group by r.mar_codigo, mar_descri having sum(s.art_stock * a.art_costo) > 0 " + _
                        " order by 4 desc "

                Dim list = db.ExecuteQuery(Of vm_StockVal)(lcSql)

                xx = 0
                auxRub = "<table border=""1""><tr><th>Marca</th><th>Total $</th><th>Total U$S</th></tr>"

                totdp = 0
                totdd = 0
                otros = False

                otr_d = 0
                otr_p = 0
                For Each rub In list


                    If rub.dolares >= 5000 Then

                        auxRub = auxRub + "<tr><td>" + rub.mar_descri + "</td>" + _
                                          "<td align=""right"">" + rub.pesos.ToString("N") + "</td>" + _
                                          "<td align=""right"">" + rub.dolares.ToString("N") + "</td></tr>"
                    Else
                        otros = True

                        otr_d = otr_d + rub.dolares
                        otr_p = otr_p + rub.pesos
                    End If


                    totdp = totdp + rub.pesos
                    totdd = totdd + rub.dolares

                    xx = xx + 1
                Next
                If otros Then
                    auxRub = auxRub + "<tr><td>Otros</td>" + _
                                      "<td align=""right"">" + otr_p.ToString("N") + "</td>" + _
                                      "<td align=""right"">" + otr_d.ToString("N") + "</td></tr>"
                End If

                auxRub = auxRub + "<tr><td>Totales</td>" + _
                                  "<td align=""right"">" + totdp.ToString("N") + "</td>" + _
                                  "<td align=""right"">" + totdd.ToString("N") + "</td></tr>"

                auxRub = auxRub + "</table>"


                totp = totp + totdp
                totd = totd + totdd


                If xx > 0 Then
                    AuxMen = AuxMen + "<h2>Deposito: " + dep.dep_descri + "</h2>" + auxRub
                End If
            Next

            AuxMen = AuxMen + "<h2>Totales: " + _
                  "$ " + totp.ToString("N") + _
                  " | U$D " + totd.ToString("N") + "</h2>"


            EnviarMail("", mail, "", CCO, AuxTit, AuxMen, "")
        End Sub

        Public Shared Sub GpDiario(ByVal mail As String, ByVal CCO As String, ByVal fecha As Date)
            Dim AuxTit, AuxMen As String

            Dim db As New Dao

            AuxTit = "Alerta - GP por familia del día: " + fecha.ToShortDateString()
            AuxMen = "<h2>GP por familia del dia: " + fecha.ToShortDateString() + "</h2>"

            Dim lcSql = _
                     " select aa.rub_codigo, r.rub_descri, SUM(rr.art_pu*rr.fac_cantid) as venta, SUM(rr.art_costo*rr.fac_cantid) as costo, " + _
                            " SUM(case when cc.fac_tipmov = 'NC' then rr.art_pu*(-rr.fac_cantid) else rr.art_pu*rr.fac_cantid end) as ventaacum " + _
                         " into #c1 from facvecab cc, facveren rr, articulos aa, rubros r " + _
                        " where Month(cc.fac_fecha) = " + fecha.Month.ToString() + " And Year(cc.fac_fecha) = " + fecha.Year.ToString() + _
                        "  and cc.cli_codigo  not in (15711,15885) " + _
                         " and cc.fac_codigo = rr.fac_codigo  " + _
                         " and rr.art_codigo > 0  " + _
                         " and rr.art_codigo = aa.art_codigo  " + _
                         " and aa.rub_codigo = r.rub_codigo " + _
                         " group by aa.rub_codigo,r.rub_descri ;" + _
                    "select rub.rub_codigo,rub.rub_descri, SUM(r.art_pu*r.fac_cantid) as venta, SUM(r.art_costo*r.fac_cantid) as costo, " + _
                            " 	SUM(case when c.fac_tipmov = 'NC' then r.art_pu*(-r.fac_cantid) else r.art_pu*r.fac_cantid end) as ventaneta, " + _
                            " max((1-(r.art_costo/(r.art_pu*0.95)))*100) as maximo, " + _
                            " min((1-(r.art_costo/(r.art_pu*0.95)))*100) as minimo " + _
                    " into #c2 from facvecab c, facveren r, articulos a, rubros rub " + _
                    " where c.fac_fecha >= '" + fecha.ToString("yyyyMMdd") + "' and c.fac_fecha < '" + fecha.AddDays(1).ToString("yyyyMMdd") + "'" + _
                            " and c.cli_codigo  not in (15711,15885) " + _
                            " and c.fac_codigo = r.fac_codigo" + _
                            " and r.art_codigo > 0" + _
                            " and r.art_codigo = a.art_codigo " + _
                            " and a.rub_codigo = rub.rub_codigo " + _
                            " group by rub.rub_codigo,rub.rub_descri ; " + _
                    " select vmes.rub_descri, isnull(vd.venta,0.00) as venta, isnull(vd.costo,0.00) as costo, isnull(vd.ventaneta,0.00) as ventaneta, " + _
                        " isnull(vd.maximo,0.00) as maximo, isnull(vd.minimo,0.00) as minimo, vmes.venta as ventames, vmes.costo as costomes, vmes.ventaacum " + _
                        " from #c1 vmes left outer join #c2 vd on vd.rub_codigo = vmes.rub_codigo " + _
                        " order by vmes.rub_codigo ; " + _
                    " drop table #c1 ; " + _
                    " drop table #c2 "

            Dim list = db.ExecuteQuery(Of vm_GpDiario)(lcSql)

            AuxMen = AuxMen + "<table border=""1""><tr><th>Familia</th><th>Venta del dia (nc neteadas)</th><th>GP Promedio del dia</th><th>GP Maximo</th><th>GP Minimo</th><th>Venta acumulada del mes</th><th>GP Promedio del mes</th></tr>"

            Dim totvtaneta As Decimal = 0.0
            Dim totvtaacum As Decimal = 0.0

            Dim totcosto As Decimal = 0.0
            Dim totventa As Decimal = 0.0

            Dim totcostomes As Decimal = 0.0
            Dim totventames As Decimal = 0.0

            Dim gp As Decimal
            For Each obj In list
                gp = IIf(obj.venta = 0, 0, Math.Round((1 - (obj.costo / (obj.venta * 0.95))) * 100, 2))

                AuxMen = AuxMen + "<tr><td>" + obj.rub_descri + "</td><td align=""right"">" + _
                        obj.ventaneta.ToString("N") + "</td><td align=""right"">" + _
                        gp.ToString("N") + "</td><td align=""right"">" + _
                        Math.Round(obj.maximo, 2).ToString("N") + "</td><td align=""right"">" + _
                        Math.Round(obj.minimo, 2).ToString("N") + "</td><td align=""right"">" + _
                        Math.Round(obj.ventaacum, 2).ToString("N") + "</td><td align=""right"">" + _
                        Math.Round((1 - (obj.costomes / (obj.ventames * 0.95))) * 100, 2).ToString("N") + "</td></tr>"

                totvtaneta = totvtaneta + obj.ventaneta
                totvtaacum = totvtaacum + obj.ventaacum
                totcosto = totcosto + obj.costo
                totventa = totventa + obj.venta
                totcostomes = totcostomes + obj.costomes
                totventames = totventames + obj.ventames
            Next

            AuxMen = AuxMen + "<tr><th>Totales</th><td align=""right"">" + _
                        totvtaneta.ToString("N") + "</td><td align=""right"">" + _
                        Math.Round((1 - (totcosto / (totventa * 0.95))) * 100, 2).ToString("N") + "</td><td align=""right"">" + _
                        "</td><td align=""right"">" + _
                        "</td><td align=""right"">" + _
                        totvtaacum.ToString("N") + "</td><td align=""right"">" + _
                        Math.Round((1 - (totcostomes / (totventames * 0.95))) * 100, 2).ToString("N") + "</td></tr>"


            AuxMen = AuxMen + "</table>"

            EnviarMail("", mail, "", CCO, AuxTit, AuxMen, "")
        End Sub

        Public Shared Sub Mps(ByVal mail As String, ByVal CCO As String, ByVal fecha As Date)
            Dim AuxTit, AuxMen As String

            Dim db As New Dao

            AuxTit = "Alerta - Pedidos MPS del día: " + fecha.ToShortDateString()
            AuxMen = "<h2>Pedidos MPS del dia: " + fecha.ToShortDateString() + "</h2>"

            Dim lcSql = _
                     " select c.ped_codigo, d.cld_destinatario, a.art_fabcodigo, a.art_descri, r.ped_cantidad, dep_descri " + _
                        " from hppedidoscab c, hppedidosren r, articulos  a, clientes_dir d, depositos dep " + _
                        " where ped_fecha >= '" + fecha.ToString("yyyyMMdd") + "' and ped_fecha < '" + fecha.AddDays(1).ToString("yyyyMMdd") + "' " + _
                            " and c.ped_anulado=0 " + _
                            " and c.ped_codigo = r.ped_codigo " + _
                            " and r.art_codigo = a.art_codigo " + _
                            " and c.cld_codigo = d.cld_codigo " + _
                            " and r.dep_codigo = dep.dep_codigo " + _
                        " order by 1"

            Dim list = db.ExecuteQuery(Of vm_Mps)(lcSql)

            AuxMen = AuxMen + "<table border=""1""><tr><th>Nº Pedido</th><th>Destintario</th><th>Codigo</th><th>Descripción</th><th>Cantidad</th><th>Deposito</th></tr>"

            For Each obj In list

                AuxMen = AuxMen + "<tr><td>" + obj.ped_codigo.ToString() + "</td><td>" + _
                        obj.cld_destinatario + "</td><td>" + _
                        obj.art_fabcodigo + "</td><td>" + _
                        obj.art_descri + "</td><td align=""right"">" + _
                        obj.ped_cantidad.ToString() + "</td><td>" + _
                        obj.dep_descri + "</td></tr>"
            Next


            AuxMen = AuxMen + "</table>"

            EnviarMail("", mail, "", CCO, AuxTit, AuxMen, "")
        End Sub

        Public Shared Sub InfoDiaria(ByVal mail As String, ByVal CCO As String, ByVal fecha As Date)
            Dim AuxTit, AuxMen As String
            id_total = 0
            Dim db As New Dao

            AuxTit = "Alerta - Información Diaria al día: " + fecha.ToShortDateString()
            AuxMen = "<h1>Etertin S.A. al dia: " + fecha.ToShortDateString() + "</h1>"

            Dim dolar As Decimal
            Try
                dolar = db.dolars.Single(Function(x) x.dol_fecha = fecha.Date).dol_venta_bna
            Catch ex As Exception
                Return
            End Try

            AuxMen = AuxMen + "<h2>Dolar: " + dolar.ToString() + "</h2>"

            AuxMen = AuxMen + InfoD_GetBancos(dolar)
            AuxMen = AuxMen + InfoD_GetStock(dolar)
            AuxMen = AuxMen + InfoD_Deudores(dolar)
            AuxMen = AuxMen + InfoD_Cartera(dolar)
            AuxMen = AuxMen + InfoD_Proveedores(dolar)

            AuxMen = AuxMen + "<table border=""1""><tr><th>Resultado Patrimonial USD</th><th align=""right"">" + id_total.ToString("N") + "</th></tr>"

            EnviarMail("", mail, "", CCO, AuxTit, AuxMen, "")
        End Sub

        Public Shared Function InfoD_GetBancos(ByVal dolar As Decimal) As String
            Dim db As New Dao
            Dim tot As Decimal
            Dim totban As Decimal
            Dim auxmen As String = "<table border=""1""><tr><th>Bancos</th><th>Total USD</th></tr>"
            Dim lcsql As String

            ''busco caja
            lcsql = _
            "select r.ren_fecha, r.ren_saldo from RendicionesCajaCab r " + _
            " where r.ren_fecha = (select max(ren_fecha) from RendicionesCajaCab x ) "

            Dim caja = db.ExecuteQuery(Of vm_Caja)(lcsql).Single()


            totban = Math.Round((caja.ren_saldo + GetSaldoCaja(caja.ren_fecha, 1871)) / dolar, 2)

            tot = totban
            auxmen = auxmen + "<tr><td>CAJA</td><td align=""right"">" + totban.ToString("N") + "</td></tr>"


            ''busco bancos
            lcsql = _
                "select r.cba_codigo, c.cba_nrocue, c.pla_codigo, r.ren_fecha, r.ren_saldo   " + _
                " from RendicionesBancoCab r, cuentas c " + _
                " where r.cba_codigo = c.cba_codigo " + _
                " AND r.ren_fecha = (select max(ren_fecha) from RendicionesBancoCab x where x.cba_codigo = r.cba_codigo ) "

            Dim list = db.ExecuteQuery(Of vm_Bancos)(lcsql)


            For Each banco In list

                totban = Math.Round((banco.ren_saldo + GetSaldoBanco(banco.ren_fecha, banco.cba_codigo, banco.pla_codigo)) / dolar, 2)
                tot = tot + totban
                auxmen = auxmen + "<tr><td>" + banco.cba_nrocue + "</td><td align=""right"">" + totban.ToString("N") + "</td></tr>"

            Next

            auxmen = auxmen + "<tr><td><b>Total USD</b></td><td align=""right""><b>" + tot.ToString("N") + "</b></td></tr>"
            auxmen = auxmen + "</table><br>"

            id_total = id_total + tot

            Return auxmen
        End Function

        Public Shared Function InfoD_GetStock(ByVal dolar As Decimal) As String
            Dim db As New Dao
            Dim AuxMen As String = ""

            Dim lcSql As String = _
             "select SUM(case when a.art_moneda = 'P' then a.art_costo * s.art_stock / " + dolar.ToString().Replace(",", ".") + " else a.art_costo * s.art_stock end) as dolares " + _
                    " from stock s, articulos a " + _
                    " where s.dep_codigo = 1 " + _
                    " and s.art_codigo = a.art_codigo "

            Dim stock As Decimal = db.ExecuteQuery(Of Decimal)(lcSql).First()

            AuxMen = AuxMen + "<table border=""1""><tr><th>Stock</th><th>Total USD</th></tr>"
            AuxMen = AuxMen + "<tr><td>Entre Rios 1866</td><td align=""right"">" + stock.ToString("N") + "</td></tr>"
            AuxMen = AuxMen + "<tr><td><b>Totales</b></td><td align=""right""><b>" + stock.ToString("N") + "</b></td></tr>"
            AuxMen = AuxMen + "</table><br>"

            id_total = id_total + stock

            Return AuxMen
        End Function

        Public Shared Function InfoD_Deudores(ByVal dolar As Decimal) As String
            Dim db As New Dao
            Dim AuxMen As String = ""

            Dim lcSql As String = _
             "select isnull(sum(CASE WHEN cc.cta_tipmov = 'NC' or cc.cta_tipmov = 'PA' THEN -cc.cta_saldo ELSE cc.cta_saldo END),0) " + _
                "from ctactecl cc where cc.cta_fecha < '" + Now.AddDays(1).ToString("yyyyMMdd") + "' and cta_saldo<>0 "

            Dim deuda As Decimal = Math.Round(db.ExecuteQuery(Of Decimal)(lcSql).First() / dolar, 2)

            AuxMen = AuxMen + "<table border=""1""><tr><th>Deudores por Venta</th><th>Total USD</th></tr>"
            AuxMen = AuxMen + "<tr><td>USD</td><td align=""right"">" + deuda.ToString("N") + "</td></tr>"
            AuxMen = AuxMen + "<tr><td><b>Total USD</b></td><td align=""right""><b>" + deuda.ToString("N") + "</b></td></tr>"
            AuxMen = AuxMen + "</table><br>"

            id_total = id_total + deuda

            Return AuxMen
        End Function

        Public Shared Function InfoD_Cartera(ByVal dolar As Decimal) As String
            Dim db As New Dao
            Dim AuxMen As String = ""

            Dim lcSql As String = _
             "SELECT SUM(che_import) as total " + _
             " FROM cheques c " + _
             " WHERE c.che_deposi = 0 And che_acredi = 0 And che_rechaz = 0 And che_devuelto = 0 And che_pagoa3 = 0 And che_suelto = 0 And che_canjeado = 0 And liq_codigo = 0	"

            Dim deuda As Decimal = Math.Round(db.ExecuteQuery(Of Decimal)(lcSql).First() / dolar, 2)

            AuxMen = AuxMen + "<table border=""1""><tr><th>Cheques en Cartera</th><th>Total USD</th></tr>"
            AuxMen = AuxMen + "<tr><td>USD</td><td align=""right"">" + deuda.ToString("N") + "</td></tr>"
            AuxMen = AuxMen + "<tr><td><b>Total USD</b></td><td align=""right""><b>" + deuda.ToString("N") + "</b></td></tr>"
            AuxMen = AuxMen + "</table><br>"

            id_total = id_total + deuda

            Return AuxMen
        End Function

        Public Shared Function InfoD_Proveedores(ByVal dolar As Decimal) As String
            Dim db As New Dao
            Dim AuxMen As String = ""


            '' CUENTA CORRIENTE
            Dim lcSql As String = _
             "select isnull(sum(CASE WHEN cc.cta_tipmov = 'NC' or cc.cta_tipmov = 'PA' THEN -cc.cta_saldo ELSE cc.cta_saldo END),0) " + _
                "from ctacteprv cc where cc.cta_fecha < '" + Now.AddDays(1).ToString("yyyyMMdd") + "' and cta_saldo<>0 "

            Dim deuda As Decimal = Math.Round(db.ExecuteQuery(Of Decimal)(lcSql).First() / dolar, 2)

            AuxMen = AuxMen + "<table border=""1""><tr><th>Proveedores</th><th>Total USD</th></tr>"
            AuxMen = AuxMen + "<tr><td>Deuda sin documentar</td><td align=""right"">" + deuda.ToString("N") + "</td></tr>"


            '' CHEQUES EMITIDOS
            lcSql = _
             "SELECT SUM(che_import) as total " + _
             " FROM chequesprop c " + _
             " WHERE che_acredi=0 and che_rechaz=0 and c.che_canjeado=0	"

            Dim cheques As Decimal = Math.Round(db.ExecuteQuery(Of Decimal)(lcSql).First() / dolar, 2)

            AuxMen = AuxMen + "<tr><td>Cheques emitidos</td><td align=""right"">" + cheques.ToString("N") + "</td></tr>"

            '' totales
            AuxMen = AuxMen + "<tr><td><b>Total USD</b></td><td align=""right""><b>" + (cheques + deuda).ToString("N") + "</b></td></tr>"
            AuxMen = AuxMen + "</table><br>"

            id_total = id_total - cheques - deuda

            Return AuxMen
        End Function

        Public Shared Function GetSaldoBanco(ByVal corte As Date, ByVal banco As Integer, ByVal pla_codigo As Integer) As Decimal
            Dim lcsql As String = _
            "select SUM(total) as total from ( "

            '' --***************** BUSCO TRANSFERENCIAS EN OP EFECTUADAS" 
            lcsql = lcsql + " select isnull(sum(-tra_total),0) as total" + _
             "     from pagos p, pagos_tra t" + _
             "     where p.pag_anulado = 0" + _
             "     and p.pag_tottra > 0" + _
             "     and p.pag_codigo=t.pag_codigo " + _
             "     and t.tra_fecha >= @desde and t.tra_fecha < @hasta" + _
             "     and t.cba_codigo = " + banco.ToString() + _
                " union"

            '' --**** CHEQUES EMITIDOS" 
            lcsql = lcsql + "     select isnull(sum(-che_import),0) as total" + _
             "     FROM chequesprop t " + _
             "     WHERE t.che_acredi=1  " + _
             "     and t.che_acrfec >= @desde and t.che_acrfec < @hasta" + _
             "     and t.cba_codigo = " + banco.ToString() + _
                " union"

            '' --**** CHEQUES DEPOSITADOS " 
            lcsql = lcsql + "     select isnull(sum(che_import),0) as total " + _
             "     FROM cheques t  " + _
             "     WHERE t.che_deposi=1 and t.che_rechaz=0   " + _
             "     and t.che_fecdep >= @desde and t.che_fecdep < @hasta " + _
             "     and t.cba_codigo = " + banco.ToString() + _
                " union "

            '' --**** TRANSFERENCIAS EN RECIBOS " 
            lcsql = lcsql + " select isnull(sum(t.tra_total),0) as total " + _
             "     from recibos r, recibos_tra t " + _
             "     where r.rec_anulad = 0 " + _
             "     and r.rec_tottra > 0 " + _
             "     and r.rec_codigo=t.rec_codigo   " + _
             "     and t.tra_fecha >= @desde and t.tra_fecha < @hasta " + _
             "     and t.cba_codigo = " + banco.ToString() + _
                " union "

            '' --**** CANJE DE CHEQUES DE TERCEROS POR DEPOSITO EN BANCO " 
            lcsql = lcsql + " select isnull(sum(che_import),0) as total " + _
             "     FROM cheques t  " + _
             "     WHERE t.che_canjeado=1 and t.che_canjtipo=2 " + _
             "     and t.che_canjfec >= @desde and t.che_canjfec < @hasta " + _
             "     and t.che_canjctadep = " + banco.ToString() + _
                " union "

            '' --**** ASIENTOS MANUALES Y TRANSFERENCIAS " 
            lcsql = lcsql + " select isnull(sum(case when r.asi_tipo='D' then r.asi_total else -r.asi_total end),0) AS total " + _
             "     from AsientosCab a, asientosren r  " + _
             "     where (a.asi_automatico = 0 or a.asi_modulo = 'TARJETASPAGOS') and a.asi_anulado = 0 " + _
             "     and a.asi_fecha >= @desde and a.asi_fecha < @hasta " + _
             "     and a.asi_codigo = r.asi_codigo  " + _
             "     and r.pla_codigo = " + pla_codigo.ToString() + _
                " union "

            '' --******************* rendiciones otros " 
            lcsql = lcsql + " select isnull(sum(case when a.reo_tipo='E' then -a.reo_total else a.reo_total end),0) as total " + _
             "     from rendicionesbancootros a " + _
             "     where a.reo_fecha >= @desde and a.reo_fecha < @hasta " + _
             "     and a.cba_codigo = " + banco.ToString() + _
                " union "

            '' --******************* cheques LIQUIDADOS RECHAZADOS " 
            lcsql = lcsql + " select isnull(sum(-che_import),0) as total  " + _
                  " FROM cheques t , cheques_liq l  " + _
                    " WHERE t.liq_codigo > 0 And t.che_rechaz = 1  " + _
                  " and t.che_recfec >= @desde and t.che_recfec < @hasta  " + _
                  " and t.liq_codigo = l.liq_codigo  " + _
                  " and l.cba_codigo =  12  " + _
                " union "

            '' --******************* liquidaciones de cheques 
            lcsql = lcsql + " select isnull(sum(l.liq_totliq),0) as total " + _
             "     from cheques_liq l " + _
             "     where l.liq_feccon >= @desde and l.liq_feccon < @hasta and l.liq_finalizado = 1 " + _
             "     and l.cba_codigo = " + banco.ToString() + _
                " ) a "

            lcsql = lcsql.Replace("@desde", "'" + corte.AddDays(1).ToString("yyyyMMdd") + "'")
            lcsql = lcsql.Replace("@hasta", "'" + Now.AddDays(1).ToString("yyyyMMdd") + "'")

            Dim db As New Dao

            Return db.ExecuteQuery(Of Decimal)(lcsql).First()

        End Function

        Public Shared Function GetSaldoCaja(ByVal corte As Date, ByVal pla_codigo As Integer) As Decimal
            Dim lcsql As String = _
            "select SUM(total) as total from ( "

            ''--******************* recibos del dia
            lcsql = lcsql + " SELECT ISNULL(sum(rec_totefe),0) as total " + _
            "     FROM RECIBOS R WHERE (r.rec_fecha >= @desde AND r.rec_Fecha<@hasta )  " + _
            "     AND r.rec_anulad = 0 " + _
            "     AND r.rec_totefe > 0 " + _
            " union "
            ''--******************* canjes en efectivo del dia
            lcsql = lcsql + " select isnull(sum(che_canjImporte),0) as total " + _
            "     from cheques ch where che_canjfec >= @desde and che_canjfec < @hasta and che_canjtipo=1  " + _
            " union	 "
            ''--******************* pagos del dia
            lcsql = lcsql + " SELECT isnull(sum(-pag_totefe),0) as total FROM pagos R  " + _
            "     WHERE (r.pag_fecha >= @desde AND r.pag_Fecha<@hasta )  " + _
            "     AND r.pag_anulado = 0 AND r.pag_totefe > 0 " + _
            " union "
            ''--******************* ingresos a fondo fijo
            lcsql = lcsql + " select isnull(sum(-mov_total),0) as total " + _
            " from FFMovimientos where mov_tipo = 'I' and mov_fecha >= @desde and mov_fecha < @hasta and mov_activo = 1  " + _
            " union "
            ''--******************* rendiciones de FF
            lcsql = lcsql + " select isnull(sum(rnd_total),0) " + _
            " from FFRendiciones where rnd_activo=1 and rnd_autori=1 " + _
            " and rnd_fecaut >= @desde and rnd_fecaut < @hasta " + _
            " union "
            ''--******************* asientos manuales y transferencias
            lcsql = lcsql + "select isnull(sum(case when r.asi_tipo='D' then r.asi_total else -r.asi_total end),0) AS total " + _
            "     from AsientosCab a, asientosren r  " + _
            "     where a.asi_automatico = 0 and a.asi_anulado = 0 " + _
            "     and a.asi_fecha >= @desde and a.asi_fecha < @hasta " + _
            "     and a.asi_codigo = r.asi_codigo  " + _
            "     and r.pla_codigo = 1871 " + _
            " ) a "

            lcsql = lcsql.Replace("@desde", "'" + corte.AddDays(1).ToString("yyyyMMdd") + "'")
            lcsql = lcsql.Replace("@hasta", "'" + Now.AddDays(1).ToString("yyyyMMdd") + "'")

            Dim db As New Dao

            Return db.ExecuteQuery(Of Decimal)(lcsql).First()

        End Function

    End Class


End Namespace