Module Job

#Region "Método principal"

    Public Sub Main()
        Dim iCountExeJOB As Integer = 0
        Dim oFiles() As String = Nothing

        'Comprobamos si el JOB está en ejecución y en caso afirmativo no lanzamos ningún proceso del JOB.
        For Each oProcess As Process In Process.GetProcesses()
            If oProcess.ProcessName.ToString = "EXO_UNI_INTERCOMPANY" Then
                iCountExeJOB += 1
            End If
        Next

        If iCountExeJOB = 0 Then
            If System.IO.Directory.Exists(My.Application.Info.DirectoryPath.ToString & "\Logs") Then
                oFiles = System.IO.Directory.GetFileSystemEntries(My.Application.Info.DirectoryPath.ToString & "\Logs")

                For Each sFile As String In oFiles
                    System.IO.File.Delete(System.IO.Path.GetFullPath(sFile))
                Next
            End If

            ' Responsables ICs
            Procesos.OAGP()

            'Formatos de fichero (Vías de pago)
            Procesos.OFRM()

            'Monedas
            Procesos.OCRN()

            'Dimensiones
            Procesos.ODIM()

            'Tipos de centro de coste
            Procesos.OCCT()

            ''Proyectos (Cuentas contables e ICs)
            'Procesos.OPRJ()

            'Categorías de balance
            Procesos.OACG()

            'Prioridades ICs
            Procesos.OBPP()

            'Dtos. por pronto pago (Condiciones de pago)
            Procesos.OCDC()

            'Propiedades ICs
            Procesos.OCQG()

            'Grupos de ICs
            Procesos.OCRG()

            'Grupos de E-mail (Personas de contacto ICs)
            Procesos.OEGP()

            'Idiomas
            Procesos.OLNG()

            'Formatos de dirección (Países)
            Procesos.OADF()

            'Países
            Procesos.OCRY()

            'Bloqueos de pago (ICs)
            Procesos.OPYB()

            'Territorios (ICs)
            Procesos.OTER()

            'Clases de expedición (ICs)
            Procesos.OSHP()

            'Indicadores de factoring (ICs)
            Procesos.OIDC()

            'Ramos (ICs)
            Procesos.OOND()

            'Estados (Direcciones ICs)
            Procesos.OCST()

            'Centros de coste
            Procesos.OPRC()

            'Bancos
            Procesos.ODSC()

            'Condiciones de pago
            Procesos.OCTG()

            'Vías de pago
            Procesos.OPYM()

            'Normas de reparto
            Procesos.OOCR()

            'Cuentas contables
            Procesos.OACT()

            'Modelos financieros
            Procesos.MODFINAN()

            'Tarjetas de crédito (ICs)
            Procesos.OCRC()

            'ICs
            Procesos.OCRD()

            'Facturas Ventas 
            'Procesos.OINV()
        End If
    End Sub

#End Region

End Module
