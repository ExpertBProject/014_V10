Public Class EXO_GLOBALES

    ''' <summary>
    ''' Función que devuelve si la empresa conectada es empresa Matriz
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function EmpresaConectadaEsMatriz(ByRef oObjGlobal As EXO_Generales.EXO_General) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        EmpresaConectadaEsMatriz = False

        Try
            oRs = CType(oObjGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs.DoQuery("SELECT t1.CompnyName " & _
                        "FROM OADM t1 WITH (NOLOCK) " & _
                        "WHERE ISNULL(t1.U_EXO_MATRIZ, 'N') = 'Y'")

            If oRs.RecordCount > 0 Then
                EmpresaConectadaEsMatriz = True
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

    ''' <summary>
    ''' Función que devuelve si la empresa conectada es empresa de Consolidación
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function EmpresaConectadaEsConsolidacion(ByRef oObjGlobal As EXO_Generales.EXO_General) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        EmpresaConectadaEsConsolidacion = False

        Try
            oRs = CType(oObjGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs.DoQuery("SELECT t1.CompnyName " & _
                        "FROM OADM t1 WITH (NOLOCK) " & _
                        "WHERE ISNULL(t1.U_EXO_CONSOLIDACION, 'N') = 'Y'")

            If oRs.RecordCount > 0 Then
                EmpresaConectadaEsConsolidacion = True
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

    ''' <summary>
    ''' Función que devuelve si la empresa conectada es empresa Sucursal
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function EmpresaConectadaEsSucursal(ByRef oObjGlobal As EXO_Generales.EXO_General) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        EmpresaConectadaEsSucursal = False

        Try
            oRs = CType(oObjGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs.DoQuery("SELECT t1.CompnyName " & _
                        "FROM OADM t1 WITH (NOLOCK) " & _
                        "WHERE ISNULL(t1.U_EXO_MATRIZ, 'N') = 'N' " & _
                        "AND ISNULL(t1.U_EXO_CONSOLIDACION, 'N') = 'N'")

            If oRs.RecordCount > 0 Then
                EmpresaConectadaEsSucursal = True
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

    Public Shared Sub Connect_Company(ByRef oObjGlobal As EXO_Generales.EXO_General, ByRef oCompany As SAPbobsCOM.Company, ByVal sDatabaseName As String)
        Try
            'Conectar DI SAP
            oCompany = New SAPbobsCOM.Company

            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish
            oCompany.Server = oObjGlobal.conexionSAP.compañia.Server
            'oCompany.LicenseServer = oObjGlobal.conexionSAP.compañia.LicenseServer 'Esto no funciona con el PL08 de la 9.2
            oCompany.LicenseServer = oObjGlobal.conexionSAP.compañia.Server
            oCompany.UserName = oObjGlobal.conexionSAP.compañia.UserName
            oCompany.Password = "159357"
            oCompany.UseTrusted = False
            oCompany.DbPassword = oObjGlobal.conexionSAP.refCompañia.OGEN.claveSQL
            oCompany.DbUserName = oObjGlobal.conexionSAP.compañia.DbUserName
            oCompany.DbServerType = oObjGlobal.conexionSAP.compañia.DbServerType
            oCompany.CompanyDB = sDatabaseName

            If oCompany.Connect <> 0 Then
                Throw New System.Exception("Error en la conexión a la compañia:" & oCompany.GetLastErrorDescription.Trim)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Sub Disconnect_Company(ByRef oCompany As SAPbobsCOM.Company)
        Try
            If Not oCompany Is Nothing Then
                If oCompany.Connected = True Then
                    oCompany.Disconnect()
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oCompany IsNot Nothing Then EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompany, Object))
            oCompany = Nothing
        End Try
    End Sub

End Class
