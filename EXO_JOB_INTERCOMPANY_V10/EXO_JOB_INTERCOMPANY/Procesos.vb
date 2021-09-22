Imports System.Data.SqlClient

Public Class Procesos

    Public Shared Sub OAGP()
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OAGP.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OAGP' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)


            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        sSQL = ""

                        If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OAGP]", "AgentCode", "AgentCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                            'Añadir
                            sSQL = "INSERT INTO [" & sDBD & "].dbo.[OAGP] " & _
                                   "SELECT [AgentCode], [AgentName], [Memo], [Locked], [DataSource], [UserSign] " & _
                                   "FROM [" & sDBO & "].dbo.[OAGP] t0 WITH (NOLOCK) " & _
                                   "WHERE t0.[AgentCode] = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'; "
                        Else
                            'Modificar"
                            sSQL = "UPDATE t1 SET [AgentName] = t0.[AgentName], " & _
                                   "[Memo] = t0.[Memo], " & _
                                   "[Locked] = t0.[Locked], " & _
                                   "[DataSource] = t0.[DataSource], " & _
                                   "[UserSign] = t0.[UserSign] " & _
                                   "FROM [" & sDBO & "].dbo.[OAGP] t0 WITH (NOLOCK) INNER JOIN " & _
                                   "[" & sDBD & "].dbo.[OAGP] t1 WITH (NOLOCK) ON t0.[AgentCode] = t1.[AgentCode] " & _
                                   "WHERE t0.[AgentCode] = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'; "
                        End If

                        sSQL &= "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()

            Conexiones.Disconnect_SQLServer(oDB)
        End Try
    End Sub

    Public Shared Sub OFRM()
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim oTransaction As SqlTransaction = Nothing

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OFRM.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t1.codeTable2 " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OFRM' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)


            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oTransaction = oDB.BeginTransaction("OFRM")

                        sSQL = ""

                        If Conexiones.GetValueDB(oDB, oTransaction, "[" & sDBD & "].dbo.[OFRM]", "AbsEntry", "Name = '" & oDt.Rows.Item(i).Item("codeTable2").ToString & "'") = "" Then
                            'Añadir
                            sSQL = "INSERT INTO [" & sDBD & "].dbo.[OFRM] " & _
                                   "SELECT (SELECT t1.AutoKey FROM [" & sDBD & "].dbo.[ONNM] t1 WITH (NOLOCK) WHERE t1.ObjectCode = '183'), t0.[Name], t0.[Encoding], t0.[FilePath], t0.[IsSystem], t0.[FrmatType], t0.[FileContnt], t0.[FrmatStats], t0.[PaymType] " & _
                                   "FROM [" & sDBO & "].dbo.[OFRM] t0 WITH (NOLOCK) " & _
                                   "WHERE t0.[AbsEntry] = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'; "

                            sSQL &= "UPDATE [" & sDBD & "].dbo.[ONNM] SET AutoKey = AutoKey + 1 WHERE ObjectCode = '183'; "
                        Else
                            'Modificar"
                            sSQL = "UPDATE t1 SET [Name] = t0.[Name], " & _
                                   "[Encoding] = t0.[Encoding], " & _
                                   "[FilePath] = t0.[FilePath], " & _
                                   "[IsSystem] = t0.[IsSystem], " & _
                                   "[FrmatType] = t0.[FrmatType], " & _
                                   "[FileContnt] = t0.[FileContnt], " & _
                                   "[FrmatStats] = t0.[FrmatStats], " & _
                                   "[PaymType] = t0.[PaymType] " & _
                                   "FROM [" & sDBO & "].dbo.[OFRM] t0 WITH (NOLOCK) INNER JOIN " & _
                                   "[" & sDBD & "].dbo.[OFRM] t1 WITH (NOLOCK) ON t0.[Name] = t1.[Name] " & _
                                   "WHERE t0.[AbsEntry] = " & oDt.Rows.Item(i).Item("codeTable").ToString & "; "
                        End If

                        sSQL &= "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, oTransaction, sSQL)

                        If oTransaction IsNot Nothing Then oTransaction.Commit()

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oTransaction IsNot Nothing Then oTransaction.Rollback()
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oTransaction IsNot Nothing Then oTransaction.Rollback()
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

            If oTransaction IsNot Nothing Then oTransaction.Rollback()
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)

            If oTransaction IsNot Nothing Then oTransaction.Rollback()
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()

            Conexiones.Disconnect_SQLServer(oDB)
        End Try
    End Sub

    Public Shared Sub OCRN()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOCRN As SAPbobsCOM.Currencies = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim oDecimals As SAPbobsCOM.CurrenciesDecimalsEnum = Nothing
        Dim sDocCurrCod As String = ""
        Dim sF100Name As String = ""
        Dim sFrgnName As String = ""
        Dim sChk100Name As String = ""
        Dim sChkName As String = ""
        Dim dMaxInDiff As Double = 0
        Dim dMaxInPcnt As Double = 0
        Dim dMaxOutDiff As Double = 0
        Dim dMaxOutPcnt As Double = 0
        Dim sCurrName As String = ""
        Dim sF100NamePl As String = ""
        Dim sFrgnNamePl As String = ""
        Dim sChk100NPl As String = ""
        Dim sChkNamePl As String = ""
        Dim oRoundSys As SAPbobsCOM.RoundingSysEnum = Nothing
        Dim oRoundPym As SAPbobsCOM.BoYesNoEnum = Nothing

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OCRN.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OCRN' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOCRN = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCurrencyCodes), SAPbobsCOM.Currencies)

                        If oOCRN.GetByKey(oDt.Rows.Item(i).Item("codeTable").ToString) = True Then
                            sXML = oOCRN.GetAsXML
                        Else
                            sXML = ""
                        End If

                        'Porque en el modo Update no funciona por XML
                        oDecimals = oOCRN.Decimals
                        sDocCurrCod = oOCRN.DocumentsCode
                        sF100Name = oOCRN.EnglishHundredthName
                        sFrgnName = oOCRN.EnglishName
                        sChk100Name = oOCRN.HundredthName
                        sChkName = oOCRN.InternationalDescription
                        dMaxInDiff = oOCRN.MaxIncomingAmtDiff
                        dMaxInPcnt = oOCRN.MaxIncomingAmtDiffPercent
                        dMaxOutDiff = oOCRN.MaxOutgoingAmtDiff
                        dMaxOutPcnt = oOCRN.MaxOutgoingAmtDiffPercent
                        sCurrName = oOCRN.Name
                        sF100NamePl = oOCRN.PluralEnglishHundredthName
                        sFrgnNamePl = oOCRN.PluralEnglishName
                        sChk100NPl = oOCRN.PluralHundredthName
                        sChkNamePl = oOCRN.PluralInternationalDescription
                        oRoundSys = oOCRN.Rounding
                        oRoundPym = oOCRN.RoundingInPayment
                        '''''''''''''''''''''''''''''''''''''''''''''

                        If sXML <> "" Then
                            oOCRN = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCurrencyCodes), SAPbobsCOM.Currencies)

                            oOCRN = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCRN]", "CurrCode", "CurrCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                                'Añadir
                                If oOCRN.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oOCRN.GetByKey(oDt.Rows.Item(i).Item("codeTable").ToString) = True Then
                                    oOCRN.Decimals = oDecimals
                                    oOCRN.DocumentsCode = sDocCurrCod
                                    oOCRN.EnglishHundredthName = sF100Name
                                    oOCRN.EnglishName = sFrgnName
                                    oOCRN.HundredthName = sChk100Name
                                    oOCRN.InternationalDescription = sChkName
                                    oOCRN.MaxIncomingAmtDiff = dMaxInDiff
                                    oOCRN.MaxIncomingAmtDiffPercent = dMaxInPcnt
                                    oOCRN.MaxOutgoingAmtDiff = dMaxOutDiff
                                    oOCRN.MaxOutgoingAmtDiffPercent = dMaxOutPcnt
                                    oOCRN.Name = sCurrName
                                    oOCRN.PluralEnglishHundredthName = sF100NamePl
                                    oOCRN.PluralEnglishName = sFrgnNamePl
                                    oOCRN.PluralHundredthName = sChk100NPl
                                    oOCRN.PluralInternationalDescription = sChkNamePl
                                    oOCRN.Rounding = oRoundSys
                                    oOCRN.RoundingInPayment = oRoundPym

                                    If oOCRN.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOCRN.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)


                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOCRN IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOCRN)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub ODIM()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing

        Dim oCmpSrvO As SAPbobsCOM.CompanyService = Nothing
        Dim oDIMServiceO As Object = Nothing
        Dim oDIMParamsO As Object = Nothing

        Dim oCmpSrvD As SAPbobsCOM.CompanyService = Nothing
        Dim oDIMServiceD As Object = Nothing
        Dim oDIMParamsD As Object = Nothing

        Dim oODIM As SAPbobsCOM.Dimension = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_ODIM.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'ODIM' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                oCmpSrvO = oCompanyO.GetCompanyService()
                oDIMServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.DimensionsService)
                oDIMParamsO = oDIMServiceO.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)

                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                oCmpSrvD = oCompanyD.GetCompanyService()
                oDIMServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.DimensionsService)
                oDIMParamsD = oDIMServiceD.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)
                            oCmpSrvO = oCompanyO.GetCompanyService()
                            oDIMServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.DimensionsService)
                            oDIMParamsO = oDIMServiceO.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)
                            oCmpSrvD = oCompanyD.GetCompanyService()
                            oDIMServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.DimensionsService)
                            oDIMParamsD = oDIMServiceD.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oDIMParamsO.DimensionCode = oDt.Rows.Item(i).Item("codeTable").ToString
                        oODIM = oDIMServiceO.GetDimension(oDIMParamsO)

                        sXML = oODIM.ToXMLString

                        If sXML <> "" Then
                            oDIMParamsD.DimensionCode = oDt.Rows.Item(i).Item("codeTable").ToString
                            oODIM = oDIMServiceD.GetDimension(oDIMParamsD)

                            oODIM.FromXMLString(sXML)

                            oDIMServiceD.UpdateDimension(oODIM)
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oDIMParamsO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oDIMParamsO)
            If oDIMServiceO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oDIMServiceO)
            If oCmpSrvO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvO)
            If oDIMParamsD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oDIMParamsD)
            If oDIMServiceD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oDIMServiceD)
            If oCmpSrvD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvD)
            If oODIM IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oODIM)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OCCT()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing

        Dim oCmpSrvO As SAPbobsCOM.CompanyService = Nothing
        Dim oCCTServiceO As Object = Nothing
        Dim oCCTParamsO As Object = Nothing

        Dim oCmpSrvD As SAPbobsCOM.CompanyService = Nothing
        Dim oCCTServiceD As Object = Nothing
        Dim oCCTParamsD As Object = Nothing

        Dim oOCCT As SAPbobsCOM.CostCenterType = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OCCT.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OCCT' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                oCmpSrvO = oCompanyO.GetCompanyService()
                oCCTServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.CostCenterTypesService)
                oCCTParamsO = oCCTServiceO.GetDataInterface(SAPbobsCOM.CostCenterTypesServiceDataInterfaces.cctsCostCenterTypeParams)

                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                oCmpSrvD = oCompanyD.GetCompanyService()
                oCCTServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.CostCenterTypesService)
                oCCTParamsD = oCCTServiceD.GetDataInterface(SAPbobsCOM.CostCenterTypesServiceDataInterfaces.cctsCostCenterTypeParams)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)
                            oCmpSrvO = oCompanyO.GetCompanyService()
                            oCCTServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.CostCenterTypesService)
                            oCCTParamsO = oCCTServiceO.GetDataInterface(SAPbobsCOM.CostCenterTypesServiceDataInterfaces.cctsCostCenterTypeParams)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)
                            oCmpSrvD = oCompanyD.GetCompanyService()
                            oCCTServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.CostCenterTypesService)
                            oCCTParamsD = oCCTServiceD.GetDataInterface(SAPbobsCOM.CostCenterTypesServiceDataInterfaces.cctsCostCenterTypeParams)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCCTParamsO.CostCenterTypeCode = oDt.Rows.Item(i).Item("codeTable").ToString
                        oOCCT = oCCTServiceO.GetCostCenterType(oCCTParamsO)

                        sXML = oOCCT.ToXMLString

                        If sXML <> "" Then
                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCCT]", "CctCode", "CctCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                                'Añadir
                                oOCCT = CType(oCCTServiceD.GetDataInterface(SAPbobsCOM.CostCenterTypesServiceDataInterfaces.cctsCostCenterType), SAPbobsCOM.CostCenterType)

                                oOCCT.FromXMLString(sXML)

                                oCCTServiceD.AddCostCenterType(oOCCT)
                            Else
                                'Modificar"
                                oCCTParamsD.CostCenterTypeCode = oDt.Rows.Item(i).Item("codeTable").ToString
                                oOCCT = oCCTServiceD.GetCostCenterType(oCCTParamsD)

                                oOCCT.FromXMLString(sXML)

                                oCCTServiceD.UpdateCostCenterType(oOCCT)
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oCCTParamsO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCCTParamsO)
            If oCCTServiceO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCCTServiceO)
            If oCmpSrvO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvO)
            If oCCTParamsD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCCTParamsD)
            If oCCTServiceD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCCTServiceD)
            If oCmpSrvD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvD)
            If oOCCT IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOCCT)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OPRJ()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing

        Dim oCmpSrvO As SAPbobsCOM.CompanyService = Nothing
        Dim oPRJServiceO As Object = Nothing
        Dim oPRJParamsO As Object = Nothing

        Dim oCmpSrvD As SAPbobsCOM.CompanyService = Nothing
        Dim oPRJServiceD As Object = Nothing
        Dim oPRJParamsD As Object = Nothing

        Dim oOPRJ As SAPbobsCOM.Project = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OPRJ.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OPRJ' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                oCmpSrvO = oCompanyO.GetCompanyService()
                oPRJServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
                oPRJParamsO = oPRJServiceO.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProjectParams)

                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                oCmpSrvD = oCompanyD.GetCompanyService()
                oPRJServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
                oPRJParamsD = oPRJServiceD.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProjectParams)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)
                            oCmpSrvO = oCompanyO.GetCompanyService()
                            oPRJServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
                            oPRJParamsO = oPRJServiceO.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProjectParams)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)
                            oCmpSrvD = oCompanyD.GetCompanyService()
                            oPRJServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
                            oPRJParamsD = oPRJServiceD.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProjectParams)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oPRJParamsO.Code = oDt.Rows.Item(i).Item("codeTable").ToString
                        oOPRJ = oPRJServiceO.GetProject(oPRJParamsO)

                        sXML = oOPRJ.ToXMLString

                        If sXML <> "" Then
                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OPRJ]", "PrjCode", "PrjCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                                'Añadir
                                oOPRJ = CType(oPRJServiceD.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProject), SAPbobsCOM.Project)

                                oOPRJ.FromXMLString(sXML)
                                oPRJServiceD.AddProject(oOPRJ)
                            Else
                                'Modificar"
                                oPRJParamsD.Code = oDt.Rows.Item(i).Item("codeTable").ToString
                                oOPRJ = oPRJServiceD.GetProject(oPRJParamsD)

                                oOPRJ.FromXMLString(sXML)

                                oPRJServiceD.UpdateProject(oOPRJ)
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oPRJParamsO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPRJParamsO)
            If oPRJServiceO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPRJServiceO)
            If oCmpSrvO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvO)
            If oPRJParamsD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPRJParamsD)
            If oPRJServiceD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPRJServiceD)
            If oCmpSrvD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvD)
            If oOPRJ IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOPRJ)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OACG()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing

        Dim oCmpSrvO As SAPbobsCOM.CompanyService = Nothing
        Dim oACGServiceO As Object = Nothing
        Dim oACGParamsO As Object = Nothing

        Dim oCmpSrvD As SAPbobsCOM.CompanyService = Nothing
        Dim oACGServiceD As Object = Nothing
        Dim oACGParamsD As Object = Nothing

        Dim oOACG As SAPbobsCOM.AccountCategory = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sAbsId As String = ""
        Dim oTransaction As SqlTransaction = Nothing

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OACG.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t1.codeTable2, t1.codeTable3 " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OACG' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                oCmpSrvO = oCompanyO.GetCompanyService()
                oACGServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.AccountCategoryService)
                oACGParamsO = oACGServiceO.GetDataInterface(SAPbobsCOM.AccountCategoryServiceDataInterfaces.acsAccountCategoryParams)

                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                oCmpSrvD = oCompanyD.GetCompanyService()
                oACGServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.AccountCategoryService)
                oACGParamsD = oACGServiceD.GetDataInterface(SAPbobsCOM.AccountCategoryServiceDataInterfaces.acsAccountCategoryParams)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)
                            oCmpSrvO = oCompanyO.GetCompanyService()
                            oACGServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.AccountCategoryService)
                            oACGParamsO = oACGServiceO.GetDataInterface(SAPbobsCOM.AccountCategoryServiceDataInterfaces.acsAccountCategoryParams)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)
                            oCmpSrvD = oCompanyD.GetCompanyService()
                            oACGServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.AccountCategoryService)
                            oACGParamsD = oACGServiceD.GetDataInterface(SAPbobsCOM.AccountCategoryServiceDataInterfaces.acsAccountCategoryParams)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        'Este primer IF es necesario porque para el valor del campo Source 'O' da error el DI API
                        If oDt.Rows.Item(i).Item("codeTable3").ToString = "O" Then
                            'Por SQL
                            oTransaction = oDB.BeginTransaction("OACG")

                            sSQL = ""

                            If Conexiones.GetValueDB(oDB, oTransaction, "[" & sDBD & "].dbo.[OACG]", "AbsId", "Name = '" & oDt.Rows.Item(i).Item("codeTable2").ToString & "' AND Source = '" & oDt.Rows.Item(i).Item("codeTable3").ToString & "'") = "" Then
                                'Añadir
                                sSQL = "INSERT INTO [" & sDBD & "].dbo.[OACG] " & _
                                       "SELECT (SELECT t1.AutoKey FROM [" & sDBD & "].dbo.[ONNM] t1 WITH (NOLOCK) WHERE t1.ObjectCode = '238'), [Name], [Source], [Locked], [DateSource], [UserSign] " & _
                                       "FROM [" & sDBO & "].dbo.[OACG] t0 WITH (NOLOCK) " & _
                                       "WHERE t0.[AbsId] = " & oDt.Rows.Item(i).Item("codeTable").ToString & "; "

                                sSQL &= "UPDATE [" & sDBD & "].dbo.[ONNM] SET AutoKey = AutoKey + 1 WHERE ObjectCode = '238'; "
                            Else
                                'Modificar"
                                sSQL = "UPDATE t1 SET [Name] = t0.[Name], " & _
                                       "[Source] = t0.[Source], " & _
                                       "[Locked] = t0.[Locked], " & _
                                       "[DateSource] = t0.[DateSource], " & _
                                       "[UserSign] = t0.[UserSign] " & _
                                       "FROM [" & sDBO & "].dbo.[OACG] t0 WITH (NOLOCK) INNER JOIN " & _
                                       "[" & sDBD & "].dbo.[OACG] t1 WITH (NOLOCK) ON t0.[Name] = t1.[Name] AND " & _
                                       "t0.Source = t1.Source " & _
                                       "WHERE t0.[AbsId] = " & oDt.Rows.Item(i).Item("codeTable").ToString & "; "
                            End If

                            sSQL &= "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "';"

                            Conexiones.ExecuteSQLDB(oDB, oTransaction, sSQL)

                            If oTransaction IsNot Nothing Then oTransaction.Commit()
                        Else
                            'Por DI API
                            oACGParamsO.CategoryCode = oDt.Rows.Item(i).Item("codeTable").ToString
                            oOACG = oACGServiceO.GetCategory(oACGParamsO)

                            sXML = oOACG.ToXMLString

                            If sXML <> "" Then
                                sAbsId = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OACG]", "AbsId", "Name = '" & oDt.Rows.Item(i).Item("codeTable2").ToString & "' AND Source = '" & oDt.Rows.Item(i).Item("codeTable3").ToString & "'")

                                If sAbsId = "" Then
                                    'Añadir
                                    oOACG = CType(oACGServiceD.GetDataInterface(SAPbobsCOM.AccountCategoryServiceDataInterfaces.acsAccountCategory), SAPbobsCOM.AccountCategory)

                                    oOACG.FromXMLString(sXML)
                                    oACGServiceD.AddCategory(oOACG)
                                Else
                                    'Modificar"
                                    oACGParamsD.CategoryCode = sAbsId
                                    oOACG = oACGServiceD.GetCategory(oACGParamsD)

                                    oOACG.FromXMLString(sXML)

                                    oACGServiceD.UpdateCategory(oOACG)
                                End If
                            End If

                            sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                            Conexiones.ExecuteSQLDB(oDB, sSQL)
                        End If

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oTransaction IsNot Nothing Then oTransaction.Rollback()
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oTransaction IsNot Nothing Then oTransaction.Rollback()
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

            If oTransaction IsNot Nothing Then oTransaction.Rollback()
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)

            If oTransaction IsNot Nothing Then oTransaction.Rollback()
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oACGParamsO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oACGParamsO)
            If oACGServiceO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oACGServiceO)
            If oCmpSrvO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvO)
            If oACGParamsD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oACGParamsD)
            If oACGServiceD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oACGServiceD)
            If oCmpSrvD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvD)
            If oOACG IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOACG)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OBPP()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOBPP As SAPbobsCOM.BPPriorities = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sPriorityDescription As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OBPP.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OBPP' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOBPP = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBPPriorities), SAPbobsCOM.BPPriorities)

                        If oOBPP.GetByKey(CInt(oDt.Rows.Item(i).Item("codeTable").ToString)) = True Then
                            sXML = oOBPP.GetAsXML
                        Else
                            sXML = ""
                        End If

                        'Porque en el modo Update no funciona por XML
                        sPriorityDescription = oOBPP.PriorityDescription
                        '''''''''''''''''''''''''''''''''''''''''''''

                        If sXML <> "" Then
                            oOBPP = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBPPriorities), SAPbobsCOM.BPPriorities)

                            oOBPP = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OBPP]", "PrioCode", "PrioCode = " & oDt.Rows.Item(i).Item("codeTable").ToString & "") = "" Then
                                'Añadir
                                If oOBPP.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oOBPP.GetByKey(CInt(oDt.Rows.Item(i).Item("codeTable").ToString)) = True Then
                                    oOBPP.PriorityDescription = sPriorityDescription

                                    If oOBPP.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOBPP.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOBPP IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOBPP)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OCDC()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing

        Dim oCmpSrvO As SAPbobsCOM.CompanyService = Nothing
        Dim oCDCServiceO As Object = Nothing
        Dim oCDCParamsO As Object = Nothing

        Dim oCmpSrvD As SAPbobsCOM.CompanyService = Nothing
        Dim oCDCServiceD As Object = Nothing
        Dim oPRJParamsD As Object = Nothing

        Dim oOCDC As SAPbobsCOM.CashDiscount = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OCDC.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OCDC' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                oCmpSrvO = oCompanyO.GetCompanyService()
                oCDCServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.CashDiscountsService)
                oCDCParamsO = oCDCServiceO.GetDataInterface(SAPbobsCOM.CashDiscountsServiceDataInterfaces.cdsCashDiscountParams)

                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                oCmpSrvD = oCompanyD.GetCompanyService()
                oCDCServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.CashDiscountsService)
                oPRJParamsD = oCDCServiceD.GetDataInterface(SAPbobsCOM.CashDiscountsServiceDataInterfaces.cdsCashDiscountParams)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)
                            oCmpSrvO = oCompanyO.GetCompanyService()
                            oCDCServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.CashDiscountsService)
                            oCDCParamsO = oCDCServiceO.GetDataInterface(SAPbobsCOM.CashDiscountsServiceDataInterfaces.cdsCashDiscountParams)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)
                            oCmpSrvD = oCompanyD.GetCompanyService()
                            oCDCServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.CashDiscountsService)
                            oPRJParamsD = oCDCServiceD.GetDataInterface(SAPbobsCOM.CashDiscountsServiceDataInterfaces.cdsCashDiscountParams)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCDCParamsO.Code = oDt.Rows.Item(i).Item("codeTable").ToString
                        oOCDC = oCDCServiceO.GetCashDiscount(oCDCParamsO)

                        sXML = oOCDC.ToXMLString

                        If sXML <> "" Then
                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCDC]", "Code", "Code = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                                'Añadir
                                oOCDC = CType(oCDCServiceD.GetDataInterface(SAPbobsCOM.CashDiscountsServiceDataInterfaces.cdsCashDiscount), SAPbobsCOM.CashDiscount)

                                oOCDC.FromXMLString(sXML)
                                oCDCServiceD.AddCashDiscount(oOCDC)
                            Else
                                'Modificar"
                                oPRJParamsD.Code = oDt.Rows.Item(i).Item("codeTable").ToString
                                oOCDC = oCDCServiceD.GetCashDiscount(oPRJParamsD)

                                oOCDC.FromXMLString(sXML)

                                oCDCServiceD.UpdateCashDiscount(oOCDC)
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oCDCParamsO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCDCParamsO)
            If oCDCServiceO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCDCServiceO)
            If oCmpSrvO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvO)
            If oPRJParamsD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPRJParamsD)
            If oCDCServiceD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCDCServiceD)
            If oCmpSrvD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvD)
            If oOCDC IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOCDC)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OCQG()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing

        Dim oCmpSrvO As SAPbobsCOM.CompanyService = Nothing
        Dim oCQGServiceO As Object = Nothing
        Dim oCQGParamsO As Object = Nothing

        Dim oCmpSrvD As SAPbobsCOM.CompanyService = Nothing
        Dim oCQGServiceD As Object = Nothing
        Dim oCQGParamsD As Object = Nothing

        Dim oOCQG As SAPbobsCOM.BusinessPartnerProperty = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OCQG.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OCQG' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                oCmpSrvO = oCompanyO.GetCompanyService()
                oCQGServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.BusinessPartnerPropertiesService)
                oCQGParamsO = oCQGServiceO.GetDataInterface(SAPbobsCOM.BusinessPartnerPropertiesServiceDataInterfaces.bppsBusinessPartnerPropertyParams)

                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                oCmpSrvD = oCompanyD.GetCompanyService()
                oCQGServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.BusinessPartnerPropertiesService)
                oCQGParamsD = oCQGServiceD.GetDataInterface(SAPbobsCOM.BusinessPartnerPropertiesServiceDataInterfaces.bppsBusinessPartnerPropertyParams)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)
                            oCmpSrvO = oCompanyO.GetCompanyService()
                            oCQGServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.BusinessPartnerPropertiesService)
                            oCQGParamsO = oCQGServiceO.GetDataInterface(SAPbobsCOM.BusinessPartnerPropertiesServiceDataInterfaces.bppsBusinessPartnerPropertyParams)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)
                            oCmpSrvD = oCompanyD.GetCompanyService()
                            oCQGServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.BusinessPartnerPropertiesService)
                            oCQGParamsD = oCQGServiceD.GetDataInterface(SAPbobsCOM.BusinessPartnerPropertiesServiceDataInterfaces.bppsBusinessPartnerPropertyParams)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCQGParamsO.PropertyCode = oDt.Rows.Item(i).Item("codeTable").ToString
                        oOCQG = oCQGServiceO.GetBusinessPartnerProperty(oCQGParamsO)

                        sXML = oOCQG.ToXMLString

                        If sXML <> "" Then
                            oCQGParamsD.PropertyCode = oDt.Rows.Item(i).Item("codeTable").ToString
                            oOCQG = oCQGServiceD.GetBusinessPartnerProperty(oCQGParamsD)

                            oOCQG.FromXMLString(sXML)

                            oCQGServiceD.UpdateBusinessPartnerProperty(oOCQG)
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oCQGParamsO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCQGParamsO)
            If oCQGServiceO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCQGServiceO)
            If oCmpSrvO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvO)
            If oCQGParamsD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCQGParamsD)
            If oCQGServiceD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCQGServiceD)
            If oCmpSrvD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvD)
            If oOCQG IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOCQG)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OCRG()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOCRG As SAPbobsCOM.BusinessPartnerGroups = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sGroupCode As String = ""
        Dim oType As SAPbobsCOM.BoBusinessPartnerGroupTypes = Nothing

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OCRG.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t1.codeTable2 " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OCRG' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOCRG = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartnerGroups), SAPbobsCOM.BusinessPartnerGroups)

                        If oOCRG.GetByKey(CInt(oDt.Rows.Item(i).Item("codeTable").ToString)) = True Then
                            sXML = oOCRG.GetAsXML
                        Else
                            sXML = ""
                        End If

                        'Porque en el modo Update no funciona por XML
                        oType = oOCRG.Type
                        '''''''''''''''''''''''''''''''''''''''''''''

                        If sXML <> "" Then
                            oOCRG = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartnerGroups), SAPbobsCOM.BusinessPartnerGroups)

                            oOCRG = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            sGroupCode = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCRG]", "GroupCode", "GroupName = '" & oDt.Rows.Item(i).Item("codeTable2").ToString & "'")

                            If sGroupCode = "" Then
                                'Añadir
                                If oOCRG.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oOCRG.GetByKey(CInt(sGroupCode)) = True Then
                                    oOCRG.Type = oType

                                    If oOCRG.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOCRG.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOCRG IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOCRG)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OEGP()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing

        Dim oCmpSrvO As SAPbobsCOM.CompanyService = Nothing
        Dim oEGPServiceO As Object = Nothing
        Dim oEGPParamsO As Object = Nothing

        Dim oCmpSrvD As SAPbobsCOM.CompanyService = Nothing
        Dim oEGPServiceD As Object = Nothing
        Dim oEGPParamsD As Object = Nothing

        Dim oOEGP As SAPbobsCOM.EmailGroup = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OEGP.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OEGP' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                oCmpSrvO = oCompanyO.GetCompanyService()
                oEGPServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.EmailGroupsService)
                oEGPParamsO = oEGPServiceO.GetDataInterface(SAPbobsCOM.EmailGroupsServiceDataInterfaces.egsEmailGroupParams)

                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                oCmpSrvD = oCompanyD.GetCompanyService()
                oEGPServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.EmailGroupsService)
                oEGPParamsD = oEGPServiceD.GetDataInterface(SAPbobsCOM.EmailGroupsServiceDataInterfaces.egsEmailGroupParams)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)
                            oCmpSrvO = oCompanyO.GetCompanyService()
                            oEGPServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.EmailGroupsService)
                            oEGPParamsO = oEGPServiceO.GetDataInterface(SAPbobsCOM.EmailGroupsServiceDataInterfaces.egsEmailGroupParams)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)
                            oCmpSrvD = oCompanyD.GetCompanyService()
                            oEGPServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.EmailGroupsService)
                            oEGPParamsD = oEGPServiceD.GetDataInterface(SAPbobsCOM.EmailGroupsServiceDataInterfaces.egsEmailGroupParams)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oEGPParamsO.EmailGroupCode = oDt.Rows.Item(i).Item("codeTable").ToString
                        oOEGP = oEGPServiceO.Get(oEGPParamsO)

                        sXML = oOEGP.ToXMLString

                        If sXML <> "" Then
                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OEGP]", "EmlGrpCode", "EmlGrpCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                                'Añadir
                                oOEGP = CType(oEGPServiceD.GetDataInterface(SAPbobsCOM.EmailGroupsServiceDataInterfaces.egsEmailGroup), SAPbobsCOM.EmailGroup)

                                oOEGP.FromXMLString(sXML)
                                oEGPServiceD.Add(oOEGP)
                            Else
                                'Modificar"
                                oEGPParamsD.EmailGroupCode = oDt.Rows.Item(i).Item("codeTable").ToString
                                oOEGP = oEGPServiceD.Get(oEGPParamsD)

                                oOEGP.FromXMLString(sXML)

                                oEGPServiceD.Update(oOEGP)
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oEGPParamsO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oEGPParamsO)
            If oEGPServiceO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oEGPServiceO)
            If oCmpSrvO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvO)
            If oEGPParamsD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oEGPParamsD)
            If oEGPServiceD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oEGPServiceD)
            If oCmpSrvD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvD)
            If oOEGP IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOEGP)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OLNG()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOLNG As SAPbobsCOM.UserLanguages = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sName As String = ""
        Dim sSysLang As String = ""
        Dim sCode As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OLNG.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t1.codeTable2 " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OLNG' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOLNG = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserLanguages), SAPbobsCOM.UserLanguages)

                        If oOLNG.GetByKey(CInt(oDt.Rows.Item(i).Item("codeTable").ToString)) = True Then
                            sXML = oOLNG.GetAsXML
                        Else
                            sXML = ""
                        End If

                        'Porque en el modo Update no funciona por XML
                        sName = oOLNG.LanguageFullName
                        sSysLang = oOLNG.RelatedSystemLanguage
                        '''''''''''''''''''''''''''''''''''''''''''''

                        If sXML <> "" Then
                            oOLNG = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserLanguages), SAPbobsCOM.UserLanguages)

                            oOLNG = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            sCode = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OLNG]", "Code", "ShortName = '" & oDt.Rows.Item(i).Item("codeTable2").ToString & "'")

                            If sCode = "" Then
                                'Añadir
                                If oOLNG.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oOLNG.GetByKey(CInt(sCode)) = True Then
                                    oOLNG.LanguageFullName = sName
                                    oOLNG.RelatedSystemLanguage = CInt(sSysLang)

                                    If oOLNG.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOLNG.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOLNG IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOLNG)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OCRY()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing

        Dim oCmpSrvO As SAPbobsCOM.CompanyService = Nothing
        Dim oCRYServiceO As Object = Nothing
        Dim oCRYParamsO As Object = Nothing

        Dim oCmpSrvD As SAPbobsCOM.CompanyService = Nothing
        Dim oCRYServiceD As Object = Nothing
        Dim oCRYParamsD As Object = Nothing

        Dim oOCRY As SAPbobsCOM.Country = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sAddrFormat As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OCRY.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OCRY' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                oCmpSrvO = oCompanyO.GetCompanyService()
                oCRYServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.CountriesService)
                oCRYParamsO = oCRYServiceO.GetDataInterface(SAPbobsCOM.CountriesServiceDataInterfaces.csCountryParams)

                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                oCmpSrvD = oCompanyD.GetCompanyService()
                oCRYServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.CountriesService)
                oCRYParamsD = oCRYServiceD.GetDataInterface(SAPbobsCOM.CountriesServiceDataInterfaces.csCountryParams)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)
                            oCmpSrvO = oCompanyO.GetCompanyService()
                            oCRYServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.CountriesService)
                            oCRYParamsO = oCRYServiceO.GetDataInterface(SAPbobsCOM.CountriesServiceDataInterfaces.csCountryParams)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)
                            oCmpSrvD = oCompanyD.GetCompanyService()
                            oCRYServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.CountriesService)
                            oCRYParamsD = oCRYServiceD.GetDataInterface(SAPbobsCOM.CountriesServiceDataInterfaces.csCountryParams)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCRYParamsO.Code = oDt.Rows.Item(i).Item("codeTable").ToString
                        oOCRY = oCRYServiceO.GetCountry(oCRYParamsO)

                        sXML = oOCRY.ToXMLString

                        If sXML <> "" Then
                            'Esto es porque el código del formato de la dirección no tiene por qué ser igual en todas las empresas
                            sAddrFormat = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OADF]", "Name", "Code = " & oOCRY.AddressFormat & "")
                            sAddrFormat = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OADF]", "Code", "Name = '" & sAddrFormat & "'")

                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCRY]", "Code", "Code = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                                'Añadir
                                oOCRY = CType(oCRYServiceD.GetDataInterface(SAPbobsCOM.CountriesServiceDataInterfaces.csCountry), SAPbobsCOM.Country)

                                oOCRY.FromXMLString(sXML)

                                If sAddrFormat = "" Then
                                    oOCRY.AddressFormat = -1
                                Else
                                    oOCRY.AddressFormat = sAddrFormat
                                End If

                                oCRYServiceD.AddCountry(oOCRY)
                            Else
                                'Modificar"
                                oCRYParamsD.Code = oDt.Rows.Item(i).Item("codeTable").ToString
                                oOCRY = oCRYServiceD.GetCountry(oCRYParamsD)

                                oOCRY.FromXMLString(sXML)

                                If sAddrFormat = "" Then
                                    oOCRY.AddressFormat = -1
                                Else
                                    oOCRY.AddressFormat = sAddrFormat
                                End If

                                oCRYServiceD.UpdateCountry(oOCRY)
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oCRYParamsO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCRYParamsO)
            If oCRYServiceO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCRYServiceO)
            If oCmpSrvO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvO)
            If oCRYParamsD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCRYParamsD)
            If oCRYServiceD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCRYServiceD)
            If oCmpSrvD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvD)
            If oOCRY IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOCRY)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OADF()
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim oTransaction As SqlTransaction = Nothing

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OADF.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t1.codeTable2 " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OADF' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)


            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oTransaction = oDB.BeginTransaction("OADF")

                        sSQL = ""

                        If Conexiones.GetValueDB(oDB, oTransaction, "[" & sDBD & "].dbo.[OADF]", "Code", "Name = '" & oDt.Rows.Item(i).Item("codeTable2").ToString & "'") = "" Then
                            'Añadir
                            sSQL = "INSERT INTO [" & sDBD & "].dbo.[OADF] " & _
                                   "SELECT (SELECT t1.AutoKey FROM [" & sDBD & "].dbo.[ONNM] t1 WITH (NOLOCK) WHERE t1.ObjectCode = '131'), [Name], [Format], [UserSign] " & _
                                   "FROM [" & sDBO & "].dbo.[OADF] t0 WITH (NOLOCK) " & _
                                   "WHERE t0.[Code] = " & oDt.Rows.Item(i).Item("codeTable").ToString & "; "

                            sSQL &= "UPDATE [" & sDBD & "].dbo.[ONNM] SET AutoKey = AutoKey + 1 WHERE ObjectCode = '131'; "
                        Else
                            'Modificar"
                            sSQL = "UPDATE t1 SET [Format] = t0.[Format], " & _
                                   "[UserSign] = t0.[UserSign] " & _
                                   "FROM [" & sDBO & "].dbo.[OADF] t0 WITH (NOLOCK) INNER JOIN " & _
                                   "[" & sDBD & "].dbo.[OADF] t1 WITH (NOLOCK) ON t0.[Name] = t1.[Name] " & _
                                   "WHERE t0.[Code] = " & oDt.Rows.Item(i).Item("codeTable").ToString & "; "
                        End If

                        sSQL &= "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, oTransaction, sSQL)

                        If oTransaction IsNot Nothing Then oTransaction.Commit()

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oTransaction IsNot Nothing Then oTransaction.Rollback()
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oTransaction IsNot Nothing Then oTransaction.Rollback()
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

            If oTransaction IsNot Nothing Then oTransaction.Rollback()
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)

            If oTransaction IsNot Nothing Then oTransaction.Rollback()
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()

            Conexiones.Disconnect_SQLServer(oDB)
        End Try
    End Sub

    Public Shared Sub OPYB()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing

        Dim oCmpSrvO As SAPbobsCOM.CompanyService = Nothing
        Dim oPYBServiceO As Object = Nothing
        Dim oPYBParamsO As Object = Nothing

        Dim oCmpSrvD As SAPbobsCOM.CompanyService = Nothing
        Dim oPYBServiceD As Object = Nothing
        Dim oPYBParamsD As Object = Nothing

        Dim oOPYB As SAPbobsCOM.PaymentBlock = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sAbsEntry As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OPYB.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t1.codeTable2 " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OPYB' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                oCmpSrvO = oCompanyO.GetCompanyService()
                oPYBServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.PaymentBlocksService)
                oPYBParamsO = oPYBServiceO.GetDataInterface(SAPbobsCOM.PaymentBlocksServiceDataInterfaces.pbsPaymentBlockParams)

                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                oCmpSrvD = oCompanyD.GetCompanyService()
                oPYBServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.PaymentBlocksService)
                oPYBParamsD = oPYBServiceD.GetDataInterface(SAPbobsCOM.PaymentBlocksServiceDataInterfaces.pbsPaymentBlockParams)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)
                            oCmpSrvO = oCompanyO.GetCompanyService()
                            oPYBServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.PaymentBlocksService)
                            oPYBParamsO = oPYBServiceO.GetDataInterface(SAPbobsCOM.PaymentBlocksServiceDataInterfaces.pbsPaymentBlockParams)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)
                            oCmpSrvD = oCompanyD.GetCompanyService()
                            oPYBServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.PaymentBlocksService)
                            oPYBParamsD = oPYBServiceD.GetDataInterface(SAPbobsCOM.PaymentBlocksServiceDataInterfaces.pbsPaymentBlockParams)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oPYBParamsO.AbsEntry = oDt.Rows.Item(i).Item("codeTable").ToString
                        oOPYB = oPYBServiceO.GetPaymentBlock(oPYBParamsO)

                        sXML = oOPYB.ToXMLString

                        If sXML <> "" Then
                            sAbsEntry = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OPYB]", "AbsEntry", "PayBlock = '" & oDt.Rows.Item(i).Item("codeTable2").ToString & "'")

                            If sAbsEntry = "" Then
                                'Añadir
                                oOPYB = CType(oPYBServiceD.GetDataInterface(SAPbobsCOM.PaymentBlocksServiceDataInterfaces.pbsPaymentBlock), SAPbobsCOM.PaymentBlock)

                                oOPYB.FromXMLString(sXML)
                                oPYBServiceD.AddPaymentBlock(oOPYB)
                            Else
                                'Modificar"
                                oPYBParamsD.AbsEntry = sAbsEntry
                                oOPYB = oPYBServiceD.GetPaymentBlock(oPYBParamsD)

                                oOPYB.FromXMLString(sXML)

                                oPYBServiceD.UpdatePaymentBlock(oOPYB)
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oPYBParamsO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPYBParamsO)
            If oPYBServiceO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPYBServiceO)
            If oCmpSrvO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvO)
            If oPYBParamsD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPYBParamsD)
            If oPYBServiceD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPYBServiceD)
            If oCmpSrvD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvD)
            If oOPYB IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOPYB)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OTER()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing

        Dim oOTER As SAPbobsCOM.Territories = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sTerritryID As String = ""
        Dim sDescript As String = ""
        Dim sInactive As String = ""
        Dim sParent As String = ""
        Dim sAux As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OTER.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t1.codeTable2, t1.codeTable3 " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OTER' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        'Porque no se puede hacer por XML, ya que es una tabla recursiva y hay que comprobar más cosas
                        'oCompanyO.XMLAsString = True
                        'oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        'oCompanyD.XMLAsString = True
                        'oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        oOTER = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oTerritories), SAPbobsCOM.Territories)

                        If oOTER.GetByKey(CInt(oDt.Rows.Item(i).Item("codeTable").ToString)) = True Then
                            sXML = oOTER.GetAsXML
                        Else
                            sXML = ""
                        End If

                        'Porque no se puede hacer por XML, ya que es una tabla recursiva y hay que comprobar más cosas
                        sDescript = oOTER.Description
                        sInactive = oOTER.Inactive
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        If sXML <> "" Then
                            oOTER = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oTerritories), SAPbobsCOM.Territories)

                            'Porque no se puede hacer por XML, ya que es una tabla recursiva y hay que comprobar más cosas
                            'oOTER = oCompanyD.GetBusinessObjectFromXML(sXML, 0)
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                            sTerritryID = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OTER]", "territryID", "descript = '" & sDescript & "'")

                            sAux = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OTER]", "descript", "territryID = " & oDt.Rows.Item(i).Item("codeTable2").ToString & "")
                            sParent = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OTER]", "territryID", "descript = '" & sAux & "'")

                            If sTerritryID = "" Then
                                'Añadir
                                oOTER.Description = sDescript
                                oOTER.Inactive = sInactive

                                If sParent <> "" Then
                                    oOTER.Parent = CInt(sParent)
                                End If

                                If oOTER.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oOTER.GetByKey(CInt(sTerritryID)) = True Then
                                    oOTER.Description = sDescript
                                    oOTER.Inactive = sInactive

                                    If sParent <> "" Then
                                        oOTER.Parent = CInt(sParent)
                                    End If

                                    If oOTER.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOTER.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOTER IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOTER)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OSHP()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOSHP As SAPbobsCOM.ShippingTypes = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sTrnspCode As String = ""
        Dim sWebSite As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OSHP.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t1.codeTable2 " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OSHP' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOSHP = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oShippingTypes), SAPbobsCOM.ShippingTypes)

                        If oOSHP.GetByKey(CInt(oDt.Rows.Item(i).Item("codeTable").ToString)) = True Then
                            sXML = oOSHP.GetAsXML
                        Else
                            sXML = ""
                        End If

                        'Porque en el modo Update no funciona por XML
                        sWebSite = oOSHP.Website
                        '''''''''''''''''''''''''''''''''''''''''''''

                        If sXML <> "" Then
                            oOSHP = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oShippingTypes), SAPbobsCOM.ShippingTypes)

                            oOSHP = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            sTrnspCode = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OSHP]", "TrnspCode", "TrnspName = '" & oDt.Rows.Item(i).Item("codeTable2").ToString & "'")

                            If sTrnspCode = "" Then
                                'Añadir
                                If oOSHP.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oOSHP.GetByKey(CInt(sTrnspCode)) = True Then
                                    oOSHP.Website = sWebSite

                                    If oOSHP.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOSHP.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOSHP IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOSHP)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OIDC()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOIDC As SAPbobsCOM.FactoringIndicators = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sName As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OIDC.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OIDC' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOIDC = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFactoringIndicators), SAPbobsCOM.FactoringIndicators)

                        If oOIDC.GetByKey(oDt.Rows.Item(i).Item("codeTable").ToString) = True Then
                            sXML = oOIDC.GetAsXML
                        Else
                            sXML = ""
                        End If

                        'Porque en el modo Update no funciona por XML
                        sName = oOIDC.IndicatorName
                        '''''''''''''''''''''''''''''''''''''''''''''

                        If sXML <> "" Then
                            oOIDC = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFactoringIndicators), SAPbobsCOM.FactoringIndicators)

                            oOIDC = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OIDC]", "Name", "Code = " & oDt.Rows.Item(i).Item("codeTable").ToString & "") = "" Then
                                'Añadir
                                If oOIDC.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oOIDC.GetByKey(oDt.Rows.Item(i).Item("codeTable").ToString) = True Then
                                    oOIDC.IndicatorName = sName

                                    If oOIDC.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOIDC.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOIDC IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOIDC)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OOND()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOOND As SAPbobsCOM.Industries = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sIndCode As String = ""
        Dim sIndDesc As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OOND.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t1.codeTable2 " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OOND' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOOND = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIndustries), SAPbobsCOM.Industries)

                        If oOOND.GetByKey(CInt(oDt.Rows.Item(i).Item("codeTable").ToString)) = True Then
                            sXML = oOOND.GetAsXML
                        Else
                            sXML = ""
                        End If

                        'Porque en el modo Update no funciona por XML
                        sIndDesc = oOOND.IndustryDescription
                        '''''''''''''''''''''''''''''''''''''''''''''

                        If sXML <> "" Then
                            oOOND = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIndustries), SAPbobsCOM.Industries)

                            oOOND = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            sIndCode = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OOND]", "IndCode", "IndName = '" & oDt.Rows.Item(i).Item("codeTable2").ToString & "'")

                            If sIndCode = "" Then
                                'Añadir
                                If oOOND.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oOOND.GetByKey(CInt(sIndCode)) = True Then
                                    oOOND.IndustryDescription = sIndDesc

                                    If oOOND.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOOND.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOOND IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOOND)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OCST()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing

        Dim oCmpSrvO As SAPbobsCOM.CompanyService = Nothing
        Dim oCSTServiceO As Object = Nothing
        Dim oCSTParamsO As Object = Nothing

        Dim oCmpSrvD As SAPbobsCOM.CompanyService = Nothing
        Dim oCSTServiceD As Object = Nothing
        Dim oCSTParamsD As Object = Nothing

        Dim oOCST As SAPbobsCOM.State = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OCST.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OCST' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t2.Code " & _
                       "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) INNER JOIN " & _
                       "[" & sDBO & "].dbo.[OCST] t2 WITH (NOLOCK) ON t1.codeTable = t2.Country " & _
                       "WHERE t1.tableName = 'OCST' " & _
                       "ORDER BY t1.dbNameOrig, t1.dbNameDest "

                oDt = New System.Data.DataTable
                Conexiones.FillDtDB(oDB, oDt, sSQL)

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                oCmpSrvO = oCompanyO.GetCompanyService()
                oCSTServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.StatesService)
                oCSTParamsO = oCSTServiceO.GetDataInterface(SAPbobsCOM.StatesServiceDataInterfaces.ssStateParams)

                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                oCmpSrvD = oCompanyD.GetCompanyService()
                oCSTServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.StatesService)
                oCSTParamsD = oCSTServiceD.GetDataInterface(SAPbobsCOM.StatesServiceDataInterfaces.ssStateParams)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)
                            oCmpSrvO = oCompanyO.GetCompanyService()
                            oCSTServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.StatesService)
                            oCSTParamsO = oCSTServiceO.GetDataInterface(SAPbobsCOM.StatesServiceDataInterfaces.ssStateParams)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString

                            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t2.Code " & _
                                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) INNER JOIN " & _
                                   "[" & sDBO & "].dbo.[OCST] t2 WITH (NOLOCK) ON t1.codeTable = t2.Country " & _
                                   "WHERE t1.tableName = 'OCST' " & _
                                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

                            oDt = New System.Data.DataTable
                            Conexiones.FillDtDB(oDB, oDt, sSQL)

                            i = 0
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)
                            oCmpSrvD = oCompanyD.GetCompanyService()
                            oCSTServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.StatesService)
                            oCSTParamsD = oCSTServiceD.GetDataInterface(SAPbobsCOM.StatesServiceDataInterfaces.ssStateParams)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCSTParamsO.Code = oDt.Rows.Item(i).Item("Code").ToString
                        oCSTParamsO.Country = oDt.Rows.Item(i).Item("codeTable").ToString
                        oOCST = oCSTServiceO.GetState(oCSTParamsO)

                        sXML = oOCST.ToXMLString

                        If sXML <> "" Then
                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCST]", "Code", "Code = '" & oDt.Rows.Item(i).Item("Code").ToString & "' AND Country = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                                'Añadir
                                oOCST = CType(oCSTServiceD.GetDataInterface(SAPbobsCOM.StatesServiceDataInterfaces.ssState), SAPbobsCOM.State)

                                oOCST.FromXMLString(sXML)

                                oCSTServiceD.AddState(oOCST)
                            Else
                                'Modificar"
                                oCSTParamsD.Code = oDt.Rows.Item(i).Item("Code").ToString
                                oCSTParamsD.Country = oDt.Rows.Item(i).Item("codeTable").ToString
                                oOCST = oCSTServiceD.GetState(oCSTParamsD)

                                oOCST.FromXMLString(sXML)

                                oCSTServiceD.UpdateState(oOCST)
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oCSTParamsO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCSTParamsO)
            If oCSTServiceO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCSTServiceO)
            If oCmpSrvO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvO)
            If oCSTParamsD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCSTParamsD)
            If oCSTServiceD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCSTServiceD)
            If oCmpSrvD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvD)
            If oOCST IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOCST)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OPRC()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing

        Dim oCmpSrvO As SAPbobsCOM.CompanyService = Nothing
        Dim oPRCServiceO As Object = Nothing
        Dim oPRCParamsO As Object = Nothing

        Dim oCmpSrvD As SAPbobsCOM.CompanyService = Nothing
        Dim oPRCServiceD As Object = Nothing
        Dim oPRCParamsD As Object = Nothing

        Dim oOPRC As SAPbobsCOM.ProfitCenter = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OPRC.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OPRC' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                oCmpSrvO = oCompanyO.GetCompanyService()
                oPRCServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService)
                oPRCParamsO = oPRCServiceO.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenterParams)

                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                oCmpSrvD = oCompanyD.GetCompanyService()
                oPRCServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService)
                oPRCParamsD = oPRCServiceD.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenterParams)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)
                            oCmpSrvO = oCompanyO.GetCompanyService()
                            oPRCServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService)
                            oPRCParamsO = oPRCServiceO.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenterParams)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)
                            oCmpSrvD = oCompanyD.GetCompanyService()
                            oPRCServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService)
                            oPRCParamsD = oPRCServiceD.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenterParams)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oPRCParamsO.CenterCode = oDt.Rows.Item(i).Item("codeTable").ToString
                        oOPRC = oPRCServiceO.GetProfitCenter(oPRCParamsO)

                        sXML = oOPRC.ToXMLString

                        If sXML <> "" Then
                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OPRC]", "PrcCode", "PrcCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                                'Añadir
                                oOPRC = CType(oPRCServiceD.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenter), SAPbobsCOM.ProfitCenter)

                                oOPRC.FromXMLString(sXML)

                                oPRCServiceD.AddProfitCenter(oOPRC)
                            Else
                                'Modificar"
                                oPRCParamsD.CenterCode = oDt.Rows.Item(i).Item("codeTable").ToString
                                oOPRC = oPRCServiceD.GetProfitCenter(oPRCParamsD)

                                oOPRC.FromXMLString(sXML)

                                oPRCServiceD.UpdateProfitCenter(oOPRC)
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oPRCParamsO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPRCParamsO)
            If oPRCServiceO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPRCServiceO)
            If oCmpSrvO IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvO)
            If oPRCParamsD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPRCParamsD)
            If oPRCServiceD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oPRCServiceD)
            If oCmpSrvD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrvD)
            If oOPRC IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOPRC)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub ODSC()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oODSC As SAPbobsCOM.Banks = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim oXml As Xml.XmlDocument = Nothing
        Dim oXmlNode As Xml.XmlNode = Nothing
        Dim sAbsEntry As String = ""
        Dim sBankName As String = ""
        Dim sCountryCod As String = ""
        Dim sPostOffice As String = ""
        Dim sIBAN As String = ""
        Dim sSwiftNum As String = ""
        Dim oDtAux As System.Data.DataTable = Nothing

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_ODSC.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t1.codeTable2 " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'ODSC' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        oCompanyD.StartTransaction()

                        oRs = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oODSC = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBanks), SAPbobsCOM.Banks)

                        If oODSC.GetByKey(CInt(oDt.Rows.Item(i).Item("codeTable").ToString)) = True Then
                            sXML = oODSC.GetAsXML
                        Else
                            sXML = ""
                        End If

                        'Porque en el modo Update no funciona por XML
                        sBankName = oODSC.BankName
                        sCountryCod = oODSC.CountryCode
                        sPostOffice = oODSC.PostOffice
                        sIBAN = oODSC.IBAN
                        sSwiftNum = oODSC.SwiftNo
                        '''''''''''''''''''''''''''''''''''''''''''''

                        If sXML <> "" Then
                            Try
                                oXml = New Xml.XmlDocument
                                oXml.LoadXml(sXML)

                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/Banks/row/DefaultBankAccountKey")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)

                                sXML = oXml.OuterXml
                            Catch ex As Exception

                            End Try

                            oODSC = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBanks), SAPbobsCOM.Banks)

                            oODSC = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            sAbsEntry = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[ODSC]", "AbsEntry", "BankCode = '" & oDt.Rows.Item(i).Item("codeTable2").ToString & "'")

                            If sAbsEntry = "" Then
                                'Añadir
                                If oODSC.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oODSC.GetByKey(CInt(sAbsEntry)) = True Then
                                    oODSC.BankName = sBankName
                                    oODSC.CountryCode = sCountryCod
                                    oODSC.PostOffice = sPostOffice
                                    oODSC.IBAN = sIBAN
                                    oODSC.SwiftNo = sSwiftNum

                                    If oODSC.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If

                                    'Porque el DI API borra estos datos de la tabla ODSC al actualizar
                                    sSQL = "SELECT t1.Account, t1.Branch, ISNULL(t1.NextCheck, 0) NextCheck " & _
                                           "FROM [" & sDBD & "].dbo.[DSC1] t1 WITH (NOLOCK) " & _
                                           "WHERE t1.AbsEntry = " & sAbsEntry & " "

                                    oDtAux = New System.Data.DataTable
                                    Conexiones.FillDtDB(oDB, oDtAux, sSQL)

                                    If oDtAux.Rows.Count > 0 Then
                                        sSQL = "UPDATE [" & sDBD & "].dbo.[ODSC] SET DfltAcct = '" & oDtAux.Rows.Item(0).Item("Account").ToString & "', " & _
                                               "DfltBranch = '" & oDtAux.Rows.Item(0).Item("Branch").ToString & "', " & _
                                               "NextChckNo = " & oDtAux.Rows.Item(0).Item("NextCheck").ToString & " " & _
                                               "WHERE AbsEntry = " & sAbsEntry & "; "

                                        oRs.DoQuery(sSQL)
                                    End If
                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oODSC.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'; "

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

            If oCompanyD.InTransaction = True Then
                oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)

            If oCompanyD.InTransaction = True Then
                oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oDtAux IsNot Nothing Then oDtAux.Dispose()
            If oODSC IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oODSC)
            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRs)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OCTG()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOCTG As SAPbobsCOM.PaymentTermsTypes = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim sSQLSelect As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim oXml As Xml.XmlDocument = Nothing
        Dim oXmlNode As Xml.XmlNode = Nothing
        Dim sGroupNum As String = ""
        Dim oBslineDate As SAPbobsCOM.BoBaselineDate = Nothing
        Dim cCredLimit As Double = 0
        Dim sDiscCode As String = ""
        Dim cVolumDscnt As Double = 0
        Dim cLatePyChrg As Double = 0
        Dim cObligLimit As Double = 0
        Dim sExtraDays As String = ""
        Dim sExtraMonth As String = ""
        Dim sTolDays As String = ""
        Dim oOpenRcpt As SAPbobsCOM.BoOpenIncPayment = Nothing
        Dim oPayDuMonth As SAPbobsCOM.BoPayTermDueTypes = Nothing
        Dim sInstNum As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OCTG.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQLSelect = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t1.codeTable2 " & _
                         "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                         "WHERE t1.tableName = 'OCTG' " & _
                         "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQLSelect)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        oCompanyD.StartTransaction()

                        oRs = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOCTG = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentTermsTypes), SAPbobsCOM.PaymentTermsTypes)

                        If oOCTG.GetByKey(CInt(oDt.Rows.Item(i).Item("codeTable").ToString)) = True Then
                            sXML = oOCTG.GetAsXML
                        Else
                            sXML = ""
                        End If

                        'Porque en el modo Update no funciona por XML
                        oBslineDate = oOCTG.BaselineDate
                        cCredLimit = oOCTG.CreditLimit
                        sDiscCode = oOCTG.DiscountCode
                        cVolumDscnt = oOCTG.GeneralDiscount
                        cLatePyChrg = oOCTG.InterestOnArrears
                        cObligLimit = oOCTG.LoadLimit
                        sExtraDays = oOCTG.NumberOfAdditionalDays
                        sExtraMonth = oOCTG.NumberOfAdditionalMonths
                        sTolDays = oOCTG.NumberOfToleranceDays
                        oOpenRcpt = oOCTG.OpenReceipt
                        oPayDuMonth = oOCTG.StartFrom
                        sInstNum = oOCTG.NumberOfInstallments

                        sSQL = ""
                        '''''''''''''''''''''''''''''''''''''''''''''

                        If sXML <> "" Then
                            Try
                                oXml = New Xml.XmlDocument
                                oXml.LoadXml(sXML)

                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/PaymentTermsTypes/row/PriceListNo")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)

                                sXML = oXml.OuterXml
                            Catch ex As Exception

                            End Try

                            oOCTG = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentTermsTypes), SAPbobsCOM.PaymentTermsTypes)

                            oOCTG = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            sGroupNum = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCTG]", "GroupNum", "PymntGroup = '" & oDt.Rows.Item(i).Item("codeTable2").ToString & "'")

                            If sGroupNum = "" Then
                                'Añadir
                                If oOCTG.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If

                                sSQL = "UPDATE [" & sDBD & "].dbo.[OCTG] SET InstNum = " & sInstNum & " " & _
                                       "WHERE GroupNum = " & sGroupNum & "; "

                                If Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[CTG1]", "CTGCode", "CTGCode = " & oDt.Rows.Item(i).Item("codeTable").ToString & "") <> "" Then
                                    sSQL &= "INSERT INTO [" & sDBD & "].dbo.[CTG1] " & _
                                            "SELECT " & sGroupNum & ", [IntsNo], [InstMonth], [InstDays], [InstPrcnt] " & _
                                            "FROM [" & sDBO & "].dbo.[CTG1] t0 WITH (NOLOCK) " & _
                                            "WHERE t0.[CTGCode] = " & oDt.Rows.Item(i).Item("codeTable").ToString & "; "
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oOCTG.GetByKey(CInt(sGroupNum)) = True Then
                                    oOCTG.BaselineDate = oBslineDate
                                    oOCTG.CreditLimit = cCredLimit
                                    oOCTG.DiscountCode = sDiscCode
                                    oOCTG.GeneralDiscount = cVolumDscnt
                                    oOCTG.InterestOnArrears = cLatePyChrg
                                    oOCTG.LoadLimit = cObligLimit
                                    oOCTG.NumberOfAdditionalDays = sExtraDays
                                    oOCTG.NumberOfAdditionalMonths = sExtraMonth
                                    oOCTG.NumberOfToleranceDays = sTolDays
                                    oOCTG.OpenReceipt = oOpenRcpt
                                    oOCTG.StartFrom = oPayDuMonth

                                    If oOCTG.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If

                                    sSQL = "UPDATE [" & sDBD & "].dbo.[OCTG] SET InstNum = " & sInstNum & " " & _
                                           "WHERE GroupNum = " & sGroupNum & "; "

                                    If Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[CTG1]", "CTGCode", "CTGCode = " & oDt.Rows.Item(i).Item("codeTable").ToString & "") <> "" Then
                                        sSQL &= "DELETE FROM [" & sDBD & "].dbo.[CTG1] WHERE CTGCode = " & sGroupNum & "; "

                                        sSQL &= "INSERT INTO [" & sDBD & "].dbo.[CTG1] " & _
                                                "SELECT " & sGroupNum & ", [IntsNo], [InstMonth], [InstDays], [InstPrcnt] " & _
                                                "FROM [" & sDBO & "].dbo.[CTG1] t0 WITH (NOLOCK) " & _
                                                "WHERE t0.[CTGCode] = " & oDt.Rows.Item(i).Item("codeTable").ToString & "; "
                                    End If
                                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOCTG.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If

                        sSQL &= "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'; "

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

            If oCompanyD.InTransaction = True Then
                oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)

            If oCompanyD.InTransaction = True Then
                oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOCTG IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOCTG)
            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRs)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OPYM()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOPYM As SAPbobsCOM.WizardPaymentMethods = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim sSQLSelect As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim oXml As Xml.XmlDocument = Nothing
        Dim oXmlNode As Xml.XmlNode = Nothing
        Dim sFormat As String = ""
        Dim sNegPymCode As String = ""
        Dim sBnkDflt As String = ""
        Dim sBankCountr As String = ""
        Dim sDflAccount As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OPYM.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQLSelect = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                         "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                         "WHERE t1.tableName = 'OPYM' " & _
                         "ORDER BY t1.dbNameOrig ASC, t1.dbNameDest ASC, t1.codeTable2 DESC "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQLSelect)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        oCompanyD.StartTransaction()

                        oRs = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOPYM = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWizardPaymentMethods), SAPbobsCOM.WizardPaymentMethods)

                        If oOPYM.GetByKey(oDt.Rows.Item(i).Item("codeTable").ToString) = True Then
                            sXML = oOPYM.GetAsXML
                        Else
                            sXML = ""
                        End If

                        If sXML <> "" Then
                            'Esto es porque el código del formato de la dirección no tiene por qué ser igual en todas las empresas
                            sFormat = ""

                            If oOPYM.Format <> "" Then
                                sFormat = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OFRM]", "Name", "AbsEntry = " & oOPYM.Format & "")
                                sFormat = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OFRM]", "AbsEntry", "Name = '" & sFormat & "'")
                            End If

                            'Esto es porque el campo método de pago negativo no tiene propiedad en el objeto de SAP WizardPaymentMethods para rellenar este campo
                            sNegPymCode = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OPYM]", "NegPymCode", "PayMethCod = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'")

                            'Esto es para que mantenga el banco de la vía de pago destino
                            sBnkDflt = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OPYM]", "BnkDflt", "PayMethCod = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'")
                            sBankCountr = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OPYM]", "BankCountr", "PayMethCod = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'")
                            sDflAccount = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OPYM]", "DflAccount", "PayMethCod = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'")

                            oXml = New Xml.XmlDocument
                            oXml.LoadXml(sXML)

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/WizardPaymentMethods/row/BankCountry")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/WizardPaymentMethods/row/DefaultBank")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/WizardPaymentMethods/row/DefaultAccount")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/WizardPaymentMethods/row/Branch")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/WizardPaymentMethods/row/GLAccount")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/WizardPaymentMethods/row/BankAccountKey")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            sXML = oXml.OuterXml

                            oOPYM = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWizardPaymentMethods), SAPbobsCOM.WizardPaymentMethods)

                            oOPYM = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OPYM]", "PayMethCod", "PayMethCod = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                                'Añadir
                                If sFormat = "" Then
                                    oOPYM.Format = ""
                                Else
                                    oOPYM.Format = sFormat
                                End If

                                If oOPYM.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                If sFormat = "" Then
                                    oOPYM.Format = ""
                                Else
                                    oOPYM.Format = sFormat
                                End If

                                oOPYM.DefaultBank = sBnkDflt
                                oOPYM.BankCountry = sBankCountr
                                oOPYM.DefaultAccount = sDflAccount

                                If oOPYM.Update() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            End If

                            oRs.DoQuery("UPDATE OPYM SET NegPymCode = '" & sNegPymCode & "' WHERE PayMethCod = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'")
                        End If

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

            If oCompanyD.InTransaction = True Then
                oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)

            If oCompanyD.InTransaction = True Then
                oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOPYM IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOPYM)
            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRs)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OOCR()
        'Dim oCompanyO As SAPbobsCOM.Company = Nothing
        'Dim oCompanyD As SAPbobsCOM.Company = Nothing

        'Dim oCmpSrvO As SAPbobsCOM.CompanyService = Nothing
        'Dim oCCTServiceO As Object = Nothing
        'Dim oCCTParamsO As Object = Nothing

        'Dim oCmpSrvD As SAPbobsCOM.CompanyService = Nothing
        'Dim oCCTServiceD As Object = Nothing
        'Dim oCCTParamsD As Object = Nothing

        'Dim oOOCR As SAPbobsCOM.DistributionRule = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        'Dim sXML As String = ""
        'Dim sDimCode As String = ""
        Dim oTransaction As SqlTransaction = Nothing

        'Tiene objeto Service pero no funciona bien, porque no rellena bien el campo TotalFactor de las líneas. Por tanto lo hacemos por SQL.
        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OOCR.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OOCR' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                'Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                'oCmpSrvO = oCompanyO.GetCompanyService()
                'oCCTServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
                'oCCTParamsO = oCCTServiceO.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRuleParams)

                'Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)
                'oCmpSrvD = oCompanyD.GetCompanyService()
                'oCCTServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
                'oCCTParamsD = oCCTServiceD.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRuleParams)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            ''Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            'Conexiones.Disconnect_Company(oCompanyO)

                            'Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)
                            'oCmpSrvO = oCompanyO.GetCompanyService()
                            'oCCTServiceO = oCmpSrvO.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
                            'oCCTParamsO = oCCTServiceO.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRuleParams)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            ''Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            'Conexiones.Disconnect_Company(oCompanyD)

                            'Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)
                            'oCmpSrvD = oCompanyD.GetCompanyService()
                            'oCCTServiceD = oCmpSrvD.GetBusinessService(SAPbobsCOM.ServiceTypes.DistributionRulesService)
                            'oCCTParamsD = oCCTServiceD.GetDataInterface(SAPbobsCOM.DistributionRulesServiceDataInterfaces.drsDistributionRuleParams)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oTransaction = oDB.BeginTransaction("OOCR")

                        'oCCTParamsO.FactorCode = oDt.Rows.Item(i).Item("codeTable").ToString
                        'oOOCR = oCCTServiceO.GetDistributionRule(oCCTParamsO)

                        'sXML = oOOCR.ToXMLString

                        'sDimCode = oOOCR.InWhichDimension

                        sSQL = ""

                        'If sXML <> "" Then
                        If Conexiones.GetValueDB(oDB, oTransaction, "[" & sDBD & "].dbo.[OOCR]", "OcrCode", "OcrCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                            'Añadir
                            sSQL = "INSERT INTO [" & sDBD & "].dbo.[OOCR] " & _
                                   "SELECT [OcrCode], [OcrName], [OcrTotal], [Direct], [Locked], [DataSource], [UserSign], " & _
                                   "[DimCode], [AbsEntry], [Active], [logInstanc], [UserSign2], [updateDate] " & _
                                   "FROM [" & sDBO & "].dbo.[OOCR] t0 WITH (NOLOCK) " & _
                                   "WHERE t0.[OcrCode] = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'; "
                        Else
                            'oCCTParamsD.FactorCode = oDt.Rows.Item(i).Item("codeTable").ToString
                            'oOOCR = oCCTServiceD.GetDistributionRule(oCCTParamsD)

                            'oOOCR.FromXMLString(sXML)

                            'oOOCR.InWhichDimension = sDimCode

                            'oCCTServiceD.UpdateDistributionRule(oOOCR)

                            'Modificar"
                            sSQL = "UPDATE t1 SET [OcrName] = t0.[OcrName], " & _
                                   "[OcrTotal] = t0.[OcrTotal],  " & _
                                   "[Direct] = t0.[Direct], " & _
                                   "[Locked] = t0.[Locked], " & _
                                   "[DataSource] = t0.[DataSource], " & _
                                   "[UserSign] = t0.[UserSign], " & _
                                   "[DimCode] = t0.[DimCode], " & _
                                   "[AbsEntry] = t0.[AbsEntry], " & _
                                   "[Active] = t0.[Active], " & _
                                   "[logInstanc] = t0.[logInstanc], " & _
                                   "[UserSign2] = t0.[UserSign2], " & _
                                   "[updateDate] = t0.[updateDate] " & _
                                   "FROM [" & sDBO & "].dbo.[OOCR] t0 WITH (NOLOCK) INNER JOIN " & _
                                   "[" & sDBD & "].dbo.[OOCR] t1 WITH (NOLOCK) ON t0.[OcrCode] = t1.[OcrCode] " & _
                                   "WHERE t0.[OcrCode] = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'; "
                        End If

                        sSQL &= "DELETE FROM [" & sDBD & "].dbo.[OCR1] WHERE [OcrCode] = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'; "

                        sSQL &= "INSERT INTO [" & sDBD & "].dbo.[OCR1] " & _
                               "SELECT [OcrCode], [PrcCode], [PrcAmount], [OcrTotal], [Direct], [UserSign], [ValidFrom], " & _
                               "[ValidTo], [logInstanc], [UserSign2], [updateDate] " & _
                               "FROM [" & sDBO & "].dbo.[OCR1] t0 WITH (NOLOCK) " & _
                               "WHERE t0.[OcrCode] = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'; "
                        'End If

                        sSQL &= "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'; "

                        Conexiones.ExecuteSQLDB(oDB, oTransaction, sSQL)

                        If oTransaction IsNot Nothing Then oTransaction.Commit()

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oTransaction IsNot Nothing Then oTransaction.Rollback()
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oTransaction IsNot Nothing Then oTransaction.Rollback()
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

            If oTransaction IsNot Nothing Then oTransaction.Rollback()
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)

            If oTransaction IsNot Nothing Then oTransaction.Rollback()
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()

            Conexiones.Disconnect_SQLServer(oDB)
        End Try
    End Sub

    Public Shared Sub OCRC()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOCRC As SAPbobsCOM.CreditCards = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim sSQLSelect As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim sCreditCard As String = ""
        Dim sCompanyId As String = ""
        Dim sCountry As String = ""
        Dim sCardName As String = ""
        Dim sAcctCode As String = ""
        Dim sPhone As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OCRC.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQLSelect = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t1.codeTable2 " & _
                         "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                         "WHERE t1.tableName = 'OCRC' " & _
                         "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQLSelect)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOCRC = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditCards), SAPbobsCOM.CreditCards)

                        If oOCRC.GetByKey(CInt(oDt.Rows.Item(i).Item("codeTable").ToString)) = True Then
                            sXML = oOCRC.GetAsXML
                        Else
                            sXML = ""
                        End If

                        'Porque en el modo Update no funciona por XML
                        sCompanyId = oOCRC.CompanyID
                        sCountry = oOCRC.CountryCode
                        sCardName = oOCRC.CreditCardName
                        sAcctCode = oOCRC.GLAccount
                        sPhone = oOCRC.Telephone
                        '''''''''''''''''''''''''''''''''''''''''''''

                        If sXML <> "" Then
                            oOCRC = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditCards), SAPbobsCOM.CreditCards)

                            oOCRC = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            sCreditCard = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCRC]", "CreditCard", "CardName = '" & oDt.Rows.Item(i).Item("codeTable2").ToString & "'")

                            If sCreditCard = "" Then
                                'Añadir
                                If oOCRC.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oOCRC.GetByKey(CInt(sCreditCard)) = True Then
                                    oOCRC.CompanyID = sCompanyId
                                    oOCRC.CountryCode = sCountry
                                    oOCRC.CreditCardName = sCardName
                                    oOCRC.GLAccount = sAcctCode
                                    oOCRC.Telephone = sPhone

                                    If oOCRC.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOCRC.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOCRC IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOCRC)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub MODFINAN()
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim oTransaction As SqlTransaction = Nothing

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_MODFINAN.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'MODFINAN' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)


            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oTransaction = oDB.BeginTransaction("MODFINAN")

                        sSQL = "TRUNCATE TABLE [" & sDBD & "].dbo.[OFRT] "

                        sSQL &= "TRUNCATE TABLE [" & sDBD & "].dbo.[FRC1] "

                        sSQL &= "TRUNCATE TABLE [" & sDBD & "].dbo.[OFRC] "

                        sSQL &= "INSERT INTO [" & sDBD & "].dbo.[OFRT] " & _
                                "SELECT AbsId, Name, DocType, FRTCounter, MoveChk1, MoveChk2, MoveTo_1, MoveTo_2, Title_1, " & _
                                "Title_2, ShowMiss, ToTitle_1, ToTitle_2, UserSign, DimCode " & _
                                "FROM [" & sDBO & "].dbo.[OFRT] WITH (NOLOCK) "

                        sSQL &= "UPDATE t1 SET [AutoKey] = t0.[AutoKey] " & _
                               "FROM [" & sDBO & "].dbo.[ONNM] t0 WITH (NOLOCK) INNER JOIN " & _
                               "[" & sDBD & "].dbo.[ONNM] t1 WITH (NOLOCK) ON t0.[ObjectCode] = t1.[ObjectCode] AND t0.[DocSubType] = t1.[DocSubType] " & _
                               "WHERE t0.[ObjectCode] = '95' "

                        sSQL &= "INSERT INTO [" & sDBD & "].dbo.[OFRC] " & _
                                "SELECT CatId, TemplateId, Name, FrgnName, Levels, FatherNum, Active, HasSons, VisOrder, SubSum, " & _
                                "SubName, Furmula, Param_1, Param_2, Param_3, Param_4, Param_5, Param_6, Param_7, Param_8, " & _
                                "Param_9, Param_10, Param_11, Param_12, Param_13, Param_14, Param_15, Param_16, Param_17, " & _
                                "Param_18, Param_19, Param_20, Param_21, Param_22, Param_23, Param_24, Param_25, OP_1, OP_2, " & _
                                "OP_3, OP_4, OP_5, OP_6, OP_7, OP_8, OP_9, OP_10, OP_11, OP_12, OP_13, OP_14, OP_15, OP_16, " & _
                                "OP_17, OP_18, OP_19, OP_20, OP_21, OP_22, OP_23, OP_24, ProfitLoss, MoveNeg, [Dummy], HideAct, " & _
                                "UserSign, ToGroup, ToTitle, LineNum, IndentChar, Reversal, TextTitle, SumType, NetIncome, " & _
                                "PLTempId, CustName, ExtFromBS, ExtData, LegalRef, PLCatId, SignAggr, Mandatory, AcctReq, " & _
                                "NotPermit, KPIFactor, CatCode, CatClass " & _
                                "FROM [" & sDBO & "].dbo.[OFRC] WITH (NOLOCK) "

                        sSQL &= "INSERT INTO [" & sDBD & "].dbo.[FRC1] " & _
                                "SELECT CatId, TemplateId, AcctCode, VisOrder, CFWId, CalcMethod, SlpCode, PrcCode, CalMethod2, CalMethod3, Linked, Sign " & _
                                "FROM [" & sDBO & "].dbo.[FRC1] WITH (NOLOCK) " & _
                                "WHERE AcctCode IN (SELECT t1.AcctCode " & _
                                                   "FROM [" & sDBD & "].dbo.[OACT] t1 WITH (NOLOCK)) "

                        sSQL &= "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, oTransaction, sSQL)

                        If oTransaction IsNot Nothing Then oTransaction.Commit()

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oTransaction IsNot Nothing Then oTransaction.Rollback()
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oTransaction IsNot Nothing Then oTransaction.Rollback()
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

            If oTransaction IsNot Nothing Then oTransaction.Rollback()
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)

            If oTransaction IsNot Nothing Then oTransaction.Rollback()
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()

            Conexiones.Disconnect_SQLServer(oDB)
        End Try
    End Sub

    Public Shared Sub OACT()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOACT As SAPbobsCOM.ChartOfAccounts = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        'Dim oPurpCode As SAPbobsCOM.SPEDContabilAccountPurposeCode = Nothing
        'Dim bPurpCode As Boolean = False
        Dim oActType As SAPbobsCOM.BoAccountTypes = Nothing
        Dim sActCurr As String = ""
        Dim oPostable As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oVatChange As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oMultiLink As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oBlocManPos As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim sBPLId As String = ""
        Dim oBudget As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oCashBox As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oCfwRlvnt As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim sCategory As String = "0"
        Dim sExportCode As String = ""
        Dim sDatevAcct As String = ""
        Dim oDatevAutoA As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oDatevFirst As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim sDfltVat As String = ""
        Dim sDetails As String = ""
        Dim oDim2Relvnt As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oDim3Relvnt As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oDim4Relvnt As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oDim5Relvnt As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oDim1Relvnt As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim sAccntntCod As String = ""
        Dim sFatherNum As String = ""
        Dim sFrgnName As String = ""
        Dim sFormatCode As String = ""
        Dim oFrozenFor As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim dFrozenFrom As Date = Nothing
        Dim sFrozenComm As String = ""
        Dim dFrozenTo As Date = Nothing
        Dim oAdvance As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim sOverCode As String = ""
        Dim sOverCode2 As String = ""
        Dim sOverCode3 As String = ""
        Dim sOverCode4 As String = ""
        Dim sOverCode5 As String = ""
        Dim oOverType As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oLocManTran As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim sAcctName As String = ""
        Dim sPlngLevel As String = ""
        Dim sProject As String = ""
        Dim oPrjRelvnt As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oProtected As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oRateTrans As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oRealAcct As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim sRefCode As String = ""
        Dim oRevalMatch As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oExmIncome As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim oTaxPostAcc As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim sTransCode As String = ""
        Dim oValidFor As SAPbobsCOM.BoYesNoEnum = Nothing
        Dim dValidFrom As Date = Nothing
        Dim sValidComm As String = ""
        Dim dValidTo As Date = Nothing
        Dim oUserFields As SAPbobsCOM.UserFields = Nothing

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OACT.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable, t2.dbTipo " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) INNER JOIN " & _
                   "[INTERCOMPANY].dbo.[DATABASES] t2 WITH (NOLOCK) ON t1.dbNameDest = t2.dbName " & _
                   "WHERE t1.tableName = 'OACT' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOACT = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts), SAPbobsCOM.ChartOfAccounts)

                        If oOACT.GetByKey(oDt.Rows.Item(i).Item("codeTable").ToString) = True Then
                            sXML = oOACT.GetAsXML
                        Else
                            sXML = ""
                        End If

                        ''Porque en el modo Update no funciona por XML
                        'Try
                        '    bPurpCode = True
                        '    oPurpCode = oOACT.AccountPurposeCode
                        'Catch exCOM As System.Runtime.InteropServices.COMException
                        '    bPurpCode = False
                        'Catch ex As Exception
                        '    bPurpCode = False
                        'End Try

                        oActType = oOACT.AccountType
                        sActCurr = oOACT.AcctCurrency
                        oPostable = oOACT.ActiveAccount
                        oVatChange = oOACT.AllowChangeVatGroup
                        oMultiLink = oOACT.AllowMultipleLinking
                        oBlocManPos = oOACT.BlockManualPosting
                        sBPLId = oOACT.BPLID
                        oBudget = oOACT.BudgetAccount
                        oCashBox = oOACT.CashAccount
                        oCfwRlvnt = oOACT.CashFlowRelevant
                        sExportCode = oOACT.DataExportCode
                        sDatevAcct = oOACT.DatevAccount
                        oDatevAutoA = oOACT.DatevAutoAccount
                        oDatevFirst = oOACT.DatevFirstDataEntry
                        sDfltVat = oOACT.DefaultVatGroup
                        sDetails = oOACT.Details
                        oDim2Relvnt = oOACT.DistributionRule2Relevant
                        oDim3Relvnt = oOACT.DistributionRule3Relevant
                        oDim4Relvnt = oOACT.DistributionRule4Relevant
                        oDim5Relvnt = oOACT.DistributionRule5Relevant
                        oDim1Relvnt = oOACT.DistributionRuleRelevant
                        sAccntntCod = oOACT.ExternalCode
                        sFatherNum = oOACT.FatherAccountKey
                        sFrgnName = oOACT.ForeignName
                        sFormatCode = oOACT.FormatCode
                        oFrozenFor = oOACT.FrozenFor
                        dFrozenFrom = oOACT.FrozenFrom
                        sFrozenComm = oOACT.FrozenRemarks
                        dFrozenTo = oOACT.FrozenTo
                        oAdvance = oOACT.LiableForAdvances
                        sOverCode = oOACT.LoadingFactorCode
                        sOverCode2 = oOACT.LoadingFactorCode2
                        sOverCode3 = oOACT.LoadingFactorCode3
                        sOverCode4 = oOACT.LoadingFactorCode4
                        sOverCode5 = oOACT.LoadingFactorCode5
                        oOverType = oOACT.LoadingType
                        oLocManTran = oOACT.LockManualTransaction
                        sAcctName = oOACT.Name
                        sPlngLevel = oOACT.PlanningLevel
                        'sProject = oOACT.ProjectCode
                        'oPrjRelvnt = oOACT.ProjectRelevant
                        oProtected = oOACT.Protected
                        oRateTrans = oOACT.RateConversion
                        oRealAcct = oOACT.ReconciledAccount
                        sRefCode = oOACT.ReferentialAccountCode
                        oRevalMatch = oOACT.RevaluationCoordinated
                        oExmIncome = oOACT.TaxExemptAccount
                        oTaxPostAcc = oOACT.TaxLiableAccount
                        sTransCode = oOACT.TransactionCode
                        oValidFor = oOACT.ValidFor
                        dValidFrom = oOACT.ValidFrom
                        sValidComm = oOACT.ValidRemarks
                        dValidTo = oOACT.ValidTo
                        oUserFields = oOACT.UserFields
                        '''''''''''''''''''''''''''''''''''''''''''''

                        If sXML <> "" Then
                            'Esto es porque el código del formato de la dirección no tiene por qué ser igual en todas las empresas
                            sCategory = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OACG]", "Name", "AbsId = " & oOACT.Category & "")
                            sCategory = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OACG]", "AbsId", "Name = '" & sCategory & "'")

                            oOACT = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts), SAPbobsCOM.ChartOfAccounts)

                            oOACT = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OACT]", "AcctCode", "AcctCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                                'Añadir
                                oOACT.ProjectRelevant = SAPbobsCOM.BoYesNoEnum.tNO
                                oOACT.ProjectCode = ""

                                If sCategory <> "" Then
                                    oOACT.Category = sCategory
                                End If

                                'Si la Empresa es de consolidación, el campo cuenta asociada debe ser No
                                If oDt.Rows.Item(i).Item("dbTipo").ToString = "C" Then
                                    oOACT.LockManualTransaction = SAPbobsCOM.BoYesNoEnum.tNO
                                End If

                                If oOACT.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'Porque en el modo Update no funciona por XML
                                If oOACT.GetByKey(oDt.Rows.Item(i).Item("codeTable").ToString) = True Then
                                    'If bPurpCode = True Then oOACT.AccountPurposeCode = oPurpCode
                                    oOACT.AccountType = oActType
                                    oOACT.AcctCurrency = sActCurr
                                    oOACT.ActiveAccount = oPostable
                                    oOACT.AllowChangeVatGroup = oVatChange
                                    oOACT.AllowMultipleLinking = oMultiLink
                                    oOACT.BlockManualPosting = oBlocManPos
                                    oOACT.BPLID = sBPLId
                                    oOACT.BudgetAccount = oBudget
                                    oOACT.CashAccount = oCashBox
                                    oOACT.CashFlowRelevant = oCfwRlvnt

                                    If sCategory <> "" Then
                                        oOACT.Category = sCategory
                                    End If

                                    oOACT.DataExportCode = sExportCode
                                    oOACT.DatevAccount = sDatevAcct
                                    oOACT.DatevAutoAccount = oDatevAutoA
                                    oOACT.DatevFirstDataEntry = oDatevFirst
                                    oOACT.DefaultVatGroup = sDfltVat
                                    oOACT.Details = sDetails
                                    oOACT.DistributionRule2Relevant = oDim2Relvnt
                                    oOACT.DistributionRule3Relevant = oDim3Relvnt
                                    oOACT.DistributionRule4Relevant = oDim4Relvnt
                                    oOACT.DistributionRule5Relevant = oDim5Relvnt
                                    oOACT.DistributionRuleRelevant = oDim1Relvnt
                                    oOACT.ExternalCode = sAccntntCod
                                    oOACT.FatherAccountKey = sFatherNum
                                    oOACT.ForeignName = sFrgnName
                                    oOACT.FormatCode = sFormatCode
                                    oOACT.FrozenFor = oFrozenFor
                                    oOACT.FrozenFrom = dFrozenFrom
                                    oOACT.FrozenRemarks = sFrozenComm
                                    oOACT.FrozenTo = dFrozenTo
                                    oOACT.LiableForAdvances = oAdvance
                                    oOACT.LoadingFactorCode = sOverCode
                                    oOACT.LoadingFactorCode2 = sOverCode2
                                    oOACT.LoadingFactorCode3 = sOverCode3
                                    oOACT.LoadingFactorCode4 = sOverCode4
                                    oOACT.LoadingFactorCode5 = sOverCode5
                                    oOACT.LoadingType = oOverType

                                    'Si la Empresa es de consolidación, el campo cuenta asociada debe ser No
                                    If oDt.Rows.Item(i).Item("dbTipo").ToString = "C" Then
                                        oOACT.LockManualTransaction = SAPbobsCOM.BoYesNoEnum.tNO
                                    Else
                                        oOACT.LockManualTransaction = oLocManTran
                                    End If

                                    oOACT.Name = sAcctName
                                    oOACT.PlanningLevel = sPlngLevel
                                    oOACT.ProjectCode = sProject
                                    'oOACT.ProjectRelevant = oPrjRelvnt
                                    'oOACT.Protected = oProtected
                                    oOACT.ProjectRelevant = SAPbobsCOM.BoYesNoEnum.tNO
                                    oOACT.ProjectCode = ""
                                    oOACT.RateConversion = oRateTrans
                                    oOACT.ReconciledAccount = oRealAcct
                                    oOACT.ReferentialAccountCode = sRefCode
                                    oOACT.RevaluationCoordinated = oRevalMatch
                                    oOACT.TaxExemptAccount = oExmIncome
                                    oOACT.TaxLiableAccount = oTaxPostAcc
                                    oOACT.TransactionCode = sTransCode
                                    oOACT.ValidFor = oValidFor
                                    oOACT.ValidFrom = dValidFrom
                                    oOACT.ValidRemarks = sValidComm
                                    oOACT.ValidTo = dValidTo

                                    For h As Integer = 0 To oUserFields.Fields.Count - 1
                                        If oOACT.UserFields.Fields.Item(oUserFields.Fields.Item(h).Name).IsNull = SAPbobsCOM.BoYesNoEnum.tNO Then
                                            oOACT.UserFields.Fields.Item(oUserFields.Fields.Item(h).Name).Value = oUserFields.Fields.Item(oUserFields.Fields.Item(h).Name).Value
                                        End If
                                    Next

                                    If oOACT.Update() <> 0 Then
                                        Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                    End If
                                End If
                                ''''''''''''''''''''''''''''''''''''''''''''''

                                'If oOACT.Update() <> 0 Then
                                '    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                'End If
                            End If
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try

                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOACT IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOACT)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OCRD()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim oOCRD As SAPbobsCOM.BusinessPartners = Nothing
        Dim oOCRD2 As SAPbobsCOM.BusinessPartners = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oDB As SqlConnection = Nothing
        Dim log As EXO_Log.EXO_Log = Nothing
        Dim sSQL As String = ""
        Dim oDt As System.Data.DataTable = Nothing
        Dim sDBO As String = ""
        Dim sDBD As String = ""
        Dim i As Integer = -1
        Dim sXML As String = ""
        Dim oXml As Xml.XmlDocument = Nothing
        Dim oXml2 As Xml.XmlDocument = Nothing
        Dim oXmlNode As Xml.XmlNode = Nothing
        Dim oXmlNode2 As Xml.XmlNode = Nothing
        Dim oXmlNodes As Xml.XmlNodeList = Nothing
        Dim oXmlNodes2 As Xml.XmlNodeList = Nothing
        Dim sListNum As String = ""
        Dim sTrnspCode As String = ""
        Dim sIndCode As String = ""
        Dim sGroupNum As String = ""
        Dim sAbsEntryOPYB As String = ""
        Dim sCreditCard As String = ""
        Dim sTerritryID As String = ""
        Dim sCodeOLNG As String = ""
        Dim sGroupCode As String = ""
        Dim sU_EXO_GRUPOEMPRESA As String = ""
        Dim sInternalCode As String = ""
        Dim sInternalKey As String = ""
        Dim sBankCode As String = ""
        Dim sCountry As String = ""
        Dim sAccount As String = ""
        Dim sAgentCode As String = ""
        Dim bExisteCE As Boolean = False
        Dim sCardType As String = ""
        Dim sSeries As String = ""

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OCRD.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)

            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
                   "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
                   "WHERE t1.tableName = 'OCRD' " & _
                   "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        oCompanyD.StartTransaction()

                        oRs = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                        oCompanyO.XMLAsString = True
                        oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oCompanyD.XMLAsString = True
                        oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                        oOCRD = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)

                        If oOCRD.GetByKey(oDt.Rows.Item(i).Item("codeTable").ToString) = True Then
                            sXML = oOCRD.GetAsXML
                        Else
                            sXML = ""
                        End If

                        ' Guardamos el grupo de empresas para actualizar después
                        sU_EXO_GRUPOEMPRESA = oOCRD.UserFields.Fields.Item("U_EXO_GRUPOEMPRESA").Value.ToString
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        sSQL = ""

                        If sXML <> "" Then
                            'Esto es porque hay ciertos campos que son autonuméricos y no tienen por qué ser igual en todas las empresas
                            sTrnspCode = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OSHP]", "TrnspName", "TrnspCode = " & oOCRD.ShippingType & "")
                            sTrnspCode = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OSHP]", "TrnspCode", "TrnspName = '" & sTrnspCode & "'")
                            sIndCode = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OOND]", "IndName", "IndCode = " & oOCRD.Industry & "")
                            sIndCode = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OOND]", "IndCode", "IndName = '" & sIndCode & "'")
                            sGroupNum = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OCTG]", "PymntGroup", "GroupNum = " & oOCRD.PayTermsGrpCode & "")
                            sGroupNum = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCTG]", "GroupNum", "PymntGroup = '" & sGroupNum & "'")
                            sAbsEntryOPYB = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OPYB]", "PayBlock", "AbsEntry = " & oOCRD.PaymentBlockDescription & "")
                            sAbsEntryOPYB = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OPYB]", "AbsEntry", "PayBlock = '" & sAbsEntryOPYB & "'")
                            sCreditCard = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OCRC]", "CardName", "CreditCard = " & oOCRD.CreditCardCode & "")
                            sCreditCard = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCRC]", "CreditCard", "CardName = '" & sCreditCard & "'")
                            sTerritryID = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OTER]", "descript", "territryID = " & oOCRD.Territory & "")
                            sTerritryID = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OTER]", "territryID", "descript = '" & sTerritryID & "'")
                            sCodeOLNG = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OLNG]", "ShortName", "Code = " & oOCRD.LanguageCode & "")
                            sCodeOLNG = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OLNG]", "Code", "ShortName = '" & sCodeOLNG & "'")
                            sGroupCode = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OCRG]", "GroupName", "GroupCode = " & oOCRD.GroupCode & "")
                            sGroupCode = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCRG]", "GroupCode", "GroupName = '" & sGroupCode & "'")

                            'Esto es porque el campo responsable no tiene propiedad en el objeto de SAP BusinessPartners para rellenar este campo
                            sAgentCode = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OCRD]", "AgentCode", "CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'")

                            'La serie de ICs en InterCompany siempre la manual
                            sCardType = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OCRD]", "CardType", "CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'")
                            sSeries = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[NNM1]", "Series", "ObjectCode = '2' AND SeriesName = 'Manual' AND DocSubType = '" & sCardType & "'")

                            oXml = New Xml.XmlDocument
                            oXml.LoadXml(sXML)

                            'NOTA: el campo Vacaciones no lo gestiona el DI API, en principio no le replico

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/HouseBank")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/HouseBankCountry")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/HouseBankAccount")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/HouseBankBranch")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/ProjectCode")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            oXmlNodes = oXml.SelectNodes("/BOM/BO/BPBankAccounts/row")

                            For j As Integer = oXmlNodes.Count - 1 To 0 Step -1
                                sInternalKey = ""
                                sBankCode = ""
                                sCountry = ""
                                sAccount = ""

                                Try
                                    sBankCode = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OCRB]", "BankCode", "AbsEntry = " & oXmlNodes.Item(j).SelectSingleNode("InternalKey").InnerText & "")
                                    sCountry = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OCRB]", "Country", "AbsEntry = " & oXmlNodes.Item(j).SelectSingleNode("InternalKey").InnerText & "")
                                    sAccount = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OCRB]", "Account", "AbsEntry = " & oXmlNodes.Item(j).SelectSingleNode("InternalKey").InnerText & "")
                                    sInternalKey = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCRB]", "AbsEntry", "BankCode = '" & sBankCode & "' AND Country = '" & sCountry & "' AND Account = '" & sAccount & "' AND CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'")

                                    If sInternalKey <> "" Then
                                        oXmlNodes.Item(j).SelectSingleNode("InternalKey").InnerText = sInternalKey
                                    Else
                                        oXmlNode = oXmlNodes.Item(j).SelectSingleNode("InternalKey")
                                        oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                    End If
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("CustomerIdNumber")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try
                            Next

                            oXmlNodes = oXml.SelectNodes("/BOM/BO/ContactEmployees/row")

                            For j As Integer = oXmlNodes.Count - 1 To 0 Step -1
                                'sInternalCode = ""

                                Try
                                    'sInternalCode = Conexiones.GetValueDB(oDB, "[" & sDBO & "].dbo.[OCPR]", "Name", "CntctCode = " & oXmlNodes.Item(j).SelectSingleNode("InternalCode").InnerText & "")
                                    'sInternalCode = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCPR]", "CntctCode", "Name = '" & sInternalCode & "' AND CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'")

                                    'If sInternalCode <> "" Then
                                    '    oXmlNodes.Item(j).SelectSingleNode("InternalCode").InnerText = sInternalCode
                                    'Else
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("InternalCode")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                    'End If
                                Catch ex As Exception

                                End Try

                                'Los medios de comunicación de las personas de contacto no funcionan por DI API (sólo se añaden para el primer contacto y 
                                'si el resto de contactos tambien tiene medios de comunicación da error), por tanto lo quitamos porque da error.
                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("BlockSendingMarketingContent")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try
                            Next

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/AttachmentEntry")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/PriceListNum")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/SalesPersonCode")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/DunningTerm")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/AutomaticPosting")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/InterestAccount")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/FeeAccount")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/BankChargesAllocationCode")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                'El grupo de empresas al ser un combo que se rellena en tiempo de ejecución, no se puede actualizar el valor
                                'por DI API, por lo que tiene que ser por SQL. Por tanto borramos el nodo, y después actualizamos el valor por SQL.
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row/U_EXO_GRUPOEMPRESA")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            'Borramos los campos de usuario que empiecen por U_STEC
                            oXmlNode = oXml.SelectSingleNode("/BOM/BO/BusinessPartners/row")

                            oXmlNodes = oXmlNode.ChildNodes

                            For j As Integer = oXmlNodes.Count - 1 To 0 Step -1
                                oXmlNode = oXmlNodes.Item(j)

                                If Left(oXmlNode.LocalName.ToUpper, 6) = "U_STEC" Then
                                    Try
                                        oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                    Catch ex As Exception

                                    End Try
                                End If
                            Next

                            sXML = oXml.OuterXml

                            oOCRD = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)

                            oOCRD = oCompanyD.GetBusinessObjectFromXML(sXML, 0)

                            If Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCRD]", "CardCode", "CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'") = "" Then
                                'Añadir
                                oOCRD.Series = sSeries

                                If sTrnspCode <> "" Then
                                    oOCRD.ShippingType = sTrnspCode
                                End If

                                If sIndCode <> "" Then
                                    oOCRD.Industry = sIndCode
                                End If

                                If sGroupNum <> "" Then
                                    oOCRD.PayTermsGrpCode = sGroupNum
                                End If

                                If sAbsEntryOPYB <> "" Then
                                    oOCRD.PaymentBlockDescription = sAbsEntryOPYB
                                End If

                                If sCreditCard <> "" Then
                                    oOCRD.CreditCardCode = sCreditCard
                                End If

                                If sTerritryID <> "" Then
                                    oOCRD.Territory = sTerritryID
                                End If

                                If sCodeOLNG <> "" Then
                                    oOCRD.LanguageCode = sCodeOLNG
                                End If

                                If sGroupCode <> "" Then
                                    oOCRD.GroupCode = sGroupCode
                                End If

                                If oOCRD.Add() <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            Else
                                'Modificar"
                                'En el modo modificar si el IC ya tiene Cuentas asociadas da un error por tanto hay que borrarlas primero por SQL
                                sSQL = "DELETE FROM [" & sDBD & "].dbo.[CRD3] WHERE CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"
                                'En el modo modificar si el IC ya tiene Indicadores de retención permitidos da un error por tanto hay que borrarlas primero por SQL
                                sSQL &= "DELETE FROM [" & sDBD & "].dbo.[CRD4] WHERE CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'"

                                oRs.DoQuery(sSQL)

                                'En el modo modificar ponemos la propia lista de precios del IC, porque sino la actualización del IC por XML no funciona
                                sListNum = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCRD]", "ListNum", "CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'")

                                oOCRD.Series = sSeries

                                If sListNum <> "" Then
                                    oOCRD.PriceListNum = sListNum
                                End If

                                If sTrnspCode <> "" Then
                                    oOCRD.ShippingType = sTrnspCode
                                End If

                                If sIndCode <> "" Then
                                    oOCRD.Industry = sIndCode
                                End If

                                If sGroupNum <> "" Then
                                    oOCRD.PayTermsGrpCode = sGroupNum
                                End If

                                If sAbsEntryOPYB <> "" Then
                                    oOCRD.PaymentBlockDescription = sAbsEntryOPYB
                                End If

                                If sCreditCard <> "" Then
                                    oOCRD.CreditCardCode = sCreditCard
                                End If

                                If sTerritryID <> "" Then
                                    oOCRD.Territory = sTerritryID
                                End If

                                If sCodeOLNG <> "" Then
                                    oOCRD.LanguageCode = sCodeOLNG
                                End If

                                If sGroupCode <> "" Then
                                    oOCRD.GroupCode = sGroupCode
                                End If

                                'Esto lo hago porque el método del DI API GetBusinessObjectFromXML para los ICs borra el campo InternalCode de ContactEmployees.
                                'Por tanto vuelvo a recuperar el XML una vez modificado el objeto oOCRD para la empresa destino, y vuelvo a rellenar el campo InternalCode.
                                sXML = oOCRD.GetAsXML

                                oXml.LoadXml(sXML)

                                oOCRD2 = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)

                                oOCRD2.GetByKey(oDt.Rows.Item(i).Item("codeTable").ToString)

                                sXML = oOCRD2.GetAsXML

                                oXml2 = New Xml.XmlDocument
                                oXml2.LoadXml(sXML)

                                oXmlNodes = oXml.SelectNodes("/BOM/BO/ContactEmployees/row")
                                oXmlNodes2 = oXml2.SelectNodes("/BOM/BO/ContactEmployees/row")

                                For h As Integer = oXmlNodes2.Count - 1 To 0 Step -1
                                    bExisteCE = False

                                    For j As Integer = oXmlNodes.Count - 1 To 0 Step -1
                                        If oXmlNodes2.Item(h).SelectSingleNode("Name").InnerText = oXmlNodes.Item(j).SelectSingleNode("Name").InnerText Then
                                            bExisteCE = True

                                            Exit For
                                        End If
                                    Next

                                    If bExisteCE = False Then
                                        oXmlNode2 = oXml.ImportNode(oXmlNodes2.Item(h).SelectSingleNode("Name").ParentNode, True)

                                        oXmlNode = oXml.SelectSingleNode("/BOM/BO/ContactEmployees")

                                        If oXmlNode Is Nothing Then
                                            oXmlNode = oXml.SelectSingleNode("/BOM/BO")

                                            oXmlNode.AppendChild(oXml.CreateElement("ContactEmployees"))
                                        End If

                                        oXmlNode = oXml.SelectSingleNode("/BOM/BO/ContactEmployees")

                                        oXmlNode.PrependChild(oXmlNode2)
                                    End If
                                Next

                                For j As Integer = oXmlNodes.Count - 1 To 0 Step -1
                                    sInternalCode = ""

                                    Try
                                        sInternalCode = Conexiones.GetValueDB(oDB, "[" & sDBD & "].dbo.[OCPR]", "CntctCode", "Name = '" & oXmlNodes.Item(j).SelectSingleNode("Name").InnerText & "' AND CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "'")

                                        If sInternalCode <> "" Then
                                            oXmlNodes.Item(j).SelectSingleNode("InternalCode").InnerText = sInternalCode
                                        End If
                                    Catch ex As Exception

                                    End Try
                                Next

                                sXML = oXml.OuterXml
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                                If oOCRD.UpdateFromXML(sXML) <> 0 Then
                                    Throw New Exception(oCompanyD.GetLastErrorCode & " / " & oCompanyD.GetLastErrorDescription)
                                End If
                            End If

                            sSQL = "UPDATE [" & sDBD & "].dbo.[OCRD] SET AgentCode = '" & sAgentCode & "' WHERE CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "' "

                            'Actualizamos los medios de comunicación porque por DI API no funciona
                            sSQL &= "UPDATE t2 SET BlockComm = ISNULL(t1.BlockComm, 'N') " & _
                                    "FROM [" & sDBO & "].dbo.[OCPR] t1 WITH (NOLOCK) INNER JOIN " & _
                                    "[" & sDBD & "].dbo.[OCPR] t2 WITH (NOLOCK) ON t1.CardCode = t2.CardCode AND " & _
                                    "t1.Name = t2.Name " & _
                                    "WHERE t2.CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "' "

                            sSQL &= "DELETE t1 FROM [" & sDBD & "].dbo.[CPRC] t1 WITH (NOLOCK) INNER JOIN " & _
                                    "[" & sDBD & "].dbo.[OCPR] t2 WITH (NOLOCK) ON t1.CntctCode = t2.CntctCode " & _
                                    "WHERE t2.CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "' "

                            sSQL &= "INSERT INTO [" & sDBD & "].dbo.[CPRC] " & _
                                    "SELECT t3.CntctCode, t1.CommMeanId, t1.[Select] " & _
                                    "FROM [" & sDBO & "].dbo.[CPRC] t1 WITH (NOLOCK) INNER JOIN " & _
                                    "[" & sDBO & "].dbo.[OCPR] t2 WITH (NOLOCK) ON t1.CntctCode = t2.CntctCode INNER JOIN " & _
                                    "[" & sDBD & "].dbo.[OCPR] t3 WITH (NOLOCK) ON t2.Name = t3.Name AND " & _
                                    "t2.CardCode = t3.CardCode " & _
                                    "WHERE t2.CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "' "

                            'El grupo de empresas al ser un combo que se rellena en tiempo de ejecución, no se puede actualizar el valor
                            'por DI API, por lo que tiene que ser por SQL
                            'If oDt.Rows.Item(i).Item("codeTable2").ToString <> "" Then
                            If sU_EXO_GRUPOEMPRESA <> "" Then
                                sSQL &= "UPDATE [" & sDBD & "].dbo.[OCRD] SET U_EXO_GRUPOEMPRESA = '" & sU_EXO_GRUPOEMPRESA & "' " & _
                                        "WHERE CardCode = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "' "
                            End If
                            'End If

                            oRs.DoQuery(sSQL)
                        End If

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If

                        sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "';"

                        Conexiones.ExecuteSQLDB(oDB, sSQL)

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)

                        If oCompanyD.InTransaction = True Then
                            oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

            If oCompanyD.InTransaction = True Then
                oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)

            If oCompanyD.InTransaction = True Then
                oCompanyD.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If oOCRD IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOCRD)
            If oOCRD2 IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOCRD2)
            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRs)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

    Public Shared Sub OINV()
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oCompanyD As SAPbobsCOM.Company = Nothing
        Dim facturaVentasOrigen As SAPbobsCOM.Documents = Nothing
        Dim facturaComprasDestino As SAPbobsCOM.Documents = Nothing

        Dim oDB As SqlConnection = Nothing
        Dim oDt As System.Data.DataTable = Nothing
        Dim sSQL As String = ""

        Dim sDBO As String = ""
        Dim sDBD As String = ""

        Dim log As EXO_Log.EXO_Log = Nothing

        Try
            log = New EXO_Log.EXO_Log(My.Application.Info.DirectoryPath.ToString & "\Logs\Log_ERRORES_OINV.txt", 1)

            Conexiones.Connect_SQLServer(oDB, log)
            sSQL = "SELECT t1.dbNameOrig, t1.dbNameDest, t1.tableName, t1.codeTable " & _
              "FROM [INTERCOMPANY].dbo.[REPLICATE] t1 WITH (NOLOCK) " & _
              "WHERE t1.tableName = 'OINV' " & _
              "ORDER BY t1.dbNameOrig, t1.dbNameDest "

            oDt = New System.Data.DataTable
            Conexiones.FillDtDB(oDB, oDt, sSQL)

            If oDt.Rows.Count > 0 Then
                sDBO = oDt.Rows.Item(0).Item("dbNameOrig").ToString
                sDBD = oDt.Rows.Item(0).Item("dbNameDest").ToString

                Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(0).Item("dbNameOrig").ToString)
                Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(0).Item("dbNameDest").ToString)

                oCompanyO.XMLAsString = True
                oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                oCompanyD.XMLAsString = True
                oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                For i = 0 To oDt.Rows.Count - 1
                    Try
                        If sDBO <> oDt.Rows.Item(i).Item("dbNameOrig").ToString Then
                            'Desconectar Company Origen y volver a conectar con la nueva Company Origen
                            Conexiones.Disconnect_Company(oCompanyO)

                            Conexiones.Connect_Company(oCompanyO, oDt.Rows.Item(i).Item("dbNameOrig").ToString)

                            oCompanyO.XMLAsString = True
                            oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                            sDBO = oDt.Rows.Item(i).Item("dbNameOrig").ToString
                        End If

                        If sDBD <> oDt.Rows.Item(i).Item("dbNameDest").ToString Then
                            'Desconectar Company Destino y volver a conectar con la nueva Company Destino
                            Conexiones.Disconnect_Company(oCompanyD)

                            Conexiones.Connect_Company(oCompanyD, oDt.Rows.Item(i).Item("dbNameDest").ToString)

                            oCompanyD.XMLAsString = True
                            oCompanyD.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                            sDBD = oDt.Rows.Item(i).Item("dbNameDest").ToString
                        End If
                        'Replicado de la factura.
                        facturaVentasOrigen = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices), SAPbobsCOM.Documents)

                        If facturaVentasOrigen.GetByKey(CInt(oDt.Rows.Item(i).Item("codeTable").ToString)) Then
                            'Busqueda del proveedor.
                            Dim oRs As SAPbobsCOM.Recordset = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                            oRs.DoQuery("SELECT ISNULL(t1.U_EXO_GRUPOEMPRESA, '') U_EXO_GRUPOEMPRESA FROM [" + sDBO + "].dbo.[OADM] t1 WITH (NOLOCK) ")
                            If oRs.RecordCount > 0 Then
                                Dim grupoEmpresa As String = oRs.Fields.Item(0).Value
                                If grupoEmpresa.Trim <> "" Then
                                    oRs.DoQuery("SELECT CardCode FROM OCRD WHERE CardType = 'S' AND U_EXO_GRUPOEMPRESA = '" + grupoEmpresa + "'")
                                    If oRs.RecordCount > 0 Then
                                        Dim codigoProveedor As String = oRs.Fields.Item(0).Value
                                        'Copiamos los valores
                                        'Cabecera
                                        facturaComprasDestino = CType(oCompanyD.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices), SAPbobsCOM.Documents)
                                        facturaComprasDestino.CardCode = codigoProveedor
                                        facturaComprasDestino.DocDate = facturaVentasOrigen.DocDate
                                        facturaComprasDestino.DocDueDate = facturaVentasOrigen.DocDueDate
                                        facturaComprasDestino.TaxDate = facturaVentasOrigen.TaxDate
                                        facturaComprasDestino.NumAtCard = facturaVentasOrigen.DocNum
                                        If facturaVentasOrigen.DiscountPercent <> 0 Then
                                            facturaComprasDestino.DiscountPercent = facturaVentasOrigen.DiscountPercent
                                        End If
                                        facturaComprasDestino.Comments = facturaVentasOrigen.Comments
                                        facturaComprasDestino.Lines.FreeText = facturaVentasOrigen.Lines.FreeText
                                        'Lineas
                                        For indiceLineas As Integer = 0 To facturaVentasOrigen.Lines.Count - 1
                                            If indiceLineas <> 0 Then
                                                facturaComprasDestino.Lines.Add()
                                            End If
                                            facturaVentasOrigen.Lines.SetCurrentLine(indiceLineas)
                                            If facturaVentasOrigen.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items Then
                                                facturaComprasDestino.Lines.ItemCode = facturaVentasOrigen.Lines.ItemCode
                                                facturaComprasDestino.Lines.Quantity = facturaVentasOrigen.Lines.Quantity
                                                facturaComprasDestino.Lines.UnitPrice = facturaVentasOrigen.Lines.UnitPrice
                                            Else
                                                facturaComprasDestino.Lines.AccountCode = facturaVentasOrigen.Lines.AccountCode
                                                facturaComprasDestino.Lines.LineTotal = facturaVentasOrigen.Lines.LineTotal
                                            End If
                                            If facturaVentasOrigen.Lines.DiscountPercent <> 0 Then
                                                facturaComprasDestino.Lines.DiscountPercent = facturaVentasOrigen.Lines.DiscountPercent
                                            End If
                                            'Verficicacion normas de reparto.
                                            'Norma 1
                                            Dim normaOrigen As String = facturaVentasOrigen.Lines.CostingCode
                                            If normaOrigen.Trim <> "" Then
                                                oRs.DoQuery("SELECT * FROM OOCR WHERE OcrCode = '" + normaOrigen + "'")
                                                If oRs.RecordCount = 0 Then
                                                    'Sera manual, la creamos
                                                    If normaOrigen.StartsWith("M") Then
                                                        'Calculo  del siguiente numero
                                                        oRs.DoQuery("select ISNULL(MAX(cast(SUBSTRING(OcrCode,2,LEN(OcrCode)-1) as int)),0)+1 from OMDR")
                                                        Dim siguienteNumero As Integer = CInt(oRs.Fields.Item(0).Value)

                                                        sSQL = "INSERT OMDR(OcrCode,OcrName,OcrTotal,Direct,Locked,DataSource,UserSign,DimCode,AbsEntry,Active,logInstanc,UserSign2,updateDate) "
                                                        sSQL += "SELECT 'M" + siguienteNumero.ToString("0000000") + "',OcrName,OcrTotal,Direct,Locked,DataSource,UserSign,DimCode,(SELECT AutoKey FROM ONNM WHERE ObjectCode = '252'),Active,logInstanc,UserSign2,updateDate "
                                                        sSQL += "FROM [" + sDBO + "].dbo.OMDR WHERE OcrCode = '" + normaOrigen + "'"
                                                        oRs.DoQuery(sSQL)


                                                        'Codigo de norma de reparto manual = 252
                                                        sSQL = "UPDATE ONNM SET AutoKey=Autokey + 1 WHERE ObjectCode = '252'"
                                                        oRs.DoQuery(sSQL)

                                                        sSQL = "INSERT MDR1(OcrCode,PrcCode,PrcAmount,OcrTotal,Direct,UserSign,ValidFrom,ValidTo,logInstanc,UserSign2,updateDate) "
                                                        sSQL += "SELECT 'M" + siguienteNumero.ToString("0000000") + "',PrcCode,PrcAmount,OcrTotal,Direct,UserSign,ValidFrom,ValidTo,logInstanc,UserSign2,updateDate "
                                                        sSQL += "FROM [" + sDBO + "].dbo.MDR1 WHERE OcrCode = '" + normaOrigen + "'"
                                                        oRs.DoQuery(sSQL)

                                                        normaOrigen = "M" + siguienteNumero.ToString("0000000")

                                                    End If
                                                End If
                                                facturaComprasDestino.Lines.CostingCode = normaOrigen
                                            End If
                                            'Norma 2
                                            normaOrigen = facturaVentasOrigen.Lines.CostingCode2
                                            If normaOrigen.Trim <> "" Then
                                                oRs.DoQuery("SELECT * FROM OOCR WHERE OcrCode = '" + normaOrigen + "'")
                                                If oRs.RecordCount = 0 Then
                                                    'Sera manual, la creamos
                                                    If normaOrigen.StartsWith("M") Then
                                                        'Calculo  del siguiente numero
                                                        oRs.DoQuery("select ISNULL(MAX(cast(SUBSTRING(OcrCode,2,LEN(OcrCode)-1) as int)),0)+1 from OMDR")
                                                        Dim siguienteNumero As Integer = CInt(oRs.Fields.Item(0).Value)

                                                        sSQL = "INSERT OMDR(OcrCode,OcrName,OcrTotal,Direct,Locked,DataSource,UserSign,DimCode,AbsEntry,Active,logInstanc,UserSign2,updateDate) "
                                                        sSQL += "SELECT 'M" + siguienteNumero.ToString("0000000") + "',OcrName,OcrTotal,Direct,Locked,DataSource,UserSign,DimCode,(SELECT AutoKey FROM ONNM WHERE ObjectCode = '252'),Active,logInstanc,UserSign2,updateDate "
                                                        sSQL += "FROM [" + sDBO + "].dbo.OMDR WHERE OcrCode = '" + normaOrigen + "'"
                                                        oRs.DoQuery(sSQL)


                                                        'Codigo de norma de reparto manual = 252
                                                        sSQL = "UPDATE ONNM SET AutoKey=Autokey + 1 WHERE ObjectCode = '252'"
                                                        oRs.DoQuery(sSQL)

                                                        sSQL = "INSERT MDR1(OcrCode,PrcCode,PrcAmount,OcrTotal,Direct,UserSign,ValidFrom,ValidTo,logInstanc,UserSign2,updateDate) "
                                                        sSQL += "SELECT 'M" + siguienteNumero.ToString("0000000") + "',PrcCode,PrcAmount,OcrTotal,Direct,UserSign,ValidFrom,ValidTo,logInstanc,UserSign2,updateDate "
                                                        sSQL += "FROM [" + sDBO + "].dbo.MDR1 WHERE OcrCode = '" + normaOrigen + "'"
                                                        oRs.DoQuery(sSQL)

                                                        normaOrigen = "M" + siguienteNumero.ToString("0000000")

                                                    End If
                                                End If
                                                facturaComprasDestino.Lines.CostingCode2 = normaOrigen
                                            End If
                                            'Norma 3
                                            normaOrigen = facturaVentasOrigen.Lines.CostingCode3
                                            If normaOrigen.Trim <> "" Then
                                                oRs.DoQuery("SELECT * FROM OOCR WHERE OcrCode = '" + normaOrigen + "'")
                                                If oRs.RecordCount = 0 Then
                                                    'Sera manual, la creamos
                                                    If normaOrigen.StartsWith("M") Then
                                                        'Calculo  del siguiente numero
                                                        oRs.DoQuery("select ISNULL(MAX(cast(SUBSTRING(OcrCode,2,LEN(OcrCode)-1) as int)),0)+1 from OMDR")
                                                        Dim siguienteNumero As Integer = CInt(oRs.Fields.Item(0).Value)

                                                        sSQL = "INSERT OMDR(OcrCode,OcrName,OcrTotal,Direct,Locked,DataSource,UserSign,DimCode,AbsEntry,Active,logInstanc,UserSign2,updateDate) "
                                                        sSQL += "SELECT 'M" + siguienteNumero.ToString("0000000") + "',OcrName,OcrTotal,Direct,Locked,DataSource,UserSign,DimCode,(SELECT AutoKey FROM ONNM WHERE ObjectCode = '252'),Active,logInstanc,UserSign2,updateDate "
                                                        sSQL += "FROM [" + sDBO + "].dbo.OMDR WHERE OcrCode = '" + normaOrigen + "'"
                                                        oRs.DoQuery(sSQL)


                                                        'Codigo de norma de reparto manual = 252
                                                        sSQL = "UPDATE ONNM SET AutoKey=Autokey + 1 WHERE ObjectCode = '252'"
                                                        oRs.DoQuery(sSQL)

                                                        sSQL = "INSERT MDR1(OcrCode,PrcCode,PrcAmount,OcrTotal,Direct,UserSign,ValidFrom,ValidTo,logInstanc,UserSign2,updateDate) "
                                                        sSQL += "SELECT 'M" + siguienteNumero.ToString("0000000") + "',PrcCode,PrcAmount,OcrTotal,Direct,UserSign,ValidFrom,ValidTo,logInstanc,UserSign2,updateDate "
                                                        sSQL += "FROM [" + sDBO + "].dbo.MDR1 WHERE OcrCode = '" + normaOrigen + "'"
                                                        oRs.DoQuery(sSQL)

                                                        normaOrigen = "M" + siguienteNumero.ToString("0000000")

                                                    End If
                                                End If
                                                facturaComprasDestino.Lines.CostingCode3 = normaOrigen
                                            End If
                                        Next
                                        'Añadimos el documento
                                        If facturaComprasDestino.Add() = 0 Then
                                            sSQL = "DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE dbNameOrig = '" & sDBO & "' AND dbNameDest = '" & sDBD & "' AND tableName = '" & oDt.Rows.Item(i).Item("tableName").ToString & "' AND codeTable = '" & oDt.Rows.Item(i).Item("codeTable").ToString & "';"
                                            Conexiones.ExecuteSQLDB(oDB, sSQL)
                                        Else
                                            log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- Error al crear la factura: " + oCompanyD.GetLastErrorDescription, EXO_Log.EXO_Log.Tipo.error)
                                        End If
                                    Else
                                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- No esta definido el proveedor asociado a la empresa origen.", EXO_Log.EXO_Log.Tipo.error)
                                    End If
                                Else
                                    log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- No esta definido el grupo de empresa en el origen.", EXO_Log.EXO_Log.Tipo.error)
                                End If
                            End If
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
                        Else
                            log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- No esite el documento origen.", EXO_Log.EXO_Log.Tipo.error)
                        End If

                        If facturaVentasOrigen IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(facturaVentasOrigen)
                        If facturaComprasDestino IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(facturaComprasDestino)

                        'Fin replicado de la factura.
                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
                    Catch ex As Exception
                        log.escribeMensaje("-- " & sDBO & "|" & sDBD & "|" & oDt.Rows.Item(i).Item("tableName").ToString & "|" & oDt.Rows.Item(i).Item("codeTable").ToString & " -- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)
                    End Try
                Next i
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            log.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            log.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDt IsNot Nothing Then oDt.Dispose()
            If facturaVentasOrigen IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(facturaVentasOrigen)
            If facturaComprasDestino IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(facturaComprasDestino)

            Conexiones.Disconnect_SQLServer(oDB)
            Conexiones.Disconnect_Company(oCompanyO)
            Conexiones.Disconnect_Company(oCompanyD)
        End Try
    End Sub

End Class
