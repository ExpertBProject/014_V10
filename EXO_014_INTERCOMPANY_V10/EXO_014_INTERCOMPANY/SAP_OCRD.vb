Public Class SAP_OCRD
    Inherits EXO_Generales.EXO_DLLBase

#Region "Constructor"

    Public Sub New(ByRef generales As EXO_Generales.EXO_General, ByRef actualizar As Boolean)
        MyBase.New(generales, actualizar)

        If actualizar Then
            cargaDatos()
        End If
    End Sub

#End Region

#Region "Inicialización"

    Private Sub cargaDatos()
        Dim oXML As String = ""
        Dim udoObj As EXO_Generales.EXO_UDO = Nothing

        Try
            If objGlobal.conexionSAP.esAdministrador Then
                'Campos de Usuario para configuración de InterCompany
                oXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "UDFs_OCRD.xml")
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("Validando: UDFs ICs", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objGlobal.conexionSAP.LoadBDFromXML(oXML)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(udoObj, Object))
        End Try

    End Sub

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.Functions.leerEmbebido(Me.GetType(), "Filtros_OCRD.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function

#End Region

#Region "Eventos"

    Public Overrides Function SBOApp_ItemEvent(ByRef infoEvento As EXO_Generales.EXO_infoItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "134"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_ComboSelect_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select

                    End Select

                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "134"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            End Select

                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "134"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "134"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select

                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function

    Public Overrides Function SBOApp_FormDataEvent(ByRef infoEvento As EXO_Generales.EXO_BusinessObjectInfo) As Boolean
        Dim oXml As New Xml.XmlDocument
        Dim sCodigo As String = ""

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "134"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select

                End Select

            Else
                Select Case infoEvento.FormTypeEx
                    Case "134"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess Then
                                    oXml.LoadXml(infoEvento.ObjectKey)
                                    sCodigo = oXml.SelectSingleNode("BusinessPartnerParams/CardCode").InnerText

                                    If GuardarInterCoOCRD("6", "OCRD", sCodigo) = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then
                                    oXml.LoadXml(infoEvento.ObjectKey)
                                    sCodigo = oXml.SelectSingleNode("BusinessPartnerParams/CardCode").InnerText

                                    If GuardarInterCoOCRD("6", "OCRD", sCodigo) = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE
                                If infoEvento.ActionSuccess Then
                                    oXml.LoadXml(infoEvento.ObjectKey)
                                    sCodigo = oXml.SelectSingleNode("BusinessPartnerParams/CardCode").InnerText

                                    If GuardarInterCoOCRD("6", "OCRD", sCodigo) = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select

                End Select

            End If

            Return MyBase.SBOApp_FormDataEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function

    Private Function EventHandler_Form_Load(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = SboApp.Forms.Item(pVal.FormUID)

            If CargarCombos(oForm) = False Then
                Exit Function
            End If

            EventHandler_Form_Load = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_ComboSelect_After(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ComboSelect_After = False

        Try
            oForm = SboApp.Forms.Item(pVal.FormUID)

            If pVal.ActionSuccess = True Then
                If pVal.ItemUID = "178" Then
                    If pVal.ColUID = "7" Then
                        SAP_OCST._sCountry = CType(CType(oForm.Items.Item("178").Specific, SAPbouiCOM.Matrix).Columns.Item("8").Cells.Item(1).Specific, SAPbouiCOM.ComboBox).Selected.Value
                    End If
                End If
            End If

            EventHandler_ComboSelect_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

#End Region

#Region "Métodos auxiliares"

    Private Function CargarCombos(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux As SAPbobsCOM.Recordset = Nothing

        CargarCombos = False

        Try
            oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsAux = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs = objGlobal.conexionSAP.compañia.GetCompanyList

            While Not oRs.EoF
                'Si las compañías del company list tienen los siguientes dos campos entonces cargamos el grupo de empresas
                sSQL = "SELECT COL.Name " & _
                        "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.syscolumns COL WITH (NOLOCK) INNER JOIN " & _
                        "[" & oRs.Fields.Item(0).Value.ToString & "].dbo.sysobjects OBJ WITH (NOLOCK) ON OBJ.id = COL.id " & _
                        "WHERE OBJ.name = 'OADM' " & _
                        "AND COL.name = 'U_EXO_CONSOLIDACION' " & _
                        "UNION ALL " & _
                        "SELECT COL.Name " & _
                        "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.syscolumns COL WITH (NOLOCK) INNER JOIN " & _
                        "[" & oRs.Fields.Item(0).Value.ToString & "].dbo.sysobjects OBJ WITH (NOLOCK) ON OBJ.id = COL.id " & _
                        "WHERE OBJ.name = 'OADM' " & _
                        "AND COL.name = 'U_EXO_MATRIZ' " & _
                        "UNION ALL " & _
                        "SELECT COL.Name " & _
                        "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.syscolumns COL WITH (NOLOCK) INNER JOIN " & _
                        "[" & oRs.Fields.Item(0).Value.ToString & "].dbo.sysobjects OBJ WITH (NOLOCK) ON OBJ.id = COL.id " & _
                        "WHERE OBJ.name = 'OADM' " & _
                        "AND COL.name = 'U_EXO_GRUPOEMPRESA'"

                oRsAux.DoQuery(sSQL)

                If oRsAux.RecordCount = 3 Then
                    'Sólo se carga el combo si la compañia conectada es Matriz o Sucursal
                    sSQL = "SELECT ISNULL(t1.U_EXO_GRUPOEMPRESA, '') U_EXO_GRUPOEMPRESA " & _
                           "FROM [" & objGlobal.conexionSAP.compañia.CompanyDB & "].dbo.[OADM] t1 WITH (NOLOCK) " & _
                           "WHERE (ISNULL(t1.U_EXO_MATRIZ, 'N') = 'Y' " & _
                           "OR (ISNULL(t1.U_EXO_CONSOLIDACION, 'N') = 'N' " & _
                           "AND ISNULL(t1.U_EXO_MATRIZ, 'N') = 'N')) "

                    oRsAux.DoQuery(sSQL)

                    If oRsAux.RecordCount > 0 Then
                        'Combo Grupo de empresas
                        sSQL = "SELECT ISNULL(t1.U_EXO_GRUPOEMPRESA, '') U_EXO_GRUPOEMPRESA " & _
                               "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.[OADM] t1 WITH (NOLOCK) "

                        oRsAux.DoQuery(sSQL)

                        If oRsAux.RecordCount > 0 Then
                            Try
                                CType(oForm.Items.Item("U_EXO_GRUPOEMPRESA").Specific, SAPbouiCOM.ComboBox).ValidValues.Add(oRsAux.Fields.Item("U_EXO_GRUPOEMPRESA").Value.ToString, oRsAux.Fields.Item("U_EXO_GRUPOEMPRESA").Value.ToString)
                            Catch exCOM As System.Runtime.InteropServices.COMException
                            Catch ex As Exception
                            End Try
                        End If
                    End If
                End If

                oRs.MoveNext()
            End While

            CargarCombos = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux, Object))
        End Try
    End Function

    Private Function GuardarInterCoOCRD(ByVal sTableCategory As String, ByVal sTableName As String, ByVal sCodigo As String) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux2 As SAPbobsCOM.Recordset = Nothing
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        'Dim sGrupoEmpresaConectada As String = ""

        GuardarInterCoOCRD = False

        Try
            If EXO_GLOBALES.EmpresaConectadaEsMatriz(objGlobal) = True Then
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("Guardando datos para InterCompany ... Espere por favor ...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                oRsAux = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                oRsAux2 = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                oRs.DoQuery("SELECT dbName " & _
                            "FROM [INTERCOMPANY].dbo.[DATABASES] WITH (NOLOCK) " & _
                            "WHERE dbTipo = 'S' " & _
                            "AND dbName <> '" & objGlobal.conexionSAP.compañia.CompanyDB & "'")

                oXml.LoadXml(oRs.GetAsXML())
                oNodes = oXml.SelectNodes("//row")

                If oRs.RecordCount > 0 Then
                    For i As Integer = 0 To oNodes.Count - 1
                        oNode = oNodes.Item(i)

                        oRsAux.DoQuery("SELECT CardCode AS Codigo, ISNULL(U_EXO_GRUPOEMPRESA, '-') AS U_EXO_GRUPOEMPRESA " & _
                                       "FROM " & sTableName & " WITH (NOLOCK) " & _
                                       "WHERE CardCode = '" & sCodigo & "'")

                        If oRsAux.RecordCount > 0 Then
                            oRsAux2.DoQuery("SELECT dbNameOrig " & _
                                            "FROM [INTERCOMPANY].dbo.[REPLICATE] WITH (NOLOCK) " & _
                                            "WHERE dbNameOrig = '" & objGlobal.conexionSAP.compañia.CompanyDB & "' " & _
                                            "AND dbNameDest = '" & oNode.SelectSingleNode("dbName").InnerText & "' " & _
                                            "AND tableCategory = " & sTableCategory & " " & _
                                            "AND tableName = '" & sTableName & "' " & _
                                            "AND codeTable = '" & sCodigo & "'")

                            'If oRsAux.Fields.Item("U_EXO_GRUPOEMPRESA").Value.ToString = "-" Then
                            If oRsAux2.RecordCount = 0 Then
                                oRsAux2.DoQuery("INSERT INTO [INTERCOMPANY].dbo.[REPLICATE] (dbNameOrig, dbNameDest, tableCategory, tableName, codeTable, dateAdd) VALUES " & _
                                                "('" & objGlobal.conexionSAP.compañia.CompanyDB & "', '" & oNode.SelectSingleNode("dbName").InnerText & "' " & _
                                                ", " & sTableCategory & ", '" & sTableName & "', '" & sCodigo & "', '" & Now.Year & "-" & Right("0" & Now.Month.ToString, 2) & "-" & Right("0" & Now.Day.ToString, 2) & "')")
                            End If
                            'Else
                            '    oRsAux.DoQuery("SELECT ISNULL(U_EXO_GRUPOEMPRESA, '') U_EXO_GRUPOEMPRESA " & _
                            '                   "FROM [" & oNode.SelectSingleNode("dbName").InnerText & "].dbo.[OADM] WITH (NOLOCK) " & _
                            '                   "WHERE ISNULL(U_EXO_GRUPOEMPRESA, '') = '" & oRsAux.Fields.Item("U_EXO_GRUPOEMPRESA").Value.ToString & "'")

                            '    If oRsAux.RecordCount > 0 Then
                            '        If oRsAux2.RecordCount = 0 Then
                            '            oRsAux.DoQuery("SELECT ISNULL(U_EXO_GRUPOEMPRESA, '') U_EXO_GRUPOEMPRESA " & _
                            '                           "FROM [OADM] WITH (NOLOCK)")

                            '            If oRsAux.RecordCount > 0 Then
                            '                sGrupoEmpresaConectada = oRsAux.Fields.Item("U_EXO_GRUPOEMPRESA").Value.ToString

                            '                oRsAux2.DoQuery("INSERT INTO [INTERCOMPANY].dbo.[REPLICATE] (dbNameOrig, dbNameDest, tableCategory, tableName, codeTable, codeTable2, dateAdd) VALUES " & _
                            '                                "('" & objGlobal.conexionSAP.compañia.CompanyDB & "', '" & oNode.SelectSingleNode("dbName").InnerText & "' " & _
                            '                                ", " & sTableCategory & ", '" & sTableName & "', '" & sCodigo & "', '" & sGrupoEmpresaConectada & "', '" & Now.Year & "-" & Right("0" & Now.Month.ToString, 2) & "-" & Right("0" & Now.Day.ToString, 2) & "')")
                            '            End If
                            '        End If
                            '    End If
                            'End If
                        Else
                            oRsAux2.DoQuery("DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE tableName = '" & sTableName & "' AND codeTable = '" & sCodigo & "'")
                        End If
                    Next
                End If

                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            End If

            GuardarInterCoOCRD = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux2, Object))
        End Try
    End Function

#End Region

End Class
