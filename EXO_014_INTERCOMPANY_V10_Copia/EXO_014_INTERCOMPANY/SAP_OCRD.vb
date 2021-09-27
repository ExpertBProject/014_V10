Imports SAPbouiCOM
Public Class SAP_OCRD
    Inherits EXO_UIAPI.EXO_DLLBase

#Region "Constructor"

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaDatos()
        End If
    End Sub

#End Region

#Region "Inicialización"

    Private Sub cargaDatos()
        Dim oXML As String = ""

        Try
            If objglobal.refDi.comunes.esAdministrador Then
                'Campos de Usuario para configuración de InterCompany
                oXML = objglobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OCRD.xml")
                objglobal.SboApp.StatusBar.SetText("Validando: UDFs ICs", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            objglobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objglobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
        End Try

    End Sub

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objglobal.funciones.leerEmbebido(Me.GetType(), "Filtros_OCRD.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function

#End Region

#Region "Eventos"

    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
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
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        End Try
    End Function

    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
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
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function

    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

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

    Private Function EventHandler_ComboSelect_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ComboSelect_After = False

        Try
            oForm = objglobal.SboApp.Forms.Item(pVal.FormUID)

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
            oRs = CType(objglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsAux = CType(objglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs = objGlobal.compañia.GetCompanyList

            While Not oRs.EoF
                'Si las compañías del company list tienen los siguientes dos campos entonces cargamos el grupo de empresas
                sSQL = "SELECT COL.Name " &
                        "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.syscolumns COL WITH (NOLOCK) INNER JOIN " &
                        "[" & oRs.Fields.Item(0).Value.ToString & "].dbo.sysobjects OBJ WITH (NOLOCK) ON OBJ.id = COL.id " &
                        "WHERE OBJ.name = 'OADM' " &
                        "AND COL.name = 'U_EXO_CONSOLIDACION' " &
                        "UNION ALL " &
                        "SELECT COL.Name " &
                        "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.syscolumns COL WITH (NOLOCK) INNER JOIN " &
                        "[" & oRs.Fields.Item(0).Value.ToString & "].dbo.sysobjects OBJ WITH (NOLOCK) ON OBJ.id = COL.id " &
                        "WHERE OBJ.name = 'OADM' " &
                        "AND COL.name = 'U_EXO_MATRIZ' " &
                        "UNION ALL " &
                        "SELECT COL.Name " &
                        "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.syscolumns COL WITH (NOLOCK) INNER JOIN " &
                        "[" & oRs.Fields.Item(0).Value.ToString & "].dbo.sysobjects OBJ WITH (NOLOCK) ON OBJ.id = COL.id " &
                        "WHERE OBJ.name = 'OADM' " &
                        "AND COL.name = 'U_EXO_GRUPOEMPRESA'"

                oRsAux.DoQuery(sSQL)

                If oRsAux.RecordCount = 3 Then
                    'Sólo se carga el combo si la compañia conectada es Matriz o Sucursal
                    sSQL = "SELECT ISNULL(t1.U_EXO_GRUPOEMPRESA, '') U_EXO_GRUPOEMPRESA " &
                           "FROM [" & objGlobal.compañia.CompanyDB & "].dbo.[OADM] t1 WITH (NOLOCK) " &
                           "WHERE (ISNULL(t1.U_EXO_MATRIZ, 'N') = 'Y' " &
                           "OR (ISNULL(t1.U_EXO_CONSOLIDACION, 'N') = 'N' " &
                           "AND ISNULL(t1.U_EXO_MATRIZ, 'N') = 'N')) "

                    oRsAux.DoQuery(sSQL)

                    If oRsAux.RecordCount > 0 Then
                        'Combo Grupo de empresas
                        sSQL = "SELECT ISNULL(t1.U_EXO_GRUPOEMPRESA, '') U_EXO_GRUPOEMPRESA " &
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
            If EXO_GLOBALES.EmpresaConectadaEsMatriz(objglobal) = True Then
                objGlobal.SBOApp.StatusBar.SetText("Guardando datos para InterCompany ... Espere por favor ...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                oRs = CType(objglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                oRsAux = CType(objglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                oRsAux2 = CType(objglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                oRs.DoQuery("SELECT dbName " &
                            "FROM [INTERCOMPANY].dbo.[DATABASES] WITH (NOLOCK) " &
                            "WHERE dbTipo = 'S' " &
                            "AND dbName <> '" & objGlobal.compañia.CompanyDB & "'")

                oXml.LoadXml(oRs.GetAsXML())
                oNodes = oXml.SelectNodes("//row")

                If oRs.RecordCount > 0 Then
                    For i As Integer = 0 To oNodes.Count - 1
                        oNode = oNodes.Item(i)

                        oRsAux.DoQuery("SELECT CardCode AS Codigo, ISNULL(U_EXO_GRUPOEMPRESA, '-') AS U_EXO_GRUPOEMPRESA " &
                                       "FROM " & sTableName & " WITH (NOLOCK) " &
                                       "WHERE CardCode = '" & sCodigo & "'")

                        If oRsAux.RecordCount > 0 Then
                            oRsAux2.DoQuery("SELECT dbNameOrig " &
                                            "FROM [INTERCOMPANY].dbo.[REPLICATE] WITH (NOLOCK) " &
                                            "WHERE dbNameOrig = '" & objGlobal.compañia.CompanyDB & "' " &
                                            "AND dbNameDest = '" & oNode.SelectSingleNode("dbName").InnerText & "' " &
                                            "AND tableCategory = " & sTableCategory & " " &
                                            "AND tableName = '" & sTableName & "' " &
                                            "AND codeTable = '" & sCodigo & "'")

                            'If oRsAux.Fields.Item("U_EXO_GRUPOEMPRESA").Value.ToString = "-" Then
                            If oRsAux2.RecordCount = 0 Then
                                Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day)
                                Dim sSQL As String = "INSERT INTO [INTERCOMPANY].dbo.[REPLICATE] (dbNameOrig, dbNameDest, tableCategory, tableName, codeTable, dateAdd) VALUES " &
                                                "('" & objGlobal.compañia.CompanyDB & "', '" & oNode.SelectSingleNode("dbName").InnerText & "' " &
                                                ", " & sTableCategory & ", '" & sTableName & "', '" & sCodigo & "', '" & dFecha.ToShortDateString & "')"
                                oRsAux2.DoQuery(sSQL)
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
                            '                                ", " & sTableCategory & ", '" & sTableName & "', '" & sCodigo & "', '" & sGrupoEmpresaConectada & "', '" & dFecha.ToShortDateString & "')")
                            '            End If
                            '        End If
                            '    End If
                            'End If
                        Else
                            oRsAux2.DoQuery("DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE tableName = '" & sTableName & "' AND codeTable = '" & sCodigo & "'")
                        End If
                    Next
                End If

                objGlobal.SBOApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
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
