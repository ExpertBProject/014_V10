Public Class EXO_OICO
    Inherits EXO_Generales.EXO_DLLBase

#Region "Variables globales"

    Private Shared _sDocEntry As String

#End Region

#Region "Constructor"

    Public Sub New(ByRef generales As EXO_Generales.EXO_General, actualizar As Boolean)
        MyBase.New(generales, actualizar)

        If actualizar Then
            cargaDatos()
            cargaAutorizaciones()
        End If
    End Sub

#End Region

#Region "Inicialización"

    Private Sub cargaDatos()
        Dim oXML As String = ""
        Dim udoObj As EXO_Generales.EXO_UDO = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        If objGlobal.conexionSAP.esAdministrador Then
            Try
                oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                oRs.DoQuery("SELECT CompnyName FROM OADM WITH (NOLOCK) WHERE ISNULL(U_EXO_CONSOLIDACION, 'N') = 'Y'")

                'Sólo generamos el UDO en las empresas de Consolidación
                If oRs.RecordCount > 0 Then
                    EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

                    'UDO Configuración InterCompany
                    oXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "UDO_EXO_OICO.xml")
                    udoObj = New EXO_Generales.EXO_UDO("EXO_OICO", objGlobal)
                    objGlobal.conexionSAP.SBOApp.StatusBar.SetText("Validando: UDO EXO_OICO", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    udoObj.validaObjeto(oXML)
                End If

            Catch exCOM As System.Runtime.InteropServices.COMException
                objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Catch ex As Exception
                objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Finally
                EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
                EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(udoObj, Object))
            End Try
        End If
    End Sub

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.Functions.leerEmbebido(Me.GetType(), "Filtros_EXO_OICO.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Dim menuXML As String = ""
        Dim res As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        Try
            oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs.DoQuery("SELECT CompnyName FROM OADM WITH (NOLOCK) WHERE ISNULL(U_EXO_CONSOLIDACION, 'N') = 'Y'")

            'Sólo cargamos el menú en las empresas de Consolidación
            If oRs.RecordCount > 0 Then
                menuXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "EXO_MENUINTER.xml")
                SboApp.LoadBatchActions(menuXML)
                res = SboApp.GetLastBatchResults
            End If

            Return Nothing

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)

            Return Nothing
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)

            Return Nothing
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

    'Para definir autorizaciones
    Private Sub cargaAutorizaciones()
        Dim autorizacionXML As String = ""
        Dim res As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        Try
            oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs.DoQuery("SELECT CompnyName FROM OADM WITH (NOLOCK) WHERE ISNULL(U_EXO_CONSOLIDACION, 'N') = 'Y'")

            'Sólo creamos la autorización en las empresas de Consolidación
            If oRs.RecordCount > 0 Then
                autorizacionXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "EXO_AUINTER.xml")
                Me.objGlobal.conexionSAP.refCompañia.LoadBDFromXML(autorizacionXML)
                res = SboApp.GetLastBatchResults
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub

#End Region

#Region "Eventos"

    Public Overrides Function SBOApp_MenuEvent(ByRef infoEvento As EXO_Generales.EXO_MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
                oForm = SboApp.Forms.ActiveForm

                Select Case oForm.TypeEx
                    Case "169"

                        Select Case infoEvento.MenuUID
                            Case "EXO-MnInterCo"
                                objGlobal.conexionSAP.cargaFormUdoBD("EXO_OICO")

                        End Select

                End Select

            Else
                oForm = SboApp.Forms.ActiveForm

                Select Case oForm.TypeEx
                    Case "UDO_FT_EXO_OICO"

                        Select Case infoEvento.MenuUID
                            Case "1282" 'Nuevo
                                If InicializarValores(oForm) = False Then
                                    Return False
                                End If

                        End Select

                End Select

            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Public Overrides Function SBOApp_FormDataEvent(ByRef infoEvento As EXO_Generales.EXO_BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument
        Dim sDocEntry As String = ""

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_OICO"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select

                End Select

            Else
                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_OICO"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess Then
                                    oXml.LoadXml(infoEvento.ObjectKey)
                                    sDocEntry = oXml.SelectSingleNode("ConfInteParams/DocEntry").InnerText

                                    oForm = SboApp.Forms.Item(infoEvento.FormUID)

                                    If AddDatabasesInterCompany(oForm, sDocEntry) = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then
                                    oXml.LoadXml(infoEvento.ObjectKey)
                                    sDocEntry = oXml.SelectSingleNode("ConfInteParams/DocEntry").InnerText

                                    'Para poder cargar la configuración de InterCompany una vez añadida
                                    _sDocEntry = sDocEntry

                                    oForm = SboApp.Forms.Item(infoEvento.FormUID)

                                    If AddDatabasesInterCompany(oForm, sDocEntry) = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

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
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Public Overrides Function SBOApp_ItemEvent(ByRef infoEvento As EXO_Generales.EXO_infoItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_OICO"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select

                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_OICO"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select

                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_OICO"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_OICO"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

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

    Private Function EventHandler_Form_Visible(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing

        EventHandler_Form_Visible = False

        Try
            If pVal.ActionSuccess = True Then
                'Recuperar el formulario
                oForm = Me.SboApp.Forms.Item(pVal.FormUID)

                oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                oRs.DoQuery("SELECT DocEntry FROM [@EXO_OICO] WITH (NOLOCK)")

                If oRs.RecordCount > 0 Then
                    oForm.DataSources.DBDataSources.Item("@EXO_OICO").Offset = 0

                    oConds = New SAPbouiCOM.Conditions
                    oCond = oConds.Add
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCond.Alias = "DocEntry"
                    oCond.CondVal = oRs.Fields.Item("DocEntry").Value.ToString

                    oForm.DataSources.DBDataSources.Item("@EXO_OICO").Query(oConds)
                    oForm.DataSources.DBDataSources.Item("@EXO_ICO1").Query(oConds)

                    If oForm.Visible = True Then
                        CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSource()
                    End If

                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                End If

                If oForm.Visible = True Then
                    If CargarCombos(oForm) = False Then
                        Exit Function
                    End If

                    If InicializarValores(oForm) = False Then
                        Exit Function
                    End If
                End If
            End If

            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oConds, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCond, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_ItemPressed_After(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = SboApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "1" Then
                If pVal.ActionSuccess = True Then
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        'Después de añadir cargamos el registro creado
                        If PosicionarRegistro(oForm, _sDocEntry) = False Then
                            Exit Function
                        End If
                    End If
                End If
            End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            _sDocEntry = "0"
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

#End Region

#Region "Métodos auxiliares"

    Private Function PosicionarRegistro(ByRef oForm As SAPbouiCOM.Form, ByVal sDocEntry As String) As Boolean
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing

        PosicionarRegistro = False

        Try
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

            oConds = New SAPbouiCOM.Conditions
            oCond = oConds.Add
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCond.Alias = "DocEntry"
            oCond.CondVal = sDocEntry

            oForm.DataSources.DBDataSources.Item("@EXO_OICO").Query(oConds)
            oForm.DataSources.DBDataSources.Item("@EXO_ICO1").Query(oConds)

            CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).LoadFromDataSource()

            PosicionarRegistro = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oConds, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCond, Object))
        End Try
    End Function

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
                'Si las compañías del company list tienen los siguientes dos campos entonces las cargamos
                sSQL = "SELECT COL.Name " & _
                        "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.syscolumns COL WITH (NOLOCK) INNER JOIN " & _
                        "[" & oRs.Fields.Item(0).Value.ToString & "].dbo.sysobjects OBJ WITH (NOLOCK) ON OBJ.id = COL.id " & _
                        "WHERE OBJ.name = 'OADM' " & _
                        "AND COL.name = 'U_EXO_CONSOLIDACION'"

                oRsAux.DoQuery(sSQL)

                If oRsAux.RecordCount = 1 Then
                    'Combo Sucursales
                    sSQL = "SELECT t1.CompnyName Name " & _
                           "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.[OADM] t1 WITH (NOLOCK) " & _
                           "WHERE ISNULL(t1.U_EXO_CONSOLIDACION, 'N') = 'N' "

                    oRsAux.DoQuery(sSQL)

                    If oRsAux.RecordCount > 0 Then
                        Try
                            CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(1).Specific, SAPbouiCOM.ComboBox).ValidValues.Add(oRs.Fields.Item(0).Value.ToString, oRs.Fields.Item(1).Value.ToString)
                        Catch exCOM As System.Runtime.InteropServices.COMException
                        Catch ex As Exception
                        End Try
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

    Private Function InicializarValores(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        InicializarValores = False

        Try
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                oRs.DoQuery("SELECT t2.dbName " & _
                            "FROM OADM t1 WITH (NOLOCK) INNER JOIN " & _
                            "[SBO-COMMON].dbo.[SRGC] t2 WITH (NOLOCK) ON t1.CompnyName = t2.cmpName " & _
                            "WHERE ISNULL(t1.U_EXO_CONSOLIDACION, 'N') = 'Y'")

                If oRs.RecordCount > 0 Then
                    oForm.DataSources.DBDataSources.Item("@EXO_OICO").SetValue("U_EXO_DBNAME", 0, oRs.Fields.Item("dbName").Value.ToString)
                End If
            End If

            InicializarValores = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

    Private Function AddDatabasesInterCompany(ByRef oForm As SAPbouiCOM.Form, ByVal sDocEntry As String) As Boolean
        Dim oRsAux As SAPbobsCOM.Recordset = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim sSQLDB As String = ""

        AddDatabasesInterCompany = False

        Try
            oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsAux = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            If objGlobal.conexionSAP.SBOApp.MessageBox("Se va a actualizar la base de datos InterCompany con las empresas indicadas. ¿Desea continuar?", 1, "Sí", "No") = 1 Then
                If Company.InTransaction = True Then
                    Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End If
                Company.StartTransaction()

                oRsAux.DoQuery("SELECT dbName " & _
                                "FROM [INTERCOMPANY].dbo.[DATABASES] WITH (NOLOCK) " & _
                                "WHERE dbName = '" & oForm.DataSources.DBDataSources.Item("@EXO_OICO").GetValue("U_EXO_DBNAME", 0).Trim & "' AND dbTipo = 'C'")

                If oRsAux.RecordCount = 0 Then
                    oRsAux.DoQuery("INSERT INTO [INTERCOMPANY].dbo.[DATABASES] (dbName, dbTipo) " & _
                                    "VALUES ('" & oForm.DataSources.DBDataSources.Item("@EXO_OICO").GetValue("U_EXO_DBNAME", 0).Trim & "', 'C')")
                End If

                For i As Integer = 0 To CInt(oForm.DataSources.DBDataSources.Item("@EXO_ICO1").Size - 1)
                    oRsAux.DoQuery("SELECT dbName " & _
                                    "FROM [INTERCOMPANY].dbo.[DATABASES] WITH (NOLOCK) " & _
                                    "WHERE dbName = '" & oForm.DataSources.DBDataSources.Item("@EXO_ICO1").GetValue("U_EXO_DBNAME", i).Trim & "' AND dbTipo = 'S'")

                    If oRsAux.RecordCount = 0 Then
                        oRsAux.DoQuery("INSERT INTO [INTERCOMPANY].dbo.[DATABASES] (dbName, dbTipo) " & _
                                        "VALUES ('" & oForm.DataSources.DBDataSources.Item("@EXO_ICO1").GetValue("U_EXO_DBNAME", i).Trim & "', 'S')")
                    End If
                Next

                oRs = objGlobal.conexionSAP.compañia.GetCompanyList

                While Not oRs.EoF
                    'Si las compañías del company list tienen la siguiente tabla
                    sSQL = "SELECT OBJ.name " & _
                            "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.sysobjects OBJ WITH (NOLOCK) " & _
                            "WHERE OBJ.name = '@EXO_OICO' "

                    oRsAux.DoQuery(sSQL)

                    If oRsAux.RecordCount > 0 Then
                        If sSQLDB <> "" Then sSQLDB &= " UNION ALL "

                        sSQLDB &= "SELECT ISNULL(t1.U_EXO_DBNAME, '') DB " & _
                                  "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.[@EXO_ICO1] t1 WITH (NOLOCK) " & _
                                  "WHERE ISNULL(t1.U_EXO_DBNAME, '') <> '' "

                        sSQLDB &= " UNION ALL "

                        sSQLDB &= "SELECT ISNULL(t1.U_EXO_DBNAME, '') DB " & _
                                  "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.[@EXO_OICO] t1 WITH (NOLOCK) " & _
                                  "WHERE ISNULL(t1.U_EXO_DBNAME, '') <> '' "
                    End If

                    oRs.MoveNext()
                End While

                sSQLDB = "DELETE FROM [INTERCOMPANY].dbo.[DATABASES] WHERE dbName NOT IN (SELECT DB FROM (" & sSQLDB & ") TABTEMP)"

                oRsAux.DoQuery(sSQLDB)

                If Company.InTransaction = True Then
                    Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If
            End If

            AddDatabasesInterCompany = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Company.InTransaction = True Then
                Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux, Object))
        End Try
    End Function

    'Private Function ComprobarDatos(ByRef oForm As SAPbouiCOM.Form) As Boolean
    '    Dim oRs As SAPbobsCOM.Recordset = Nothing
    '    Dim oRsAux As SAPbobsCOM.Recordset = Nothing
    '    Dim sGrupoEmpresa As String = ""
    '    Dim sSQL As String = ""

    '    ComprobarDatos = False

    '    Try
    '        oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
    '        oRsAux = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

    '        For i As Integer = 0 To oForm.DataSources.DBDataSources.Item("@EXO_ICO1").Size - 1
    '            If oForm.DataSources.DBDataSources.Item("@EXO_ICO1").GetValue("U_EXO_DBNAME", i).Trim <> "" Then
    '                sSQL = "SELECT ISNULL(t1.U_EXO_GRUPOEMPRESA, '') U_EXO_GRUPOEMPRESA " & _
    '                       "FROM [" & oForm.DataSources.DBDataSources.Item("@EXO_ICO1").GetValue("U_EXO_DBNAME", i).Trim & "].dbo.[OADM] t1 WITH (NOLOCK) " & _
    '                       "WHERE ISNULL(t1.U_EXO_GRUPOEMPRESA, '') <> ''"

    '                oRsAux.DoQuery(sSQL)

    '                If oRsAux.RecordCount > 0 Then
    '                    sGrupoEmpresa = oRsAux.Fields.Item("U_EXO_GRUPOEMPRESA").Value.ToString
    '                Else
    '                    sGrupoEmpresa = ""
    '                End If

    '                For j As Integer = 0 To oForm.DataSources.DBDataSources.Item("@EXO_ICO1").Size - 1
    '                    If oForm.DataSources.DBDataSources.Item("@EXO_ICO1").GetValue("U_EXO_DBNAME", i).Trim <> oForm.DataSources.DBDataSources.Item("@EXO_ICO1").GetValue("U_EXO_DBNAME", j).Trim Then
    '                        sSQL = "SELECT ISNULL(t1.U_EXO_GRUPOEMPRESA, '') U_EXO_GRUPOEMPRESA " & _
    '                               "FROM [" & oForm.DataSources.DBDataSources.Item("@EXO_ICO1").GetValue("U_EXO_DBNAME", j).Trim & "].dbo.[OADM] t1 WITH (NOLOCK) " & _
    '                               "WHERE ISNULL(t1.U_EXO_GRUPOEMPRESA, '') <> ''"

    '                        oRsAux.DoQuery(sSQL)

    '                        If oRsAux.RecordCount > 0 Then
    '                            If sGrupoEmpresa = oRsAux.Fields.Item("U_EXO_GRUPOEMPRESA").Value.ToString Then

    '                            End If
    '                        End If

    '                        Exit For
    '                    End If
    '                Next
    '            End If
    '        Next

    '        If iContGrupo > 1 Then
    '            SboApp.StatusBar.SetText("Existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    '            Exit Function
    '        End If

    '        ComprobarDatos = True

    '    Catch exCOM As System.Runtime.InteropServices.COMException
    '        Throw exCOM
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
    '        EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux, Object))
    '    End Try
    'End Function

#End Region

End Class
