Public Class EXO_OCON
    Inherits EXO_Generales.EXO_DLLBase

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

        Try
            If objGlobal.conexionSAP.esAdministrador Then
                'Campos de Usuario para configuración de InterCompany
                oXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "UDFs_OJDT.xml")
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("Validando: UDFs OJDT", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
        Dim fXML As String = objGlobal.Functions.leerEmbebido(Me.GetType(), "Filtros_EXO_OCON.xml")
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
                menuXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "EXO_MENUCONSO.xml")
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
                autorizacionXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "EXO_AUCONSO.xml")
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
        'Dim oPermisoUser As SAPbobsCOM.BoPermission = Nothing

        Try
            If infoEvento.BeforeAction = True Then
                oForm = SboApp.Forms.ActiveForm

                Select Case oForm.TypeEx
                    Case "169"

                        Select Case infoEvento.MenuUID
                            Case "EXO-MnConso"
                                'oPermisoUser = objGlobal.conexionSAP.refCompañia.autorizacionUsuario("EXO_AUCONSO")

                                'If oPermisoUser = SAPbobsCOM.BoPermission.boper_Full OrElse oPermisoUser = SAPbobsCOM.BoPermission.boper_ReadOnly Then
                                If EventHandler_Form_Load() = False Then
                                    GC.Collect()
                                    Return False
                                End If
                                'Else
                                'Me.SboApp.StatusBar.SetText("El usuario no tiene permisos para acceder a este formulario.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                'End If

                        End Select

                End Select

            Else
                oForm = SboApp.Forms.ActiveForm

                Select Case oForm.TypeEx
                    Case "EXO_OCON"

                        Select Case infoEvento.MenuUID

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

    Public Overrides Function SBOApp_ItemEvent(ByRef infoEvento As EXO_Generales.EXO_infoItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_OCON"

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
                        Case "EXO_OCON"

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
                        Case "EXO_OCON"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

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
                        Case "EXO_OCON"

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

    Private Function EventHandler_Form_Load() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_Generales.EXO_XML(objGlobal.conexionSAP.refCompañia, objGlobal.conexionSAP.refSBOApp)

        EventHandler_Form_Load = False

        Try
            'Buscar XML de update
            SboApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Path = objGlobal.conexionSAP.pathPantallas
            If Path = "" Then
                Return False
            End If

            oFP = CType(SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = EXO_Xml.LoadFormXml(Path & "\EXO_OCON.srf", False).ToString

            Try
                'Lo metemos en un try catch porque si intenta acceder un usuario sin autorización salta un error interno
                oForm = SboApp.Forms.AddEx(oFP)

                If CargarCombos(oForm) = False Then
                    Exit Function
                End If
            Catch exCOM As System.Runtime.InteropServices.COMException
            Catch ex As Exception
            End Try

            EventHandler_Form_Load = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_ItemPressed_After(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = SboApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "3" Then
                If pVal.ActionSuccess = True Then
                    If ComprobarDatos(oForm) = False Then
                        Exit Function
                    End If

                    If Consolidar(oForm) = False Then
                        Exit Function
                    End If
                End If
            End If

            EventHandler_ItemPressed_After = True

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
                'Si las compañías del company list tienen los siguientes tres campos entonces cargamos el grupo de empresas
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
                    sSQL = "SELECT DocEntry " & _
                           "FROM [@EXO_ICO1] t1 WITH (NOLOCK) " & _
                           "WHERE t1.U_EXO_DBNAME = '" & oRs.Fields.Item(0).Value.ToString & "' " & _
                           "AND t1.U_EXO_PRCNT > 0 "

                    oRsAux.DoQuery(sSQL)

                    If oRsAux.RecordCount > 0 Then
                        'Combo Grupo de empresas
                        sSQL = "SELECT ISNULL(t1.U_EXO_GRUPOEMPRESA, '') U_EXO_GRUPOEMPRESA " & _
                               "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.[OADM] t1 WITH (NOLOCK) " & _
                               "WHERE (ISNULL(t1.U_EXO_MATRIZ, 'N') = 'Y' " & _
                               "OR (ISNULL(t1.U_EXO_CONSOLIDACION, 'N') = 'N' " & _
                               "AND ISNULL(t1.U_EXO_MATRIZ, 'N') = 'N')) " & _
                               "AND ISNULL(t1.U_EXO_GRUPOEMPRESA, '') <> ''"

                        oRsAux.DoQuery(sSQL)

                        If oRsAux.RecordCount > 0 Then
                            Try
                                CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.ComboBox).ValidValues.Add(oRsAux.Fields.Item("U_EXO_GRUPOEMPRESA").Value.ToString, oRsAux.Fields.Item("U_EXO_GRUPOEMPRESA").Value.ToString)
                            Catch exCOM As System.Runtime.InteropServices.COMException
                            Catch ex As Exception
                            End Try
                        End If
                    End If
                End If

                oRs.MoveNext()
            End While

            If CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.ComboBox).ValidValues.Count > 0 Then
                CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.ComboBox).ValidValues.Add("ALL", "ALL")
                CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.ComboBox).Select("ALL", SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If

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

    Private Function ComprobarDatos(ByRef oForm As SAPbouiCOM.Form) As Boolean
        ComprobarDatos = False

        Try
            If IsDate(oForm.DataSources.UserDataSources.Item("DocDateD").Value) = False AndAlso IsDate(oForm.DataSources.UserDataSources.Item("DocDateH").Value) = False Then
                SboApp.StatusBar.SetText("Debe indicar al menos una fecha.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Exit Function
            ElseIf IsDate(oForm.DataSources.UserDataSources.Item("DocDateD").Value) AndAlso IsDate(oForm.DataSources.UserDataSources.Item("DocDateH").Value) Then
                If CDate(oForm.DataSources.UserDataSources.Item("DocDateD").Value) > CDate(oForm.DataSources.UserDataSources.Item("DocDateH").Value) Then
                    SboApp.StatusBar.SetText("La fecha de contabilización desde debe ser menor o igual a la fecha de contabilización hasta.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    Exit Function
                End If
            ElseIf oForm.DataSources.UserDataSources.Item("Asiento").Value.Trim = "" Then
                SboApp.StatusBar.SetText("Debe indicar los asientos a importar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Exit Function
            End If

            ComprobarDatos = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function Consolidar(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux As SAPbobsCOM.Recordset = Nothing
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim cPrcnt As Double = 0
        Dim sTipo As String = ""
        Dim sdbName As String = ""
        Dim sFile As String = ""
        Dim log As EXO_Log.EXO_Log = Nothing

        Consolidar = False

        Try
            oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsAux = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            If oForm.DataSources.UserDataSources.Item("Asiento").Value.Trim = "ALL" Then
                sSQL = "SELECT t1.U_EXO_DBNAME, t1.U_EXO_PRCNT, t1.U_EXO_TIPOCONSO " & _
                       "FROM [@EXO_ICO1] t1 WITH (NOLOCK) " 
            Else
                oRs = objGlobal.conexionSAP.compañia.GetCompanyList

                While Not oRs.EoF
                    oRsAux.DoQuery("SELECT t1.U_EXO_DBNAME " & _
                                   "FROM [@EXO_ICO1] t1 WITH (NOLOCK) " & _
                                   "WHERE t1.U_EXO_DBNAME = '" & oRs.Fields.Item(0).Value.ToString & "'")

                    If oRsAux.RecordCount > 0 Then
                        oRsAux.DoQuery("SELECT t1.CompnyName " & _
                                       "FROM [" & oRs.Fields.Item(0).Value.ToString & "].dbo.[OADM] t1 WITH (NOLOCK) " & _
                                       "WHERE ISNULL(t1.U_EXO_GRUPOEMPRESA, '') = '" & oForm.DataSources.UserDataSources.Item("Asiento").Value.Trim & "'")

                        If oRsAux.RecordCount > 0 Then
                            sSQL = "SELECT t1.U_EXO_DBNAME, t1.U_EXO_PRCNT, t1.U_EXO_TIPOCONSO " & _
                                   "FROM [@EXO_ICO1] t1 WITH (NOLOCK) " & _
                                   "WHERE t1.U_EXO_DBNAME = '" & oRs.Fields.Item(0).Value.ToString & "'"

                            Exit While
                        End If
                    End If

                    oRs.MoveNext()
                End While
            End If

            If sSQL <> "" Then
                oRsAux.DoQuery(sSQL)

                oXml.LoadXml(oRsAux.GetAsXML())
                oNodes = oXml.SelectNodes("//row")

                If oRsAux.RecordCount > 0 Then
                    sFile = System.IO.Path.GetTempPath() & Guid.NewGuid().ToString() & ".txt"

                    If System.IO.File.Exists(sFile) = False Then
                        log = New EXO_Log.EXO_Log(sFile, 1)
                    End If

                    For i As Integer = 0 To oNodes.Count - 1
                        oNode = oNodes.Item(i)

                        sdbName = oNode.SelectSingleNode("U_EXO_DBNAME").InnerText
                        cPrcnt = CDbl(oNode.SelectSingleNode("U_EXO_PRCNT").InnerText.Replace(".", ","))
                        sTipo = oNode.SelectSingleNode("U_EXO_TIPOCONSO").InnerText

                        If OJDT(oForm, sdbName, cPrcnt, sTipo, log) = False Then
                            Exit Function
                        End If
                    Next

                    Process.Start(sFile)

                    Me.SboApp.StatusBar.SetText("Fin de la consolidación contable. Abriendo fichero de log. Espere por favor ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Else
                    Me.SboApp.StatusBar.SetText("El grupo de empresa seleccionado no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            Else
                Me.SboApp.StatusBar.SetText("El grupo de empresa seleccionado no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            Consolidar = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux, Object))
        End Try
    End Function

    Private Function OJDT(ByRef oForm As SAPbouiCOM.Form, ByVal sdbName As String, ByVal cPrcnt As Double, ByVal sTipo As String, ByRef log As EXO_Log.EXO_Log) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim oXmlRs As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim oXmlRsAux As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodesAux As System.Xml.XmlNodeList = Nothing
        Dim oNodeAux As System.Xml.XmlNode = Nothing
        Dim oCompanyO As SAPbobsCOM.Company = Nothing
        Dim oOJDT As SAPbobsCOM.JournalEntries = Nothing
        Dim sXML As String = ""
        Dim oXml As Xml.XmlDocument = Nothing
        Dim oXmlNode As Xml.XmlNode = Nothing
        Dim oXmlNodes As Xml.XmlNodeList = Nothing
        Dim sGrupoEmpresa As String = ""
        Dim sCeCos As String = ""
        Dim sImportes As String = ""
        Dim dicOMDR As Dictionary(Of String, String) = New Dictionary(Of String, String)
        Dim sCostingCode As String = ""
        Dim sCostingCode2 As String = ""
        Dim sCostingCode3 As String = ""
        Dim cSumCredit As Double = 0
        Dim cSumDebit As Double = 0

        OJDT = False

        Try
            oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsAux = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT DISTINCT t1.TransId, t1.Number, ISNULL(t1.U_EXO_COMPANY1, '') U_EXO_COMPANY1, ISNULL(t1.U_EXO_COMPANY2, '') U_EXO_COMPANY2 " & _
                   "FROM [" & sdbName & "].dbo.[OJDT] t1 WITH (NOLOCK) INNER JOIN " & _
                   "[" & sdbName & "].dbo.[JDT1] t2 WITH (NOLOCK) ON t1.TransId = t2.TransId " & _
                   "WHERE t1.TransType <> '-3' " & _
                   "AND t1.TransType <> '-2' " & _
                   "AND ISNULL(t1.U_EXO_COMPANY1, '') <> '" & objGlobal.conexionSAP.compañia.CompanyDB & "' " & _
                   "AND ISNULL(t1.U_EXO_COMPANY2, '') <> '" & objGlobal.conexionSAP.compañia.CompanyDB & "' " & _
                   "AND t1.TransId NOT IN (SELECT DISTINCT t3.TransId " & _
                                           "FROM [" & sdbName & "].dbo.[JDT1] t3 WITH (NOLOCK) INNER JOIN " & _
                                           "[" & objGlobal.conexionSAP.compañia.CompanyDB & "].dbo.[@EXO_OCTE] t4 WITH (NOLOCK) ON t3.Account = t4.U_EXO_ACCTCODE " & _
                                           "WHERE t3.TransId = t1.TransId) "
            '"AND ISNULL(t1.StornoToTr, 0) = 0 " & _
            '"AND t1.TransId NOT IN (SELECT t5.StornoToTr " & _
            '                        "FROM [" & sdbName & "].dbo.[OJDT] t5 WITH (NOLOCK) " & _
            '                        "WHERE t5.StornoToTr = t1.TransId) "

            If IsDate(oForm.DataSources.UserDataSources.Item("DocDateD").Value) = True Then
                sSQL &= "AND CONVERT(DATE, t1.RefDate, 112) >= CONVERT(DATE, '" & oForm.DataSources.UserDataSources.Item("DocDateD").ValueEx & "', 112) "
            End If

            If IsDate(oForm.DataSources.UserDataSources.Item("DocDateH").Value) = True Then
                sSQL &= "AND CONVERT(DATE, t1.RefDate, 112) <= CONVERT(DATE, '" & oForm.DataSources.UserDataSources.Item("DocDateH").ValueEx & "', 112) "
            End If

            oRs.DoQuery(sSQL)

            oXmlRs.LoadXml(oRs.GetAsXML())
            oNodes = oXmlRs.SelectNodes("//row")

            If oRs.RecordCount > 0 Then
                oRsAux.DoQuery("SELECT ISNULL(U_EXO_GRUPOEMPRESA, '') U_EXO_GRUPOEMPRESA " & _
                               "FROM [" & sdbName & "].dbo.[OADM] ")

                If oRsAux.RecordCount > 0 Then
                    sGrupoEmpresa = oRsAux.Fields.Item("U_EXO_GRUPOEMPRESA").Value.ToString
                End If

                Me.SboApp.StatusBar.SetText("Iniciando consolidación compañía " & sdbName, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                EXO_GLOBALES.Connect_Company(objGlobal, oCompanyO, sdbName)

                oCompanyO.XMLAsString = True
                oCompanyO.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                objGlobal.conexionSAP.compañia.XMLAsString = True
                objGlobal.conexionSAP.compañia.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode

                For i As Integer = 0 To oNodes.Count - 1
                    Try
                        Me.SboApp.StatusBar.SetText(sdbName & " - Asiento " & (i + 1).ToString & " de " & oNodes.Count, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        oNode = oNodes.Item(i)

                        cSumCredit = 0
                        cSumDebit = 0

                        If Me.Company.InTransaction = True Then
                            Me.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        Me.Company.StartTransaction()

                        oOJDT = CType(oCompanyO.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries), SAPbobsCOM.JournalEntries)

                        If oOJDT.GetByKey(CInt(oNode.SelectSingleNode("TransId").InnerText)) = True Then
                            sXML = oOJDT.GetAsXML
                        Else
                            sXML = ""
                        End If

                        If sXML <> "" Then
                            oXml = New Xml.XmlDocument
                            oXml.LoadXml(sXML)

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/JournalEntries/row/Memo")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/JournalEntries/row/Reference")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/JournalEntries/row/Reference2")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/JournalEntries/row/Reference3")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/JournalEntries/row/TransactionCode")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/JournalEntries/row/ProjectCode")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/JournalEntries/row/JdtNum")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/JournalEntries/row/Indicator")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/JournalEntries/row/Series")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            Try
                                oXmlNode = oXml.SelectSingleNode("/BOM/BO/PrimaryFormItems")
                                oXmlNode.ParentNode.RemoveChild(oXmlNode)
                            Catch ex As Exception

                            End Try

                            oXmlNodes = oXml.SelectNodes("/BOM/BO/JournalEntries_Lines/row")

                            For j As Integer = oXmlNodes.Count - 1 To 0 Step -1
                                'Añadir etiquetas Debit, Credit, CreditSys y DebitSys si no existen. Error DI API encontrado en el tipo de IVA S0
                                oXmlNode = oXmlNodes.Item(j).SelectSingleNode("Debit")

                                If oXmlNode Is Nothing Then
                                    oXmlNode = oXml.CreateNode(System.Xml.XmlNodeType.Element, "Debit", "")
                                    oXmlNode.InnerText = "0"
                                    oXmlNode = oXmlNodes.Item(j).AppendChild(oXmlNode)
                                End If

                                oXmlNode = oXmlNodes.Item(j).SelectSingleNode("Credit")

                                If oXmlNode Is Nothing Then
                                    oXmlNode = oXml.CreateNode(System.Xml.XmlNodeType.Element, "Credit", "")
                                    oXmlNode.InnerText = "0"
                                    oXmlNode = oXmlNodes.Item(j).AppendChild(oXmlNode)
                                End If

                                oXmlNode = oXmlNodes.Item(j).SelectSingleNode("CreditSys")

                                If oXmlNode Is Nothing Then
                                    oXmlNode = oXml.CreateNode(System.Xml.XmlNodeType.Element, "CreditSys", "")
                                    oXmlNode.InnerText = "0"
                                    oXmlNode = oXmlNodes.Item(j).AppendChild(oXmlNode)
                                End If

                                oXmlNode = oXmlNodes.Item(j).SelectSingleNode("DebitSys")

                                If oXmlNode Is Nothing Then
                                    oXmlNode = oXml.CreateNode(System.Xml.XmlNodeType.Element, "DebitSys", "")
                                    oXmlNode.InnerText = "0"
                                    oXmlNode = oXmlNodes.Item(j).AppendChild(oXmlNode)
                                End If

                                If sTipo = "S" Then
                                    'Si el tipo de consolidación es 'Sistema'
                                    oXmlNodes.Item(j).SelectSingleNode("Credit").InnerText = Math.Round((CDbl(oXmlNodes.Item(j).SelectSingleNode("CreditSys").InnerText.Replace(".", ",")) * cPrcnt) / 100, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero).ToString.Replace(",", ".")

                                    oXmlNodes.Item(j).SelectSingleNode("Debit").InnerText = Math.Round((CDbl(oXmlNodes.Item(j).SelectSingleNode("DebitSys").InnerText.Replace(".", ",")) * cPrcnt) / 100, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero).ToString.Replace(",", ".")

                                    cSumCredit += CDbl(oXmlNodes.Item(j).SelectSingleNode("Credit").InnerText.Replace(".", ","))
                                    cSumDebit += CDbl(oXmlNodes.Item(j).SelectSingleNode("Debit").InnerText.Replace(".", ","))
                                Else
                                    'Si el tipo de consolidación es 'Local'
                                    oXmlNodes.Item(j).SelectSingleNode("Credit").InnerText = Math.Round((CDbl(oXmlNodes.Item(j).SelectSingleNode("Credit").InnerText.Replace(".", ",")) * cPrcnt) / 100, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero).ToString.Replace(",", ".")

                                    oXmlNodes.Item(j).SelectSingleNode("Debit").InnerText = Math.Round((CDbl(oXmlNodes.Item(j).SelectSingleNode("Debit").InnerText.Replace(".", ",")) * cPrcnt) / 100, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero).ToString.Replace(",", ".")

                                    cSumCredit += CDbl(oXmlNodes.Item(j).SelectSingleNode("Credit").InnerText.Replace(".", ","))
                                    cSumDebit += CDbl(oXmlNodes.Item(j).SelectSingleNode("Debit").InnerText.Replace(".", ","))
                                End If

                                'Ajustamos el importe al último asiento si hay diferencias
                                If j = 0 Then
                                    If Math.Round(cSumCredit, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero) > Math.Round(cSumDebit, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero) Then
                                        If CDbl(oXmlNodes.Item(j).SelectSingleNode("Debit").InnerText.Replace(".", ",")) <> 0 Then
                                            oXmlNodes.Item(j).SelectSingleNode("Debit").InnerText = Math.Round(CDbl(oXmlNodes.Item(j).SelectSingleNode("Debit").InnerText.Replace(".", ",")) + Math.Round(Math.Round(cSumCredit, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero) - Math.Round(cSumDebit, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero), objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero), objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero).ToString.Replace(",", ".")
                                        Else
                                            oXmlNodes.Item(j).SelectSingleNode("Credit").InnerText = Math.Round(CDbl(oXmlNodes.Item(j).SelectSingleNode("Credit").InnerText.Replace(".", ",")) - Math.Round(Math.Round(cSumCredit, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero) - Math.Round(cSumDebit, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero), objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero), objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero).ToString.Replace(",", ".")
                                        End If
                                    ElseIf Math.Round(cSumCredit, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero) < Math.Round(cSumDebit, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero) Then
                                        If CDbl(oXmlNodes.Item(j).SelectSingleNode("Credit").InnerText.Replace(".", ",")) <> 0 Then
                                            oXmlNodes.Item(j).SelectSingleNode("Credit").InnerText = Math.Round(CDbl(oXmlNodes.Item(j).SelectSingleNode("Credit").InnerText.Replace(".", ",")) + Math.Round(Math.Round(cSumDebit, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero) - Math.Round(cSumCredit, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero), objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero), objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero).ToString.Replace(",", ".")
                                        Else
                                            oXmlNodes.Item(j).SelectSingleNode("Debit").InnerText = Math.Round(CDbl(oXmlNodes.Item(j).SelectSingleNode("Debit").InnerText.Replace(".", ",")) - Math.Round(Math.Round(cSumDebit, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero) - Math.Round(cSumCredit, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero), objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero), objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero).ToString.Replace(",", ".")
                                        End If
                                    End If
                                End If

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("Line_ID")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("FCDebit")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("FCCredit")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("ShortName")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("ContraAccount")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("LineMemo")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("Reference1")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("Reference2")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("ProjectCode")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("BaseSum")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("TaxGroup")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("DebitSys")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("CreditSys")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("VatLine")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("SystemBaseAmount")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("VatAmount")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("SystemVatAmount")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("GrossValue")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                Try
                                    oXmlNode = oXmlNodes.Item(j).SelectSingleNode("AdditionalReference")
                                    oXmlNode.ParentNode.RemoveChild(oXmlNode)
                                Catch ex As Exception

                                End Try

                                'Normas de reparto manuales
                                sCostingCode = ""
                                Try
                                    sCostingCode = oXmlNodes.Item(j).SelectSingleNode("CostingCode").InnerText
                                Catch ex As Exception

                                End Try

                                If sCostingCode <> "" Then
                                    sCeCos = ""
                                    sImportes = ""

                                    RepartoCeCos(oCompanyO, sCeCos, sImportes, cPrcnt, sCostingCode)
                                    If sCeCos <> "" Then
                                        If dicOMDR.ContainsKey(sCostingCode) = False Then
                                            oXmlNodes.Item(j).SelectSingleNode("CostingCode").InnerText = Create_RepartoManual("1", sCeCos, sImportes)

                                            dicOMDR.Add(sCostingCode, oXmlNodes.Item(j).SelectSingleNode("CostingCode").InnerText)
                                        Else
                                            oXmlNodes.Item(j).SelectSingleNode("CostingCode").InnerText = dicOMDR.Item(sCostingCode)
                                        End If
                                    End If
                                End If

                                sCostingCode2 = ""
                                Try
                                    sCostingCode2 = oXmlNodes.Item(j).SelectSingleNode("CostingCode2").InnerText
                                Catch ex As Exception

                                End Try

                                If sCostingCode2 <> "" Then
                                    sCeCos = ""
                                    sImportes = ""

                                    RepartoCeCos(oCompanyO, sCeCos, sImportes, cPrcnt, sCostingCode2)
                                    If sCeCos <> "" Then
                                        If dicOMDR.ContainsKey(sCostingCode2) = False Then
                                            oXmlNodes.Item(j).SelectSingleNode("CostingCode2").InnerText = Create_RepartoManual("2", sCeCos, sImportes)

                                            dicOMDR.Add(sCostingCode2, oXmlNodes.Item(j).SelectSingleNode("CostingCode2").InnerText)
                                        Else
                                            oXmlNodes.Item(j).SelectSingleNode("CostingCode2").InnerText = dicOMDR.Item(sCostingCode2)
                                        End If
                                    End If
                                End If

                                sCostingCode3 = ""
                                Try
                                    sCostingCode3 = oXmlNodes.Item(j).SelectSingleNode("CostingCode3").InnerText
                                Catch ex As Exception

                                End Try

                                If sCostingCode3 <> "" Then
                                    sCeCos = ""
                                    sImportes = ""

                                    RepartoCeCos(oCompanyO, sCeCos, sImportes, cPrcnt, sCostingCode3)
                                    If sCeCos <> "" Then
                                        If dicOMDR.ContainsKey(sCostingCode3) = False Then
                                            oXmlNodes.Item(j).SelectSingleNode("CostingCode3").InnerText = Create_RepartoManual("3", sCeCos, sImportes)

                                            dicOMDR.Add(sCostingCode3, oXmlNodes.Item(j).SelectSingleNode("CostingCode3").InnerText)
                                        Else
                                            oXmlNodes.Item(j).SelectSingleNode("CostingCode3").InnerText = dicOMDR.Item(sCostingCode3)
                                        End If
                                    End If
                                End If

                            Next

                            sXML = oXml.OuterXml

                            oOJDT = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries), SAPbobsCOM.JournalEntries)

                            oOJDT = Me.Company.GetBusinessObjectFromXML(sXML, 0)

                            oOJDT.AutoVAT = SAPbobsCOM.BoYesNoEnum.tNO
                            oOJDT.UserFields.Fields.Item("U_EXO_GRUPOEMPRESA").Value = sGrupoEmpresa

                            If oOJDT.Add() <> 0 Then
                                Throw New Exception(Me.Company.GetLastErrorCode & " / " & Me.Company.GetLastErrorDescription)
                            End If

                            If oNode.SelectSingleNode("U_EXO_COMPANY1").InnerText = "" Then
                                oRsAux.DoQuery("UPDATE [" & sdbName & "].dbo.[OJDT] SET U_EXO_COMPANY1 = '" & objGlobal.conexionSAP.compañia.CompanyDB & "' " & _
                                               "WHERE TransId = " & oNode.SelectSingleNode("TransId").InnerText)
                            ElseIf oNode.SelectSingleNode("U_EXO_COMPANY2").InnerText = "" Then
                                oRsAux.DoQuery("UPDATE [" & sdbName & "].dbo.[OJDT] SET U_EXO_COMPANY2 = '" & objGlobal.conexionSAP.compañia.CompanyDB & "' " & _
                                               "WHERE TransId = " & oNode.SelectSingleNode("TransId").InnerText)
                            End If
                        End If

                        If Me.Company.InTransaction = True Then
                            Me.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                        End If

                    Catch exCOM As System.Runtime.InteropServices.COMException
                        log.escribeMensaje("DB " & sdbName & " ASIENTO " & oNode.SelectSingleNode("Number").InnerText & " ----- " & exCOM.Message, EXO_Log.EXO_Log.Tipo.error)

                        If Me.Company.InTransaction = True Then
                            Me.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    Catch ex As Exception
                        log.escribeMensaje("DB " & sdbName & " ASIENTO " & oNode.SelectSingleNode("Number").InnerText & " ----- " & ex.Message, EXO_Log.EXO_Log.Tipo.error)

                        If Me.Company.InTransaction = True Then
                            Me.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                    End Try
                Next
            End If

            OJDT = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            If Me.Company.InTransaction = True Then
                Me.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            Throw exCOM
        Catch ex As Exception
            If Me.Company.InTransaction = True Then
                Me.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOJDT, Object))

            EXO_GLOBALES.Disconnect_Company(oCompanyO)

            Me.SboApp.StatusBar.SetText("Fin consolidación compañía " & sdbName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End Try
    End Function

    Private Sub RepartoCeCos(ByRef oCompany As SAPbobsCOM.Company, ByRef sCeCos As String, ByRef sImportes As String, ByVal cPrcnt As Double, ByVal sCostingCode As String)
        Dim oRsAux As SAPbobsCOM.Recordset = Nothing
        Dim oXmlAux As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodesAux As System.Xml.XmlNodeList = Nothing
        Dim oNodeAux As System.Xml.XmlNode = Nothing
        Dim cDiferencias As Double = 0
        Dim cOcrTotal As Double = 0
        Dim cPrcAmount As Double = 0
        Dim cSumOcrTotal As Double = 0

        Try
            oRsAux = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            If sCostingCode <> "" Then
                'Comprobamos que sea manual en el origen
                oRsAux.DoQuery("SELECT t1.PrcCode, t1.PrcAmount, t1.OcrTotal " & _
                               "FROM MDR1 t1 WITH (NOLOCK) " & _
                               "WHERE t1.OcrCode = '" & sCostingCode & "'")

                oXmlAux.LoadXml(oRsAux.GetAsXML())
                oNodesAux = oXmlAux.SelectNodes("//row")

                If oRsAux.RecordCount > 0 Then
                    cOcrTotal = Math.Round((CDbl(oNodesAux.Item(0).SelectSingleNode("OcrTotal").InnerText.Replace(".", ",")) * cPrcnt) / 100, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero)

                    For h As Integer = 0 To oNodesAux.Count - 1
                        oNodeAux = oNodesAux.Item(h)

                        cPrcAmount = Math.Round((CDbl(oNodesAux.Item(h).SelectSingleNode("PrcAmount").InnerText.Replace(".", ",")) * cPrcnt) / 100, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero)

                        cSumOcrTotal += cPrcAmount

                        If h = oNodesAux.Count - 1 Then
                            If sImportes = "" Then
                                sImportes = (cPrcAmount + (cOcrTotal - cSumOcrTotal)).ToString.Replace(".", ",")
                            Else
                                sImportes &= ";" & (cPrcAmount + (cOcrTotal - cSumOcrTotal)).ToString.Replace(".", ",")
                            End If
                        Else
                            If sImportes = "" Then
                                sImportes = cPrcAmount.ToString.Replace(".", ",")
                            Else
                                sImportes &= ";" & cPrcAmount.ToString.Replace(".", ",")
                            End If
                        End If

                        If sCeCos = "" Then
                            sCeCos = oNodesAux.Item(h).SelectSingleNode("PrcCode").InnerText
                        Else
                            sCeCos &= ";" & oNodesAux.Item(h).SelectSingleNode("PrcCode").InnerText
                        End If
                    Next
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux, Object))
        End Try
    End Sub

    Private Function Create_RepartoManual(ByVal sDimCode As String, ByVal sArrayCeCos As String, ByVal sArrayImportes As String) As String
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim iNumerador As Integer = 0
        Dim sSQL As String = ""
        Dim sCode As String = ""
        Dim TotalReparto As Double = 0
        Dim oMCecos() As String
        Dim oMImportes() As String
        Dim sImporte As Double = 0

        Create_RepartoManual = ""

        Try
            'Recupero el numerador de SAP
            oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT TOP 1 AutoKey FROM ONNM WITH (NOLOCK) WHERE ObjectCode = '252'"

            oRs.DoQuery(sSQL)

            If oRs.RecordCount = 0 Then
                Throw New Exception("No se pudo recuperar el numerador de repartos manuales")
            End If

            'Control de número
            If Integer.TryParse(oRs.Fields.Item(0).Value.ToString, iNumerador) = False Then
                iNumerador = 0
            End If
            If iNumerador = 0 Then
                Throw New Exception("No se pudo recuperar el numerador de repartos manuales")
            End If

            sSQL = "UPDATE ONNM SET AUTOKEY = " & iNumerador.ToString.Trim & " + 1 WHERE OBJECTCODE = '252'"

            oRs.DoQuery(sSQL)

            'Recupero el total de los CeCos
            oMImportes = Split(sArrayImportes, ";")

            For iIdx As Integer = 0 To oMImportes.Length - 1
                If Double.TryParse(oMImportes(iIdx).Trim, sImporte) = False Then
                    sImporte = 0
                End If
                If sImporte = 0 Then
                    Throw New Exception("No se admiten importes a 0 en los repartos manuales.")
                End If
                TotalReparto = TotalReparto + sImporte
            Next

            sCode = "M" & iNumerador.ToString("0000000")

            'Creo la cabecera
            sSQL = "INSERT INTO OMDR (OcrCode, OcrName, OcrTotal, Direct, Locked, DataSource, UserSign, DimCode, AbsEntry, Active)"
            sSQL = sSQL & " VALUES ("
            sSQL = sSQL & "'" & sCode & "'"
            sSQL = sSQL & ",'Norma de reparto manual'"
            sSQL = sSQL & "," & Replace(Replace(TotalReparto.ToString, ".", ""), ",", ".")
            sSQL = sSQL & ",'N'"
            sSQL = sSQL & ",'N'"
            sSQL = sSQL & ",'I'"
            sSQL = sSQL & "," & Me.Company.UserSignature
            sSQL = sSQL & "," & sDimCode
            sSQL = sSQL & "," & iNumerador
            sSQL = sSQL & ",'Y'"
            sSQL = sSQL & " )"

            oRs.DoQuery(sSQL)

            'Monto las líneas
            'Leeo de los CeCos
            oMCecos = Split(sArrayCeCos, ";")

            For iIdx As Integer = 0 To oMCecos.Length - 1
                If oMCecos(iIdx).Trim = "" Then
                    Throw New Exception("Error al crear la línea de repartos manuales. El Centro de Coste no puede ser blanco")
                End If

                sSQL = "INSERT INTO MDR1 (OcrCode, PrcCode, PrcAmount, OcrTotal, Direct, ValidFrom)"
                sSQL = sSQL & " VALUES ("
                sSQL = sSQL & "'" & sCode & "'"
                sSQL = sSQL & ",'" & oMCecos(iIdx).Trim & "'"

                If Double.TryParse(oMImportes(iIdx), sImporte) = False Then
                    sImporte = 0
                End If
                If sImporte = 0 Then
                    Throw New Exception("No se admiten importes a 0 en los repartos manuales.")
                End If

                sSQL = sSQL & "," & Replace(Replace(sImporte.ToString, ".", ""), ",", ".")
                sSQL = sSQL & "," & Replace(Replace(TotalReparto.ToString, ".", ""), ",", ".")
                sSQL = sSQL & ",'N'"
                sSQL = sSQL & ",'1900-01-01 00:00:00.000'"
                sSQL = sSQL & " )"

                oRs.DoQuery(sSQL)
            Next

            Create_RepartoManual = sCode

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

#End Region

End Class
