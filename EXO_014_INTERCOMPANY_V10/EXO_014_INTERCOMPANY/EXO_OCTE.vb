Imports SAPbouiCOM
Public Class EXO_OCTE
    Inherits EXO_UIAPI.EXO_DLLBase

#Region "Constructor"

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaDatos()
            cargaAutorizaciones()
        End If
    End Sub

#End Region

#Region "Inicialización"

    Private Sub cargaDatos()
        Dim oXML As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        If objGlobal.refDi.comunes.esAdministrador Then
            Try
                oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                oRs.DoQuery("SELECT CompnyName FROM OADM WITH (NOLOCK) WHERE ISNULL(U_EXO_CONSOLIDACION, 'N') = 'Y'")

                'Sólo generamos el UDO en las empresas de Consolidación
                If oRs.RecordCount > 0 Then
                    EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

                    'UDO Configuración InterCompany
                    oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_OCTE.xml")
                    objGlobal.refDi.comunes.LoadBDFromXML(oXML)
                    objGlobal.SBOApp.StatusBar.SetText("Validando: UDO EXO_OCTE", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If

            Catch exCOM As System.Runtime.InteropServices.COMException
                objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Catch ex As Exception
                objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Finally
                EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            End Try
        End If
    End Sub

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "Filtros_EXO_OCTE.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Dim menuXML As String = ""
        Dim res As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        Try
            oRs = CType(objglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs.DoQuery("SELECT CompnyName FROM OADM WITH (NOLOCK) WHERE ISNULL(U_EXO_CONSOLIDACION, 'N') = 'Y'")

            'Sólo cargamos el menú en las empresas de Consolidación
            If oRs.RecordCount > 0 Then
                menuXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MENUCTAEX.xml")
                objGlobal.SboApp.LoadBatchActions(menuXML)
                res = objglobal.SboApp.GetLastBatchResults
            End If

            Return Nothing

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return Nothing
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
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
            oRs = CType(objglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs.DoQuery("SELECT CompnyName FROM OADM WITH (NOLOCK) WHERE ISNULL(U_EXO_CONSOLIDACION, 'N') = 'Y'")

            'Sólo creamos la autorización en las empresas de Consolidación
            If oRs.RecordCount > 0 Then
                autorizacionXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_AUCTAEX.xml")
                objGlobal.refDi.comunes.LoadBDFromXML(autorizacionXML)
                res = objglobal.SboApp.GetLastBatchResults
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub

#End Region

#Region "Eventos"
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.MenuUID
                    Case "EXO-MnCtaEx"
                        oForm = objGlobal.SBOApp.Forms.ActiveForm
                        If oForm.TypeEx = "169" Then
                            If EventHandler_Form_Visible() = False Then
                                GC.Collect()
                                Return False
                            End If
                        End If
                End Select
            End If
            Return MyBase.SBOApp_MenuEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function


    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_OCTE"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select

                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_OCTE"

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
                        Case "EXO_OCTE"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_OCTE"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

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

    Private Function EventHandler_Form_Visible() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection = Nothing
        Dim oCFL As SAPbouiCOM.ChooseFromList = Nothing
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams = Nothing

        EventHandler_Form_Visible = False

        Try
            'Recuperar el formulario
            oForm = Me.objGlobal.SBOApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "EXO_OCTE", "")

            If oForm.Visible = True Then
                oForm.Title = "Cuentas contables excluidas para la consolidación"

                If CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("DocEntry").Visible = True Then
                    CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("DocEntry").Visible = False
                End If

                oCFLs = oForm.ChooseFromLists

                oCFLCreationParams = Me.objglobal.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

                oCFLCreationParams.MultiSelection = False
                oCFLCreationParams.ObjectType = "1"
                oCFLCreationParams.UniqueID = "EXO_CFL1"

                oCFL = oCFLs.Add(oCFLCreationParams)

                oCFLCreationParams = Me.objglobal.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

                oCFLCreationParams.MultiSelection = False
                oCFLCreationParams.ObjectType = "1"
                oCFLCreationParams.UniqueID = "EXO_CFL2"

                oCFL = oCFLs.Add(oCFLCreationParams)

                CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_ACCTCODE").ChooseFromListUID = "EXO_CFL1"
                CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_ACCTCODE").ChooseFromListAlias = "AcctCode"
                CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_ACCTNAME").ChooseFromListUID = "EXO_CFL2"
                CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_ACCTNAME").ChooseFromListAlias = "AcctCode"
            End If

            EventHandler_Form_Visible = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCFL, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCFLCreationParams, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCFLs, Object))
            'EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCond, Object))
            'EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oConds, Object))
        End Try
    End Function

    Private Function EventHandler_Choose_FromList_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oCFLEvento As ItemEvent = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_Choose_FromList_Before = False

        Try
            oForm = Me.objglobal.SboApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "3" AndAlso pVal.ColUID = "U_EXO_ACCTCODE" Then
                oCFLEvento = CType(pVal, ItemEvent)

                oConds = New SAPbouiCOM.Conditions

                oCond = oConds.Add
                oCond.Alias = "Postable"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"

                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            ElseIf pVal.ItemUID = "3" AndAlso pVal.ColUID = "U_EXO_ACCTNAME" Then
                oCFLEvento = CType(pVal, ItemEvent)

                oConds = New SAPbouiCOM.Conditions

                oCond = oConds.Add
                oCond.Alias = "Postable"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "Y"

                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            End If

            EventHandler_Choose_FromList_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCFLEvento, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oConds, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCond, Object))
        End Try
    End Function

    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oDataTable As SAPbouiCOM.DataTable = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim iRow As Integer

        EventHandler_Choose_FromList_After = False

        Try
            oForm = Me.objglobal.SboApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing
                GC.Collect()
                Return True
            End If

            oCFLEvento = CType(pVal, ItemEvent)

            Select Case pVal.ItemUID
                Case "3" 'Matrix

                    Select Case pVal.ColUID
                        Case "U_EXO_ACCTCODE"
                            iRow = oCFLEvento.Row

                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                iRow = CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).RowCount
                            End If

                            oDataTable = pVal.SelectedObjects

                            If oDataTable IsNot Nothing Then
                                Try
                                    CType(CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_ACCTCODE").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("AcctCode", 0).ToString
                                Catch ex As Exception

                                End Try

                                Try
                                    CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).SetCellWithoutValidation(iRow, "U_EXO_ACCTNAME", oDataTable.GetValue("AcctName", 0).ToString)
                                Catch ex As Exception

                                End Try

                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If

                        Case "U_EXO_ACCTNAME"
                            iRow = oCFLEvento.Row

                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                iRow = CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).RowCount
                            End If

                            oDataTable = oCFLEvento.SelectedObjects

                            If oDataTable IsNot Nothing Then
                                Try
                                    CType(CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_ACCTCODE").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("AcctCode", 0).ToString
                                Catch ex As Exception

                                End Try

                                Try
                                    CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).SetCellWithoutValidation(iRow, "U_EXO_ACCTNAME", oDataTable.GetValue("AcctName", 0).ToString)
                                Catch ex As Exception

                                End Try

                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If

                    End Select
            End Select

            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCFLEvento, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDataTable, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

#End Region

End Class
