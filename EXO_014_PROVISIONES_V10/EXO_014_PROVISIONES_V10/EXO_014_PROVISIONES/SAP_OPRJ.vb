Public Class SAP_OPRJ
    Inherits EXO_Generales.EXO_DLLBase
    Public Sub New(ByRef generales As EXO_Generales.EXO_General, ByRef actualizar As Boolean)
        MyBase.New(generales, actualizar)

        If actualizar Then
            cargaCampos()
        End If
    End Sub
#Region "Inicialización"

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.Functions.leerEmbebido(Me.GetType(), "Filtros.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As Xml.XmlDocument
        Return Nothing
    End Function

    Private Sub cargaCampos()
        If objGlobal.conexionSAP.esAdministrador Then


            Dim autorizacionXML As String = ""
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            oXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "UDFs_OPRJ.xml")
            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("Validando: UDFs OPRJ   ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.conexionSAP.LoadBDFromXML(oXML)

        End If

    End Sub
#End Region

#Region "Eventos"
    Public Overrides Function SBOApp_ItemEvent(ByRef infoEvento As EXO_Generales.EXO_infoItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "711"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                                Case SAPbouiCOM.BoEventTypes.et_GRID_SORT

                            End Select

                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "711"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select

                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "711"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE


                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "711"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

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

    Private Function EventHandler_Form_Load(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection = Nothing
        Dim oCFL As SAPbouiCOM.ChooseFromList = Nothing
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams = Nothing

        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario

            oForm = SboApp.Forms.Item(pVal.FormUID)

            oCFLs = oForm.ChooseFromLists

            oCFLCreationParams = CType(Me.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "62"
            oCFLCreationParams.UniqueID = "EXO_CFL1"

            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLCreationParams = CType(Me.SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "62"
            oCFLCreationParams.UniqueID = "EXO_CFL2"

            oCFL = oCFLs.Add(oCFLCreationParams)

            CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_OCRCODE").ChooseFromListUID = "EXO_CFL1"
            CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_OCRCODE").ChooseFromListAlias = "OcrCode"
            CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_OCRCODE5").ChooseFromListUID = "EXO_CFL2"
            CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_OCRCODE5").ChooseFromListAlias = "OcrCode"


            EventHandler_Form_Load = True

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

    Private Function EventHandler_Choose_FromList_Before(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oCFLEvento As EXO_Generales.EXO_infoItemEvent = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_Choose_FromList_Before = False

        Try
            oForm = Me.SboApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "3" AndAlso pVal.ColUID = "U_EXO_OCRCODE" Then
                oCFLEvento = CType(pVal, EXO_Generales.EXO_infoItemEvent)

                oConds = New SAPbouiCOM.Conditions

                oCond = oConds.Add
                oCond.Alias = "DimCode"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "1"

                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            ElseIf pVal.ItemUID = "3" AndAlso pVal.ColUID = "U_EXO_OCRCODE5" Then
                oCFLEvento = CType(pVal, EXO_Generales.EXO_infoItemEvent)

                oConds = New SAPbouiCOM.Conditions

                oCond = oConds.Add
                oCond.Alias = "DimCode"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "5"

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

    Private Function EventHandler_Choose_FromList_After(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oCFLEvento As EXO_Generales.EXO_infoItemEvent = Nothing
        Dim oDataTable As EXO_Generales.EXO_infoItemEvent.EXO_SeleccionadosCHFL = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim iRow As Integer

        EventHandler_Choose_FromList_After = False

        Try
            oForm = Me.SboApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing
                GC.Collect()
                Return True
            End If

            oCFLEvento = CType(pVal, EXO_Generales.EXO_infoItemEvent)

            Select Case pVal.ItemUID
                Case "3" 'Matrix

                    Select Case pVal.ColUID
                        Case "U_EXO_OCRCODE"
                            iRow = oCFLEvento.Row
                            oDataTable = pVal.SelectedObjects

                            If oDataTable IsNot Nothing Then
                                Try
                                    CType(CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_OCRCODE").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("OcrCode", 0).ToString
                                Catch ex As Exception

                                End Try

                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If

                        Case "U_EXO_OCRCODE5"
                            iRow = oCFLEvento.Row
                            oDataTable = oCFLEvento.SelectedObjects

                            If oDataTable IsNot Nothing Then
                                Try
                                    CType(CType(oForm.Items.Item("3").Specific, SAPbouiCOM.Matrix).Columns.Item("U_EXO_OCRCODE5").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("OcrCode", 0).ToString
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
