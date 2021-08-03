Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_PROVPOR
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaCampos()
            cargaAutorizaciones()
        End If
    End Sub
#Region "Inicialización"

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "Filtros.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As Xml.XmlDocument
        Dim menuXML As String = ""
        Dim res As String = ""
        'Return Nothing

        'Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MENUDEFPROV.xml")
        'Dim menu As Xml.XmlDocument = New Xml.XmlDocument
        'menu.LoadXml(menuXML)
        'Return menu

        menuXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MENUDEFPROV.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        res = objGlobal.SBOApp.GetLastBatchResults

    End Function

    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim sXML As String = ""
            Dim res As String = ""

            'UDO Grupo de comisiones para artículos
            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_PROVPOR.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_PROVPOR", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults
        End If


    End Sub

    'Para definir autorizaciones
    Private Sub cargaAutorizaciones()
        Dim autorizacionXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_AUPROVPOR.xml")
        objGlobal.refDi.comunes.LoadBDFromXML(autorizacionXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
    End Sub

#End Region

#Region "Eventos"
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.MenuUID
                    Case "EXO-MnProvPor"
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_PROVPOR")

                    Case Else
                        oForm = objGlobal.SBOApp.Forms.ActiveForm

                        Select Case oForm.TypeEx

                            Case "UDO_FT_EXO_PROVPOR"

                                Select Case infoEvento.MenuUID

                                End Select

                        End Select

                End Select
            Else
                oForm = objGlobal.SBOApp.Forms.ActiveForm

                Select Case oForm.TypeEx
                    Case "UDO_EXO_PROVPOR"

                        Select Case infoEvento.MenuUID

                        End Select

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
    Private Function EventHandler_Form_Visible(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing

        EventHandler_Form_Visible = False

        Try

            If pVal.ActionSuccess = True Then
                'Recuperar el formulario
                oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)



                If oForm.Visible = True Then
                    'caja codigo interno no visualizar
                    'oForm.Items.Item("0_U_E").Width = 0
                    'oForm.Items.Item("1_U_E").Width = 0
                    If CargarCombos(oForm) = False Then
                        Exit Function
                    End If

                End If


                EventHandler_Form_Visible = True
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_PROVPOR"

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
                        Case "UDO_FT_EXO_PROVPOR"

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
                        Case "UDO_FT_EXO_PROVPOR"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_Form_Visible(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_PROVPOR"

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
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oXml As New Xml.XmlDocument



        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx

                    Case "UDO_FT_EXO_PROVPOR"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD


                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                'If ComprobarDatos(oForm) = False Then
                                '    Return False
                                'End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
                                If ComprobarDatos(oForm) = False Then
                                    Return False
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select

                End Select

            Else

                Select Case infoEvento.FormTypeEx
                    Case "UDO_FT_EXO_PROVPOR"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                                If infoEvento.ActionSuccess = True Then

                                End If




                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                                If infoEvento.ActionSuccess = True Then
                                    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)


                                End If


                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If infoEvento.ActionSuccess = True Then
                                    oForm = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)


                                End If

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
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oDataTable As SAPbouiCOM.DataTable = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_Choose_FromList_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If (pVal.ItemUID = "0_U_G" AndAlso pVal.ColUID = "C_0_2") OrElse (pVal.ItemUID = "1_U_G" AndAlso pVal.ColUID = "C_1_2") Then
                oCFLEvento = CType(pVal, IChooseFromListEvent)

                oConds = New SAPbouiCOM.Conditions

                oCond = oConds.Add
                oCond.Alias = "DimCode"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "1"



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
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing
                GC.Collect()
                Return True
            End If

            oCFLEvento = CType(pVal, IChooseFromListEvent)

            Select Case pVal.ItemUID
                Case "0_U_G" 'Matrix

                    Select Case pVal.ColUID
                        Case "C_0_2"
                            iRow = oCFLEvento.Row


                            oDataTable = oCFLEvento.SelectedObjects

                            If oDataTable IsNot Nothing Then
                                Try
                                    CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("OcrCode", 0).ToString
                                Catch ex As Exception
                                    CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("OcrCode", 0).ToString
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
#Region "Métodos auxiliares"
    Private Function CargarCombos(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sSQL As String = ""
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        CargarCombos = False


        Try
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT ItmsGrpCod,ItmsGrpNam from OITB where Locked='N'"
            oRs.DoQuery(sSQL)

            oXml.LoadXml(oRs.GetAsXML())
            oNodes = oXml.SelectNodes("//row")
            If oRs.RecordCount > 0 Then
                For j As Integer = 0 To oNodes.Count - 1
                    oNode = oNodes.Item(j)
                    CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").ValidValues.Add(oNode.SelectSingleNode("ItmsGrpCod").InnerText, oNode.SelectSingleNode("ItmsGrpNam").InnerText)

                Next

            End If


            'CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Editable

            'CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_4").Editable
            'objGlobal.conexionSAP.refSBOApp.cargaCombo(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.ComboBox).ValidValues, "SELECT SlpCode, SlpName FROM OSLP")

            'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '    CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.ComboBox).Select("-1", SAPbouiCOM.BoSearchKey.psk_ByValue)
            'End If

            CargarCombos = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ComprobarDatos(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSql As String = ""
        Dim dtFechaD As String
        Dim dtFechaH As String
        ComprobarDatos = False

        Try
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            dtFechaD = ((CType(oForm.Items.Item("20_U_E").Specific, SAPbouiCOM.EditText).String.ToString))
            dtFechaH = ((CType(oForm.Items.Item("21_U_E").Specific, SAPbouiCOM.EditText).String))
            dtFechaD = (Format(CDate(dtFechaD), "yyyy/MM/dd"))
            dtFechaH = (Format(CDate(dtFechaH), "yyyy/MM/dd"))
            sSql = "SELECT DocEntry FROM [@EXO_PROVPOR] WHERE ('" & dtFechaD & "' BETWEEN U_EXO_FDESDE And U_EXO_FHASTA) OR ('" & dtFechaH & "' BETWEEN  U_EXO_FDESDE And U_EXO_FHASTA) "
            oRs.DoQuery(sSql)
            If oRs.RecordCount > 0 Then
                objGlobal.SBOApp.MessageBox("Ya existe un registro con las fechas seleccionadas")
                Exit Function
            End If

            ComprobarDatos = True

        Catch ex As Exception
            Throw ex
        Finally
            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRs)
        End Try
    End Function
#End Region
End Class
