Imports SAPbouiCOM
Public Class SAP_VENTASCOM
    Inherits EXO_UIAPI.EXO_DLLBase
    Private Shared _bItemCodeChanged As Boolean = False
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaCampos()
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
        Return Nothing
    End Function

    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then

            Dim autorizacionXML As String = ""
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

        End If

    End Sub
#End Region

#Region "Eventos"
    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "133", "179", "149", "139", "140", "180", "65303", "142", "143", "182", "141", "181", "392"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                    If EventHandler_Validate_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                                Case SAPbouiCOM.BoEventTypes.et_GRID_SORT

                            End Select

                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "133", "179", "149", "139", "140", "180", "65303", "142", "143", "182", "141", "181", "392"

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
                        Case "133", "179", "149", "139", "140", "180", "65303", "142", "143", "182", "141", "181", "392"

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

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "133", "179", "149", "139", "140", "180", "65303", "142", "143", "182", "141", "181", "392"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

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

    Private Function EventHandler_Validate_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sCodProyecto As String
        Dim strSql As String
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        EventHandler_Validate_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oForm.Freeze(True)

            'ventas y compras
            If pVal.ItemUID = "38" AndAlso (pVal.ColUID = "1" OrElse pVal.ColUID = "3" OrElse pVal.ColUID = "2" OrElse pVal.ColUID = "31") Then 'ItemCode o ItemName o Num. catálogo IC
                If _bItemCodeChanged = True Then ' Por ser el campo choosefromlist hay que utilizar una variable para saber si se ha modificado el campo

                    'Si el proyecto es diferente de vacio, compruebo los datos de normas de reparto
                    sCodProyecto = CStr(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("31").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).String)

                    If Obtener_CentrosCoste(objGlobal, oForm, pVal.Row) = False Then
                        Exit Function
                    End If



                End If
            Else
                'asiento
                If pVal.ItemUID = "76" AndAlso (pVal.ColUID = "16") Then 'proyecto asiento
                    If _bItemCodeChanged = True Then ' Por ser el campo choosefromlist hay que utilizar una variable para saber si se ha modificado el campo

                        'Si el proyecto es diferente de vacio, compruebo los datos de normas de reparto
                        sCodProyecto = CStr(CType(CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("16").Cells.Item(pVal.Row).Specific, SAPbouiCOM.EditText).String)

                        If Obtener_CentrosCoste(objGlobal, oForm, pVal.Row) = False Then
                            Exit Function
                        End If



                    End If
                End If
            End If

            EventHandler_Validate_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            If pVal.ItemUID = "38" AndAlso (pVal.ColUID = "1" OrElse pVal.ColUID = "3" OrElse pVal.ColUID = "2" OrElse pVal.ColUID = "31") Then 'ItemCode o ItemName o Num. catálogo IC
                If _bItemCodeChanged = True Then

                End If
            End If

            _bItemCodeChanged = False


            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oCFLEvento As IChooseFromListEvent = Nothing
        Dim oDataTable As SAPbouiCOM.DataTable = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sCFL_ID As String = ""

        EventHandler_Choose_FromList_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oForm = Nothing
                GC.Collect()
                Return True
            End If

            oCFLEvento = CType(pVal, IChooseFromListEvent)

            Select Case oCFLEvento.ChooseFromListUID
                Case "6", "7", "1", "29", "27", "28", "16" 'Artículo, Descripción o Número de catálogo de IC
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    oDataTable = oCFLEvento.SelectedObjects

                    If oDataTable IsNot Nothing Then
                        If pVal.ItemUID = "38" AndAlso (pVal.ColUID = "1" OrElse pVal.ColUID = "3" OrElse pVal.ColUID = "2" OrElse pVal.ColUID = "31") Then 'ItemCode o ItemName o Num. catálogo IC
                            _bItemCodeChanged = True

                        End If
                        If pVal.ItemUID = "76" AndAlso (pVal.ColUID = "16") Then 'ItemCode o ItemName o Num. catálogo IC
                            _bItemCodeChanged = True

                        End If
                    End If

            End Select

            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDataTable, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
#End Region

#Region "Auxiliares"
    Public Shared Function Obtener_CentrosCoste(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form, ByVal iRow As Integer) As Boolean
        Dim strSql As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sCodProyecto As String = ""
        Dim SCeco As String = ""
        Dim sDime As String = ""
        Try
            Obtener_CentrosCoste = False
            oRs = CType(oObjGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            If oForm.TypeEx = "392" Then
                sCodProyecto = CStr(CType(CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("16").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).String)
            Else
                sCodProyecto = CStr(CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("31").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).String)
            End If


            strSql = "Select PrjCode,PrjName, U_EXO_OCRCODE,U_EXO_OCRCODE5 " _
            & " From OPRJ where PrjCode='" & sCodProyecto & "' "
            oRs.DoQuery(strSql)
            If oRs.RecordCount > 0 Then
                SCeco = oRs.Fields.Item("U_EXO_OCRCODE").Value.ToString
                sDime = oRs.Fields.Item("U_EXO_OCRCODE5").Value.ToString
                If oForm.TypeEx = "392" Then
                    If CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("2006").Width = 0 Then
                        CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("2006").Width = 100
                    End If

                    If CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("2006").Editable = False Then
                        CType(CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("2006").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).Active = True
                        CType(CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("2006").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).String = SCeco
                    Else
                        'CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).SetCellWithoutValidation(iRow, "OcrCode", oRs.Fields.Item("U_EXO_OCRCODE").Value.ToString)

                        CType(CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("2006").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).String = SCeco

                    End If
                Else
                    If CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("2004").Width = 0 Then
                        CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("2004").Width = 100
                    End If

                    If CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("2004").Editable = False Then
                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("2004").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).Active = True
                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("2004").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).String = SCeco
                    Else
                        'CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).SetCellWithoutValidation(iRow, "OcrCode", oRs.Fields.Item("U_EXO_OCRCODE").Value.ToString)

                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("2004").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).String = SCeco

                    End If
                End If

                'compras, que salga solo ceco
                If oForm.TypeEx = "142" OrElse oForm.TypeEx = "143" OrElse oForm.TypeEx = "182" OrElse oForm.TypeEx = "141" OrElse oForm.TypeEx = "181" OrElse oForm.TypeEx = "392" Then
                    sDime = ""
                Else
                    sDime = oRs.Fields.Item("U_EXO_OCRCODE5").Value.ToString
                End If
                If oForm.TypeEx = "392" Then
                    If CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("2005").Width = 0 Then
                        CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("2005").Width = 100
                    End If

                    If CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("2005").Editable = False Then
                        CType(CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("2005").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).Active = True
                        CType(CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("2005").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).String = sDime
                    Else
                        'CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).SetCellWithoutValidation(iRow, "OcrCode", oRs.Fields.Item("U_EXO_OCRCODE").Value.ToString)

                        CType(CType(oForm.Items.Item("76").Specific, SAPbouiCOM.Matrix).Columns.Item("2005").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).String = sDime

                    End If
                Else
                    If CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("2000").Width = 0 Then
                        CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("2000").Width = 100
                    End If

                    If CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("2000").Editable = False Then
                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("2000").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).Active = True
                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("2000").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).String = sDime
                    Else
                        CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("2000").Cells.Item(iRow).Specific, SAPbouiCOM.EditText).String = sDime
                    End If
                End If


                Obtener_CentrosCoste = True
            End If

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
