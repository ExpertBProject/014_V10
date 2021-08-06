Imports SAPbouiCOM
Public Class SAP_OCDC
    Inherits EXO_UIAPI.EXO_DLLBase

#Region "Variables globales"

    Private Shared _sCodigo As String

#End Region

#Region "Constructor"

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

    End Sub

#End Region

#Region "Inicialización"

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objglobal.funciones.leerEmbebido(Me.GetType(), "Filtros_OCDC.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function

#End Region

#Region "Eventos"

    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
                oForm = objglobal.SBOApp.Forms.ActiveForm

                Select Case oForm.TypeEx
                    Case "653"

                        Select Case infoEvento.MenuUID
                            Case "1283"
                                _sCodigo = CType(oForm.Items.Item("7").Specific, SAPbouiCOM.EditText).Value

                        End Select

                End Select

            Else
                oForm = objglobal.SBOApp.Forms.ActiveForm

                Select Case oForm.TypeEx
                    Case "653"

                        Select Case infoEvento.MenuUID

                        End Select

                End Select

            End If

            Return MyBase.objGlobal.SBOApp.MenuEvent(infoEvento)

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

    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "653"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select

                End Select

            Else
                Select Case infoEvento.FormTypeEx
                    Case "653"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess Then
                                    If GuardarInterCoOCDC("1", "OCDC", _sCodigo) = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then
                                    If GuardarInterCoOCDC("1", "OCDC", _sCodigo) = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE
                                If infoEvento.ActionSuccess Then
                                    If GuardarInterCoOCDC("1", "OCDC", _sCodigo) = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                        End Select

                End Select

            End If

            Return MyBase.objGlobal.SBOApp.FormDataEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)

            Return False
        End Try
    End Function

    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "653"

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
                        Case "653"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

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
                        Case "653"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "653"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                            End Select

                    End Select
                End If
            End If

            Return MyBase.objGlobal.SBOApp.ItemEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function

    Private Function EventHandler_ItemPressed_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_Before = False

        Try
            oForm = objglobal.SboApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "1" Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE OrElse oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    _sCodigo = CType(oForm.Items.Item("7").Specific, SAPbouiCOM.EditText).Value
                End If
            End If

            EventHandler_ItemPressed_Before = True

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

    Private Function GuardarInterCoOCDC(ByVal sTableCategory As String, ByVal sTableName As String, ByVal sCodigo As String) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux2 As SAPbobsCOM.Recordset = Nothing
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing

        GuardarInterCoOCDC = False

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

                        oRsAux.DoQuery("SELECT Code AS Codigo " &
                                       "FROM " & sTableName & " WITH (NOLOCK) " &
                                       "WHERE Code = '" & sCodigo & "'")

                        If oRsAux.RecordCount > 0 Then
                            oRsAux2.DoQuery("SELECT dbNameOrig " &
                                            "FROM [INTERCOMPANY].dbo.[REPLICATE] WITH (NOLOCK) " &
                                            "WHERE dbNameOrig = '" & objGlobal.compañia.CompanyDB & "' " &
                                            "AND dbNameDest = '" & oNode.SelectSingleNode("dbName").InnerText & "' " &
                                            "AND tableCategory = " & sTableCategory & " " &
                                            "AND tableName = '" & sTableName & "' " &
                                            "AND codeTable = '" & sCodigo & "'")

                            If oRsAux2.RecordCount = 0 Then
                                oRsAux2.DoQuery("INSERT INTO [INTERCOMPANY].dbo.[REPLICATE] (dbNameOrig, dbNameDest, tableCategory, tableName, codeTable, dateAdd) VALUES " &
                                                "('" & objGlobal.compañia.CompanyDB & "', '" & oNode.SelectSingleNode("dbName").InnerText & "' " &
                                                ", " & sTableCategory & ", '" & sTableName & "', '" & sCodigo & "', '" & Now.Year & "-" & Right("0" & Now.Month.ToString, 2) & "-" & Right("0" & Now.Day.ToString, 2) & "')")
                            End If
                        Else
                            oRsAux2.DoQuery("DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE tableName = '" & sTableName & "' AND codeTable = '" & sCodigo & "'")
                        End If
                    Next
                End If

                objGlobal.SBOApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            End If

            GuardarInterCoOCDC = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            _sCodigo = ""
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux2, Object))
        End Try
    End Function

#End Region

End Class
