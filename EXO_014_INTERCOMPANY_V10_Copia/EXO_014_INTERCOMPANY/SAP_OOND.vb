Imports SAPbouiCOM
Public Class SAP_OOND
    Inherits EXO_UIAPI.EXO_DLLBase

#Region "Constructor"

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)
    End Sub

#End Region

#Region "Inicialización"

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "Filtros_OOND.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function

#End Region

#Region "Eventos"

    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "60174"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select

                End Select

            Else
                Select Case infoEvento.FormTypeEx
                    Case "60174"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess Then
                                    If GuardarInterCoOOND("1", "OOND") = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then
                                    If GuardarInterCoOOND("1", "OOND") = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE
                                If infoEvento.ActionSuccess Then
                                    If GuardarInterCoOOND("1", "OOND") = False Then
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

#End Region

#Region "Métodos auxiliares"

    Private Function GuardarInterCoOOND(ByVal sTableCategory As String, ByVal sTableName As String) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux2 As SAPbobsCOM.Recordset = Nothing
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim oXmlAux As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodesAux As System.Xml.XmlNodeList = Nothing
        Dim oNodeAux As System.Xml.XmlNode = Nothing

        GuardarInterCoOOND = False

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
                    oRsAux2.DoQuery("DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE tableName = '" & sTableName & "'")

                    For i As Integer = 0 To oNodes.Count - 1
                        oNode = oNodes.Item(i)

                        oRsAux.DoQuery("SELECT IndCode AS Codigo, IndName AS Codigo2 " &
                                       "FROM " & sTableName & " WITH (NOLOCK) ")

                        oXmlAux.LoadXml(oRsAux.GetAsXML())
                        oNodesAux = oXmlAux.SelectNodes("//row")

                        If oRsAux.RecordCount > 0 Then
                            For j As Integer = 0 To oNodesAux.Count - 1
                                oNodeAux = oNodesAux.Item(j)

                                oRsAux2.DoQuery("SELECT dbNameOrig " &
                                                "FROM [INTERCOMPANY].dbo.[REPLICATE] WITH (NOLOCK) " &
                                                "WHERE dbNameOrig = '" & objGlobal.compañia.CompanyDB & "' " &
                                                "AND dbNameDest = '" & oNode.SelectSingleNode("dbName").InnerText & "' " &
                                                "AND tableCategory = " & sTableCategory & " " &
                                                "AND tableName = '" & sTableName & "' " &
                                                "AND codeTable2 = '" & oNodeAux.SelectSingleNode("Codigo2").InnerText & "'")

                                If oRsAux2.RecordCount = 0 Then
                                    Dim dFecha As Date = New Date(Now.Year, Now.Month, Now.Day)
                                    oRsAux2.DoQuery("INSERT INTO [INTERCOMPANY].dbo.[REPLICATE] (dbNameOrig, dbNameDest, tableCategory, tableName, codeTable, codeTable2, dateAdd) VALUES " &
                                                    "('" & objGlobal.compañia.CompanyDB & "', '" & oNode.SelectSingleNode("dbName").InnerText & "' " &
                                                    ", " & sTableCategory & ", '" & sTableName & "', '" & oNodeAux.SelectSingleNode("Codigo").InnerText & "', '" & oNodeAux.SelectSingleNode("Codigo2").InnerText & "', '" & dFecha.ToShortDateString & "')")
                                End If
                            Next
                        End If
                    Next
                End If

                objGlobal.SBOApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            End If

            GuardarInterCoOOND = True

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
