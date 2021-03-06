'Public Class SAP_OPRJ
'    Inherits EXO_UIAPI.EXO_DLLBase

'#Region "Constructor"

'    Public Sub New(ByRef generales As EXO_UIAPI.EXO_UIAPI, actualizar As Boolean)
'        MyBase.New(generales, actualizar)
'    End Sub

'#End Region

'#Region "Inicialización"

'    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
'        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "Filtros_OPRJ.xml")
'        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
'        filtro.LoadFromXML(fXML)
'        Return filtro
'    End Function

'    Public Overrides Function menus() As System.Xml.XmlDocument
'        Return Nothing
'    End Function

'#End Region

'#Region "Eventos"

'    Public Overrides Function objglobal.SboApp.FormDataEvent(ByRef infoEvento As EXO_Generales.EXO_BusinessObjectInfo) As Boolean
'        Try
'            If infoEvento.BeforeAction = True Then
'                Select Case infoEvento.FormTypeEx
'                    Case "711"
'                        Select Case infoEvento.EventType

'                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

'                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

'                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

'                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

'                        End Select

'                End Select

'            Else
'                Select Case infoEvento.FormTypeEx
'                    Case "711"
'                        Select Case infoEvento.EventType

'                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
'                                If infoEvento.ActionSuccess Then
'                                    If GuardarInterCoOPRJ("1", "OPRJ") = False Then
'                                        Return False
'                                    End If
'                                End If

'                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
'                                If infoEvento.ActionSuccess Then
'                                    If GuardarInterCoOPRJ("1", "OPRJ") = False Then
'                                        Return False
'                                    End If
'                                End If

'                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE
'                                If infoEvento.ActionSuccess Then
'                                    If GuardarInterCoOPRJ("1", "OPRJ") = False Then
'                                        Return False
'                                    End If
'                                End If

'                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

'                        End Select

'                End Select

'            End If

'           Return MyBase.SBOApp_FormDataEvent(infoEvento)

'        Catch exCOM As System.Runtime.InteropServices.COMException
'            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
'            Return False
'        Catch ex As Exception
'            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
'            Return False
'        End Try
'    End Function

'#End Region

'#Region "Métodos auxiliares"

'    Private Function GuardarInterCoOPRJ(ByVal sTableCategory As String, ByVal sTableName As String) As Boolean
'        Dim oRs As SAPbobsCOM.Recordset = Nothing
'        Dim oRsAux As SAPbobsCOM.Recordset = Nothing
'        Dim oRsAux2 As SAPbobsCOM.Recordset = Nothing
'        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
'        Dim oNodes As System.Xml.XmlNodeList = Nothing
'        Dim oNode As System.Xml.XmlNode = Nothing
'        Dim oXmlAux As System.Xml.XmlDocument = New System.Xml.XmlDocument
'        Dim oNodesAux As System.Xml.XmlNodeList = Nothing
'        Dim oNodeAux As System.Xml.XmlNode = Nothing

'        GuardarInterCoOPRJ = False

'        Try
'            If EXO_GLOBALES.EmpresaConectadaEsMatriz(objGlobal) = True Then
'                objGlobal.conexionSAP.objglobal.SboApp.StatusBar.SetText("Guardando datos para InterCompany ... Espere por favor ...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

'                oRs = CType(objglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
'                oRsAux = CType(objglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
'                oRsAux2 = CType(objglobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

'                oRs.DoQuery("SELECT dbName " & _
'                            "FROM [INTERCOMPANY].dbo.[DATABASES] WITH (NOLOCK) " & _
'                            "WHERE dbName <> '" & objGlobal.conexionSAP.compañia.CompanyDB & "'")

'                oXml.LoadXml(oRs.GetAsXML())
'                oNodes = oXml.SelectNodes("//row")

'                If oRs.RecordCount > 0 Then
'                    oRsAux2.DoQuery("DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE tableName = '" & sTableName & "' AND dateAdd = '" & dFecha.ToShortDateString & "'")
'                    oRsAux2.DoQuery("DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE tableName = '" & sTableName & "' AND codeTable NOT IN (SELECT PrjCode " & _
'                                                                                                                                            "FROM " & sTableName & " WITH (NOLOCK))")

'                    For i As Integer = 0 To oNodes.Count - 1
'                        oNode = oNodes.Item(i)

'                        oRsAux.DoQuery("SELECT PrjCode AS Codigo " & _
'                                       "FROM " & sTableName & " WITH (NOLOCK) " & _
'                                       "WHERE UpdateDate = '" & dFecha.ToShortDateString & "' " & _
'                                       "OR UpdateDate IS NULL ")

'                        oXmlAux.LoadXml(oRsAux.GetAsXML())
'                        oNodesAux = oXmlAux.SelectNodes("//row")

'                        If oRsAux.RecordCount > 0 Then
'                            For j As Integer = 0 To oNodesAux.Count - 1
'                                oNodeAux = oNodesAux.Item(j)

'                                oRsAux2.DoQuery("SELECT dbNameOrig " & _
'                                                "FROM [INTERCOMPANY].dbo.[REPLICATE] WITH (NOLOCK) " & _
'                                                "WHERE dbNameOrig = '" & objGlobal.conexionSAP.compañia.CompanyDB & "' " & _
'                                                "AND dbNameDest = '" & oNode.SelectSingleNode("dbName").InnerText & "' " & _
'                                                "AND tableCategory = " & sTableCategory & " " & _
'                                                "AND tableName = '" & sTableName & "' " & _
'                                                "AND codeTable = '" & oNodeAux.SelectSingleNode("Codigo").InnerText & "'")

'                                If oRsAux2.RecordCount = 0 Then
'                                    oRsAux2.DoQuery("INSERT INTO [INTERCOMPANY].dbo.[REPLICATE] (dbNameOrig, dbNameDest, tableCategory, tableName, codeTable, dateAdd) VALUES " & _
'                                                    "('" & objGlobal.conexionSAP.compañia.CompanyDB & "', '" & oNode.SelectSingleNode("dbName").InnerText & "' " & _
'                                                    ", " & sTableCategory & ", '" & sTableName & "', '" & oNodeAux.SelectSingleNode("Codigo").InnerText & "', '" & dFecha.ToShortDateString & "')")
'                                End If
'                            Next
'                        End If
'                    Next
'                End If

'                objGlobal.conexionSAP.objglobal.SboApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
'            End If

'            GuardarInterCoOPRJ = True

'        Catch exCOM As System.Runtime.InteropServices.COMException
'            Throw exCOM
'        Catch ex As Exception
'            Throw ex
'        Finally
'            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
'            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux, Object))
'            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux2, Object))
'        End Try
'    End Function

'#End Region

'End Class
