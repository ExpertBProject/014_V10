Imports SAPbouiCOM
Public Class SAP_OPWZ
    Inherits EXO_UIAPI.EXO_DLLBase

#Region "Variables globales"

    Private Shared _ActionSuccess As Boolean = True

#End Region

#Region "Constructor"

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)
    End Sub

#End Region

#Region "Inicialización"

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "Filtros_OPWZ.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function

#End Region

#Region "Eventos"

    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "504"

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
                        Case "504"

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
                        Case "504"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "504"

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

            Return MyBase.SBOApp_ItemEvent(infoEvento)

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
        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument

        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            'Buscar XML de update
            objGlobal.SBOApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Path = objGlobal.refDi.OGEN.pathGeneral & "\01.Pantallas"
            If Path = "" Then
                Return False
            End If

            XmlDoc.Load(Path & "\SAP_" & oForm.BusinessObject.Type & ".srf")
            XmlDoc.SelectSingleNode("Application/forms/action/form/@uid").Value = oForm.UniqueID

            objGlobal.SBOApp.LoadBatchActions(XmlDoc.InnerXml.ToString)

            'Posicionamos campos
            oForm.Items.Item("EXO_001").Top = oForm.Items.Item("8").Top 'Nueva fecha de ejecución de pago
            oForm.Items.Item("EXO_001").Left = oForm.Items.Item("14").Width + oForm.Items.Item("14").Left + 20
            oForm.Items.Item("EXO_002").Top = oForm.Items.Item("14").Top
            oForm.Items.Item("EXO_002").Left = oForm.Items.Item("EXO_001").Width + oForm.Items.Item("EXO_001").Left + 10

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

    Private Function EventHandler_ItemPressed_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "4" Then
                If oForm.PaneLevel = 1 Then
                    If CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.EditText).Value <> "" Then 'Nueva fecha de ejecución de pago distinto de vacío
                        If CType(oForm.Items.Item("14").Specific, SAPbouiCOM.EditText).Value <> CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.EditText).Value Then
                            If CType(oForm.Items.Item("132").Specific, SAPbouiCOM.OptionBtn).Selected = True AndAlso
                               CType(oForm.Items.Item("226").Specific, SAPbouiCOM.CheckBox).Checked = False Then
                                'Check cargar ejecución de pago grabada True, Option Vista de ejecuciones de pago realizadas False
                                _ActionSuccess = True

                                If ComprobarDatos(oForm) = False Then
                                    _ActionSuccess = False

                                    Exit Function
                                End If

                                CType(oForm.Items.Item("14").Specific, SAPbouiCOM.EditText).Active = True
                                CType(oForm.Items.Item("14").Specific, SAPbouiCOM.EditText).Value = CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.EditText).Value
                            End If
                        End If
                    End If
                End If
            End If

            EventHandler_ItemPressed_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            _ActionSuccess = False

            Throw exCOM
        Catch ex As Exception
            _ActionSuccess = False

            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "4" Then
                If oForm.PaneLevel = 2 Then
                    If CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.EditText).Value <> "" Then 'Nueva fecha de ejecución de pago distinto de vacío
                        If CType(oForm.Items.Item("14").Specific, SAPbouiCOM.EditText).Value <> CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.EditText).Value Then
                            If CType(oForm.Items.Item("132").Specific, SAPbouiCOM.OptionBtn).Selected = True AndAlso
                               CType(oForm.Items.Item("226").Specific, SAPbouiCOM.CheckBox).Checked = False Then
                                'Check cargar ejecución de pago grabada True, Option Vista de ejecuciones de pago realizadas False
                                If _ActionSuccess = False Then
                                    oForm.PaneLevel = 1

                                    Exit Function
                                End If

                                If ActualizarDatos(oForm) = False Then
                                    oForm.PaneLevel = 1

                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                ElseIf oForm.PaneLevel = 1 Then
                    If CType(oForm.Items.Item("132").Specific, SAPbouiCOM.OptionBtn).Selected = True AndAlso
                       CType(oForm.Items.Item("226").Specific, SAPbouiCOM.CheckBox).Checked = False Then
                        'Check cargar ejecución de pago grabada True, Option Vista de ejecuciones de pago realizadas False
                        oForm.Items.Item("EXO_001").Visible = True
                        oForm.Items.Item("EXO_002").Visible = True
                    Else
                        oForm.Items.Item("EXO_001").Visible = False
                        oForm.Items.Item("EXO_002").Visible = False
                    End If
                End If
            End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.PaneLevel = 1

            Throw exCOM
        Catch ex As Exception
            oForm.PaneLevel = 1

            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

#End Region

#Region "Métodos auxiliares"

    Private Function ComprobarDatos(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sWizardName As String = ""
        Dim sIdEntry As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim sFecha As String = ""

        ComprobarDatos = False

        Try
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsAux = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT t1.IdNumber " &
                   "FROM OPWZ t1 WITH (NOLOCK) " &
                   "WHERE t1.WizardName = '" & CType(oForm.Items.Item("135").Specific, SAPbouiCOM.EditText).Value & "'"

            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                sIdEntry = oRs.Fields.Item("IdNumber").Value.ToString

                sSQL = "SELECT DISTINCT UPPER(t1.InvCurr) InvCurr, t1.LineRate " &
                       "FROM PWZ3 t1 WITH (NOLOCK) " &
                       "WHERE t1.IdEntry = " & sIdEntry & " " &
                       "AND UPPER(t1.InvCurr) <> 'EUR' "

                oRs.DoQuery(sSQL)

                oXml.LoadXml(oRs.GetAsXML())
                oNodes = oXml.SelectNodes("//row")

                If oRs.RecordCount > 0 Then
                    sFecha = CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.EditText).Value

                    For i As Integer = 0 To oNodes.Count - 1
                        oNode = oNodes.Item(i)

                        oRsAux.DoQuery("SELECT Rate " &
                                       "FROM ORTT WITH (NOLOCK) " &
                                       "WHERE UPPER(Currency) = '" & oNode.SelectSingleNode("InvCurr").InnerText & "' " &
                                       "AND CONVERT(DATE, RateDate, 112) = CONVERT(DATE, '" & sFecha & "', 112)")

                        If oRsAux.RecordCount = 0 Then
                            objGlobal.SBOApp.MessageBox("No se ha definido tipo de cambio para la moneda " & oNode.SelectSingleNode("InvCurr").InnerText & " y fecha " & Right(CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.EditText).Value, 2) & "/" & Mid(CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.EditText).Value, 5, 2) & "/" & Left(CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.EditText).Value, 4) & ".")

                            Exit Function
                        End If
                    Next
                End If
            End If

            ComprobarDatos = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux, Object))
        End Try
    End Function

    Private Function ActualizarDatos(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Dim sWizardName As String = ""
        Dim sIdEntry As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux2 As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing
        Dim sFecha As String = ""
        'Dim cDocTotal As Double = 0
        'Dim sObjType As String = ""

        ActualizarDatos = False

        Try
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsAux = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oRsAux2 = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            sSQL = "SELECT t1.IdNumber " &
                   "FROM OPWZ t1 WITH (NOLOCK) " &
                   "WHERE t1.WizardName = '" & CType(oForm.Items.Item("135").Specific, SAPbouiCOM.EditText).Value & "'"

            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                sIdEntry = oRs.Fields.Item("IdNumber").Value.ToString

                If objGlobal.compañia.InTransaction = True Then
                    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                End If
                objGlobal.compañia.StartTransaction()

                sFecha = CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.EditText).Value

                'Update PmntDate en OPWZ
                oRsAux.DoQuery("UPDATE OPWZ SET PmntDate = CONVERT(DATE, '" & sFecha & "', 112) " &
                               "WHERE IdNumber = " & sIdEntry)

                'Update PymDate en PWZ3
                If CType(oForm.Items.Item("2310000025").Specific, SAPbouiCOM.OptionBtn).Selected = True Then
                    oRsAux.DoQuery("UPDATE PWZ3 SET PymDate = CONVERT(DATE, '" & sFecha & "', 112) " &
                                   "WHERE IdEntry = " & sIdEntry & " " &
                                   "AND Checked = 'Y'")
                End If


                sSQL = "SELECT UPPER(t1.InvCurr) InvCurr, t1.LineRate, t1.ObjType, t1.InvKey " &
                       "FROM PWZ3 t1 WITH (NOLOCK) " &
                       "WHERE t1.IdEntry = " & sIdEntry & " " &
                       "AND UPPER(t1.InvCurr) <> 'EUR' "

                oRs.DoQuery(sSQL)

                oXml.LoadXml(oRs.GetAsXML())
                oNodes = oXml.SelectNodes("//row")

                If oRs.RecordCount > 0 Then
                    For i As Integer = 0 To oNodes.Count - 1
                        oNode = oNodes.Item(i)

                        'cDocTotal = 0

                        'sObjType = oNode.SelectSingleNode("ObjType").InnerText

                        'If sObjType = "13" Then 'Factura de ventas
                        '    oRsAux2.DoQuery("SELECT DocTotalFC " & _
                        '                    "FROM OINV WITH (NOLOCK) " & _
                        '                    "WHERE DocEntry = " & oNode.SelectSingleNode("InvKey").InnerText)

                        '    If oRsAux2.RecordCount > 0 Then
                        '        cDocTotal = CDbl(oRsAux2.Fields.Item("DocTotalFC").Value.ToString.Replace(".", ","))
                        '    End If
                        'ElseIf sObjType = "18" Then 'Factura de compras
                        '    oRsAux2.DoQuery("SELECT DocTotalFC " & _
                        '                    "FROM OPCH WITH (NOLOCK) " & _
                        '                    "WHERE DocEntry = " & oNode.SelectSingleNode("InvKey").InnerText)

                        '    If oRsAux2.RecordCount > 0 Then
                        '        cDocTotal = CDbl(oRsAux2.Fields.Item("DocTotalFC").Value.ToString.Replace(".", ","))
                        '    End If
                        'ElseIf sObjType = "30" Then 'Asiento
                        '    oRsAux2.DoQuery("SELECT FcTotal " & _
                        '                    "FROM OJDT WITH (NOLOCK) " & _
                        '                    "WHERE TransId = " & oNode.SelectSingleNode("InvKey").InnerText)

                        '    If oRsAux2.RecordCount > 0 Then
                        '        cDocTotal = CDbl(oRsAux2.Fields.Item("FcTotal").Value.ToString.Replace(".", ","))
                        '    End If
                        'ElseIf sObjType = "203" Then 'Factura de anticipo ventas
                        '    oRsAux2.DoQuery("SELECT DocTotalFC " & _
                        '                    "FROM ODPI WITH (NOLOCK) " & _
                        '                    "WHERE DocEntry = " & oNode.SelectSingleNode("InvKey").InnerText)

                        '    If oRsAux2.RecordCount > 0 Then
                        '        cDocTotal = CDbl(oRsAux2.Fields.Item("DocTotalFC").Value.ToString.Replace(".", ","))
                        '    End If
                        'ElseIf sObjType = "204" Then 'Factura de anticipo compras
                        '    oRsAux2.DoQuery("SELECT DocTotalFC " & _
                        '                    "FROM ODPO WITH (NOLOCK) " & _
                        '                    "WHERE DocEntry = " & oNode.SelectSingleNode("InvKey").InnerText)

                        '    If oRsAux2.RecordCount > 0 Then
                        '        cDocTotal = CDbl(oRsAux2.Fields.Item("DocTotalFC").Value.ToString.Replace(".", ","))
                        '    End If
                        'ElseIf sObjType = "14" Then 'Abono de ventas
                        '    oRsAux2.DoQuery("SELECT DocTotalFC " & _
                        '                    "FROM ORIN WITH (NOLOCK) " & _
                        '                    "WHERE DocEntry = " & oNode.SelectSingleNode("InvKey").InnerText)

                        '    If oRsAux2.RecordCount > 0 Then
                        '        cDocTotal = CDbl(oRsAux2.Fields.Item("DocTotalFC").Value.ToString.Replace(".", ","))
                        '    End If
                        'ElseIf sObjType = "19" Then 'Abono de compras
                        '    oRsAux2.DoQuery("SELECT DocTotalFC " & _
                        '                    "FROM ORPC WITH (NOLOCK) " & _
                        '                    "WHERE DocEntry = " & oNode.SelectSingleNode("InvKey").InnerText)

                        '    If oRsAux2.RecordCount > 0 Then
                        '        cDocTotal = CDbl(oRsAux2.Fields.Item("DocTotalFC").Value.ToString.Replace(".", ","))
                        '    End If
                        'End If

                        oRsAux2.DoQuery("SELECT Rate " &
                                        "FROM ORTT WITH (NOLOCK) " &
                                        "WHERE UPPER(Currency) = '" & oNode.SelectSingleNode("InvCurr").InnerText & "' " &
                                        "AND CONVERT(DATE, RateDate, 112) = CONVERT(DATE, '" & sFecha & "', 112)")

                        If oRsAux2.RecordCount > 0 Then
                            If CDbl(oNode.SelectSingleNode("LineRate").InnerText.Replace(".", ",")) <> CDbl(oRsAux2.Fields.Item("Rate").Value.ToString.Replace(".", ",")) Then
                                'Update tipo de cambio e importes en PWZ3
                                'oRsAux.DoQuery("UPDATE PWZ3 SET LineRate = " & oRsAux2.Fields.Item("Rate").Value.ToString.Replace(",", ".") & ", " & _
                                '               "PayAmount = " & Math.Round(CDbl(oRsAux2.Fields.Item("Rate").Value.ToString.Replace(".", ",")) * cDocTotal, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero).ToString.Replace(",", ".") & ", " & _
                                '               "PayAmntSys = " & Math.Round(CDbl(oRsAux2.Fields.Item("Rate").Value.ToString.Replace(".", ",")) * cDocTotal, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero).ToString.Replace(",", ".") & ", " & _
                                '               "BcgPmnt = " & Math.Round(CDbl(oRsAux2.Fields.Item("Rate").Value.ToString.Replace(".", ",")) * cDocTotal, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero).ToString.Replace(",", ".") & ", " & _
                                '               "BcgPmntSc = " & Math.Round(CDbl(oRsAux2.Fields.Item("Rate").Value.ToString.Replace(".", ",")) * cDocTotal, objGlobal.conexionSAP.OADM.decimalesImportes, MidpointRounding.AwayFromZero).ToString.Replace(",", ".") & " " & _
                                '               "WHERE IdEntry = " & sIdEntry)

                                oRsAux.DoQuery("UPDATE PWZ3 SET LineRate = " & oRsAux2.Fields.Item("Rate").Value.ToString.Replace(",", ".") & " " &
                                               "WHERE IdEntry = " & sIdEntry)
                            End If
                        End If
                    Next
                End If

                If objGlobal.compañia.InTransaction = True Then
                    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If
            End If

            ActualizarDatos = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            Throw exCOM
        Catch ex As Exception
            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsAux2, Object))
        End Try
    End Function

#End Region

End Class
