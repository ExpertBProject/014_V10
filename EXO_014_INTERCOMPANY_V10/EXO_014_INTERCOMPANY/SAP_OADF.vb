﻿Public Class SAP_OADF
    Inherits EXO_Generales.EXO_DLLBase

#Region "Variables globales"

    Private Shared _sCodigo As String

#End Region

#Region "Constructor"

    Public Sub New(ByRef generales As EXO_Generales.EXO_General, actualizar As Boolean)
        MyBase.New(generales, actualizar)
    End Sub

#End Region

#Region "Inicialización"

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.Functions.leerEmbebido(Me.GetType(), "Filtros_OADF.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function

#End Region

#Region "Eventos"

    Public Overrides Function SBOApp_MenuEvent(ByRef infoEvento As EXO_Generales.EXO_MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
                oForm = SboApp.Forms.ActiveForm

                Select Case oForm.TypeEx
                    Case "419"

                        Select Case infoEvento.MenuUID
                            Case "1283"
                                _sCodigo = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value

                        End Select

                End Select

            Else
                oForm = SboApp.Forms.ActiveForm

                Select Case oForm.TypeEx
                    Case "419"

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

    Public Overrides Function SBOApp_FormDataEvent(ByRef infoEvento As EXO_Generales.EXO_BusinessObjectInfo) As Boolean
        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "419"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select

                End Select

            Else
                Select Case infoEvento.FormTypeEx
                    Case "419"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If infoEvento.ActionSuccess Then
                                    If GuardarInterCoOADF("1", "OADF", _sCodigo) = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then
                                    If GuardarInterCoOADF("1", "OADF", _sCodigo) = False Then
                                        Return False
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE
                                If infoEvento.ActionSuccess Then
                                    If GuardarInterCoOADF("1", "OADF", _sCodigo) = False Then
                                        Return False
                                    End If
                                End If

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
        End Try
    End Function

    Public Overrides Function SBOApp_ItemEvent(ByRef infoEvento As EXO_Generales.EXO_infoItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "419"

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
                        Case "419"

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
                        Case "419"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "419"

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
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function

    Private Function EventHandler_ItemPressed_Before(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_Before = False

        Try
            oForm = SboApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "1" Then
                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE OrElse oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    _sCodigo = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).Value
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

    Private Function GuardarInterCoOADF(ByVal sTableCategory As String, ByVal sTableName As String, ByVal sCodigo As String) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux As SAPbobsCOM.Recordset = Nothing
        Dim oRsAux2 As SAPbobsCOM.Recordset = Nothing
        Dim oXml As System.Xml.XmlDocument = New System.Xml.XmlDocument
        Dim oNodes As System.Xml.XmlNodeList = Nothing
        Dim oNode As System.Xml.XmlNode = Nothing

        GuardarInterCoOADF = False

        Try
            If EXO_GLOBALES.EmpresaConectadaEsMatriz(objGlobal) = True Then
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("Guardando datos para InterCompany ... Espere por favor ...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                oRs = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                oRsAux = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                oRsAux2 = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

                oRs.DoQuery("SELECT dbName " & _
                            "FROM [INTERCOMPANY].dbo.[DATABASES] WITH (NOLOCK) " & _
                            "WHERE dbTipo = 'S' " & _
                            "AND dbName <> '" & objGlobal.conexionSAP.compañia.CompanyDB & "'")

                oXml.LoadXml(oRs.GetAsXML())
                oNodes = oXml.SelectNodes("//row")

                If oRs.RecordCount > 0 Then
                    For i As Integer = 0 To oNodes.Count - 1
                        oNode = oNodes.Item(i)

                        oRsAux.DoQuery("SELECT Code AS Codigo, Name AS Codigo2 " & _
                                       "FROM " & sTableName & " WITH (NOLOCK) " & _
                                       "WHERE Name = '" & sCodigo & "'")

                        If oRsAux.RecordCount > 0 Then
                            oRsAux2.DoQuery("SELECT dbNameOrig " & _
                                            "FROM [INTERCOMPANY].dbo.[REPLICATE] WITH (NOLOCK) " & _
                                            "WHERE dbNameOrig = '" & objGlobal.conexionSAP.compañia.CompanyDB & "' " & _
                                            "AND dbNameDest = '" & oNode.SelectSingleNode("dbName").InnerText & "' " & _
                                            "AND tableCategory = " & sTableCategory & " " & _
                                            "AND tableName = '" & sTableName & "' " & _
                                            "AND codeTable2 = '" & sCodigo & "'")

                            If oRsAux2.RecordCount = 0 Then
                                oRsAux2.DoQuery("INSERT INTO [INTERCOMPANY].dbo.[REPLICATE] (dbNameOrig, dbNameDest, tableCategory, tableName, codeTable, codeTable2, dateAdd) VALUES " & _
                                                "('" & objGlobal.conexionSAP.compañia.CompanyDB & "', '" & oNode.SelectSingleNode("dbName").InnerText & "' " & _
                                                ", " & sTableCategory & ", '" & sTableName & "', '" & oRsAux.Fields.Item("Codigo").Value.ToString & "', '" & sCodigo & "', '" & Now.Year & "-" & Right("0" & Now.Month.ToString, 2) & "-" & Right("0" & Now.Day.ToString, 2) & "')")
                            End If
                        Else
                            oRsAux2.DoQuery("DELETE FROM [INTERCOMPANY].dbo.[REPLICATE] WHERE tableName = '" & sTableName & "' AND codeTable2 = '" & sCodigo & "'")
                        End If
                    Next
                End If

                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            End If

            GuardarInterCoOADF = True

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
