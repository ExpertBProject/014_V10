Imports SAPbouiCOM
Public Class SAP_OJDT
    Inherits EXO_UIAPI.EXO_DLLBase

#Region "Constructor"

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)
    End Sub

#End Region

#Region "Inicialización"

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "Filtros_OJDT.xml")
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
        Dim oXml As New Xml.XmlDocument
        Dim sCodigo As String = ""

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.FormTypeEx
                    Case "392"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE

                        End Select

                End Select

            Else
                Select Case infoEvento.FormTypeEx
                    Case "392"
                        Select Case infoEvento.EventType

                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                                If infoEvento.ActionSuccess Then
                                    oXml.LoadXml(infoEvento.ObjectKey)
                                    sCodigo = oXml.SelectSingleNode("JournalEntryParams/JdtNum").InnerText

                                    If GuardarDatos(sCodigo) = False Then
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

#End Region

#Region "Métodos auxiliares"

    Private Function GuardarDatos(ByVal sCodigo As String) As Boolean
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        GuardarDatos = False

        Try
            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            oRs.DoQuery("UPDATE OJDT SET U_EXO_GRUPOEMPRESA = '', " &
                        "U_EXO_COMPANY1 = '', " &
                        "U_EXO_COMPANY2 = '' " &
                        "WHERE TransId = " & sCodigo)

            GuardarDatos = True

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
