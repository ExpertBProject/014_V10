Imports SAPbouiCOM
Public Class SAP_OADM
    Inherits EXO_UIAPI.EXO_DLLBase

#Region "Constructor"

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaDatos()
        End If
    End Sub

#End Region

#Region "Inicialización"

    Private Sub cargaDatos()
        Dim oXML As String = ""

        Try
            If objGlobal.refDi.comunes.esAdministrador Then
                'Campos de Usuario para configuración de InterCompany
                oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OADM.xml")
                objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs Detalles de empresa", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            objglobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objglobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
        End Try
    End Sub

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "Filtros_OADM.xml")
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
                        Case "136"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE

                            End Select

                    End Select

                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "136"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            End Select

                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "136"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_Form_Load(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "136"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

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

    Private Function EventHandler_Form_Load(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument

        EventHandler_Form_Load = False

        Try
            'Recuperar el formulario
            oForm = objglobal.SboApp.Forms.Item(pVal.FormUID)

            'Buscar XML de update
            objglobal.SboApp.StatusBar.SetText("Presentando información...Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Path = objGlobal.refDi.OGEN.pathGeneral & "\01.Pantallas"
            If Path = "" Then
                Return False
            End If

            XmlDoc.Load(Path & "\SAP_" & oForm.BusinessObject.Type & ".srf")
            XmlDoc.SelectSingleNode("Application/forms/action/form/@uid").Value = oForm.UniqueID

            Try
                objglobal.SboApp.LoadBatchActions(XmlDoc.InnerXml.ToString)
            Catch exCOM As System.Runtime.InteropServices.COMException
            Catch ex As Exception
            End Try

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

#End Region

End Class
