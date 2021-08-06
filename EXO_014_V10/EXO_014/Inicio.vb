Imports SAPbouiCOM
Public Class Inicio
    Inherits EXO_UIAPI.EXO_DLLBase

    Private eventosOPDN As EXO_OPDN
    Private eventosASIS_APROV As EXO_ASIS_APROV

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        eventosOPDN = New EXO_OPDN(objGlobal)
        eventosASIS_APROV = New EXO_ASIS_APROV(objGlobal)

        ''Dim contenidoXML As String = objGlobal.Functions.leerEmbebido(Me.GetType(), "UFINDISC.xml")
        ''Me.objGlobal.conexionSAP.refCompañia.LoadBDFromXML(contenidoXML)
        If actualizar Then
            cargaCampos()
        End If

    End Sub

    Public Sub cargaCampos()

        If objGlobal.refDi.comunes.esAdministrador() Then
            objGlobal.compañia.escribeMensaje("El usuario es administrador")
            'Definicion descuentos financieros
            Dim contenidoXML As String

            contenidoXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDF_OPDN.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(contenidoXML)



        Else
            objGlobal.compañia.escribeMensaje("El usuario NO es administrador")
        End If
    End Sub

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters

        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_Filtros.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)
        Return filtro

    End Function

    Public Overrides Function menus() As System.Xml.XmlDocument
        Return Nothing
    End Function

    Public Overrides Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean
        Dim res As Boolean = True
        Dim tipoForm As String = ""
        Try
            tipoForm = infoEvento.FormTypeEx
        Catch ex As Exception
        End Try

        Select Case tipoForm

            Case "143"
                If Not infoEvento.BeforeAction And infoEvento.ActionSuccess Then
                    res = eventosOPDN.SBOApp_FormDataEvent(infoEvento)
                End If
        End Select

        Return res

    End Function

    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Dim res As Boolean = True
        Dim tipoForm As String = ""
        Try
            tipoForm = infoEvento.FormTypeEx
        Catch ex As Exception
        End Try

        Select Case tipoForm
            Case "540010007"
                Try
                    res = eventosASIS_APROV.SBOApp_ItemEvent(infoEvento)

                Catch ex As Exception

                End Try

        End Select
        Return res

    End Function




End Class