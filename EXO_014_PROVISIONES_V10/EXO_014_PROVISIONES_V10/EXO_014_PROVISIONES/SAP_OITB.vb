Public Class SAP_OITB
    Inherits EXO_Generales.EXO_DLLBase
    Public Sub New(ByRef generales As EXO_Generales.EXO_General, ByRef actualizar As Boolean)
        MyBase.New(generales, actualizar)

        If actualizar Then
            cargaCampos()
        End If
    End Sub
#Region "Inicialización"

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.Functions.leerEmbebido(Me.GetType(), "Filtros.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As Xml.XmlDocument
        Return Nothing
    End Function

    Private Sub cargaCampos()
        If objGlobal.conexionSAP.esAdministrador Then


            Dim autorizacionXML As String = ""
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            oXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "UDFs_OITB.xml")
            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("Validando: UDFs OITB  ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.conexionSAP.LoadBDFromXML(oXML)

        End If

    End Sub
#End Region
End Class
