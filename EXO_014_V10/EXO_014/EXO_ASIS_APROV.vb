Imports SAPbouiCOM

Public Class EXO_ASIS_APROV

    'Private refDI As EXO_DIAPI.EXO_DIAPI
    'Private refUI As EXO_UIAPI.EXO_UIAPI
    Private objGlobal As EXO_UIAPI.EXO_UIAPI

    Public Sub New(objGlobal As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objGlobal
        'refDI = objGlobal.compañia
        'refUI = objGlobal.compañia.refSBOApp

    End Sub

    Public Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Dim res As Boolean = True

        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)

        Try

            Select Case infoEvento.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                    If infoEvento.BeforeAction Then

                        Dim oItem As SAPbouiCOM.Item

                        oItem = oForm.Items.Add("txtFCon", BoFormItemTypes.it_EDIT)
                        oItem.Top = oForm.Items.Item("1980000002").Top - 30
                        oItem.Left = oForm.Items.Item("1980000002").Left
                        oItem.Height = oForm.Items.Item("1980000002").Height
                        oItem.Width = oForm.Items.Item("1980000002").Width
                        oItem.FromPane = 4
                        oItem.ToPane = 4

                        oForm.DataSources.UserDataSources.Add("UDSFCon", BoDataType.dt_DATE)
                        CType(oItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "", "UDSFCon")

                        oItem = oForm.Items.Add("lblFCon", BoFormItemTypes.it_STATIC)
                        oItem.Top = oForm.Items.Item("1980000001").Top - 30
                        oItem.Left = oForm.Items.Item("1980000001").Left
                        oItem.Height = oForm.Items.Item("1980000001").Height
                        oItem.Width = oForm.Items.Item("1980000001").Width
                        oItem.LinkTo = "txtFCon"
                        oItem.FromPane = 4
                        oItem.ToPane = 4

                        CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Fecha Contable"

                        oItem = oForm.Items.Add("txtFDoc", BoFormItemTypes.it_EDIT)
                        oItem.Top = oForm.Items.Item("1980000002").Top - 15
                        oItem.Left = oForm.Items.Item("1980000002").Left
                        oItem.Height = oForm.Items.Item("1980000002").Height
                        oItem.Width = oForm.Items.Item("1980000002").Width
                        oItem.FromPane = 4
                        oItem.ToPane = 4

                        oForm.DataSources.UserDataSources.Add("UDSFDoc", BoDataType.dt_DATE)
                        CType(oItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "", "UDSFDoc")


                        'CType(oItem.Specific, SAPbouiCOM.EditText).DataBind.SetBound(True, "OITM", "U_EXO_VTP")

                        oItem = oForm.Items.Add("lblFDoc", BoFormItemTypes.it_STATIC)
                        oItem.Top = oForm.Items.Item("1980000001").Top - 15
                        oItem.Left = oForm.Items.Item("1980000001").Left
                        oItem.Height = oForm.Items.Item("1980000001").Height
                        oItem.Width = oForm.Items.Item("1980000001").Width
                        oItem.LinkTo = "txtFDoc"
                        oItem.FromPane = 4
                        oItem.ToPane = 4

                        CType(oItem.Specific, SAPbouiCOM.StaticText).Caption = "Fecha Documento"

                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If Not infoEvento.BeforeAction Then
                        If oForm.PaneLevel = 8 And infoEvento.ItemUID = "_wiz_next_" Then
                            ActualizarFechas(oForm, oForm.DataSources.UserDataSources.Item("UDSFDoc").Value, oForm.DataSources.UserDataSources.Item("UDSFCon").Value)
                        End If
                    End If
            End Select

        Catch ex As Exception
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try

        Return res

    End Function

    Private Sub ActualizarFechas(oForm As Form, FechaDoc As String, FechaCont As String)

        Dim oMat As SAPbouiCOM.Matrix

        oMat = CType(oForm.Items.Item("540000036").Specific, SAPbouiCOM.Matrix)
        'Dim aaa As String = oMat.SerializeAsXML(BoMatrixXmlSelect.mxs_All)
        Dim Idx As Integer = 0
        Dim XmlDoc As New Xml.XmlDocument
        Dim DocEntry As String = ""

        Dim ocombo As SAPbouiCOM.ComboBox = CType(oForm.Items.Item("540000014").Specific, SAPbouiCOM.ComboBox)

        Dim Objeto As String = ocombo.Selected.Value
        XmlDoc.LoadXml(oMat.SerializeAsXML(BoMatrixXmlSelect.mxs_All))
        'Dim Objeto As String = ocombo.Selected.Value
        Dim XMLNodes As Xml.XmlNodeList = Nothing
        Dim tabla As String = ""
        XMLNodes = XmlDoc.SelectNodes("//Rows/Row/Columns/Column[ID='540000011' and Value > 0]/Value")
        For Idx = 0 To XMLNodes.Count - 1
            'Recupero DocEntry del Objeto Creado

            DocEntry = XMLNodes(Idx).InnerText
           
            Dim oDOC As SAPbobsCOM.Documents = Nothing
            If Objeto = "22" Then
                oDOC = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders), SAPbobsCOM.Documents)
                tabla = "OPOR"
            ElseIf Objeto = "540000006" Then
                oDOC = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations), SAPbobsCOM.Documents)
                tabla = "OPQT"
            ElseIf Objeto = "540010007" Then
                oDOC = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseRequest), SAPbobsCOM.Documents)
                tabla = "OPRQ"
            End If

            Dim sDconum As String = objGlobal.refDi.SQL.executeScalar("select docnum from " + tabla + " where docentry='" + DocEntry + "' ").AsString


            Dim HayCambios As Boolean = False

            oDOC.GetByKey(DocEntry)

            If FechaDoc <> "" Then
                oDOC.TaxDate = FechaDoc
                HayCambios = True
            End If

            If FechaCont <> "" Then
                oDOC.DocDate = FechaCont
                HayCambios = True
            End If

            If HayCambios = True Then
                If oDOC.Update() <> 0 Then

                    objGlobal.SBOApp.StatusBar.SetText("Ocurrió un error actualizando la fecha del documento " + sDconum + ". " + objGlobal.compañia.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else
                    objGlobal.SBOApp.StatusBar.SetText("Se actualizaron las fechas del documento " + sDconum + ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If
            End If
        Next

        objGlobal.SBOApp.StatusBar.SetText("Finalizó proceso de actualización de fechas", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        EXO_CleanCOM.CLiberaCOM.FormMatrix(oMat)

    End Sub


End Class