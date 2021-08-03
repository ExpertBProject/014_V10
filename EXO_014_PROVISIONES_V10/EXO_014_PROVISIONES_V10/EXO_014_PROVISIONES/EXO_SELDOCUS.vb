Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Net
Imports System.Text
Imports SAPbouiCOM

Public Class EXO_SELDOCUS
    Inherits EXO_UIAPI.EXO_DLLBase

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, usaLicencia, idAddOn)

        If actualizar Then
            cargaCampos()
            cargaAutorizaciones()
        End If
    End Sub
#Region "Inicialización"

    Public Overrides Function filtros() As SAPbouiCOM.EventFilters
        Dim fXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "Filtros.xml")
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(fXML)
        Return filtro
    End Function

    Public Overrides Function menus() As Xml.XmlDocument
        Dim menuXML As String = ""
        Dim res As String = ""

        menuXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MENUCONPROV.xml")
        objglobal.SboApp.LoadBatchActions(menuXML)
        res = objglobal.SboApp.GetLastBatchResults
        'Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_MENUCONPROV.xml")
        'Dim menu As Xml.XmlDocument = New Xml.XmlDocument
        'menu.LoadXml(menuXML)
        'Return menu


    End Function

    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim autorizacionXML As String = ""
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_OINV.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDFs OINV ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
        End If
    End Sub

    'Para definir autorizaciones
    Private Sub cargaAutorizaciones()
        Dim autorizacionXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "EXO_AUSELDOCUS.xml")
        objGlobal.refDi.comunes.LoadBDFromXML(autorizacionXML)
        Dim res As String = objglobal.SboApp.GetLastBatchResults
    End Sub

#End Region

#Region "Eventos"
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.MenuUID
                    Case "EXO-MnProvDoc"
                        OpenFormOINVEDI(objGlobal, Me.GetType)
                End Select
            End If
            Return MyBase.SBOApp_MenuEvent(infoEvento)
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
    Public Overrides Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_DOCUSPROV"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_ComboSelect_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
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
                        Case "EXO_DOCUSPROV"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                    If EventHandler_Matrix_Link_Press_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                            End Select

                    End Select
                End If

            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_DOCUSPROV"

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

                                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_DOCUSPROV"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

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
    Private Function EventHandler_Matrix_Link_Press_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_Matrix_Link_Press_Before = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "EXO_GR" Then
                If pVal.ColUID = "DocEntry" Then
                    CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item("DocEntry"), SAPbouiCOM.EditTextColumn).LinkedObjectType = CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).DataTable.GetValue("ObjType", pVal.Row).ToString
                End If
            End If

            EventHandler_Matrix_Link_Press_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim Datos(3) As String
        Dim sCECO, SDime4, sDime5, sProyecto As String
        Dim sComentarios As String = ""

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "btVer" Or pVal.ItemUID = "Check_0" Then
                If pVal.ActionSuccess = True Then
                    'cargarGrid
                    'comprobar si ha metido fechas

                    If CType(oForm.Items.Item("EXO_001").Specific, SAPbouiCOM.EditText).Value = "" Then
                        objGlobal.SBOApp.MessageBox("Antes de consultar los documentos a enviar debe seleccionar una fecha desde.")
                        Exit Function
                    End If

                    If CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.EditText).Value = "" Then
                        objGlobal.SBOApp.MessageBox("Antes de consultar los documentos a enviar debe seleccionar una fecha hasta.")
                        Exit Function
                    End If

                    If CType(oForm.Items.Item("Cmb_0").Specific, SAPbouiCOM.ComboBox).Value = "" Then
                        objGlobal.SBOApp.MessageBox("Antes de consultar los documentos a enviar debe seleccionar un grupo de artículos.")
                        Exit Function
                    End If

                    EXO_SELDOCUS.CargarGrid(objGlobal, oForm)
                End If

            End If

            If pVal.ItemUID = "btGenerar" Then
                If pVal.ActionSuccess = True Then
                    If objGlobal.SBOApp.MessageBox("Se van a generar los asientos contables de los documentos seleccionados. ¿Continuar?", 1, "Aceptar", "Cancelar") = 1 Then

                        objGlobal.SBOApp.StatusBar.SetText("...Preparando datos para la generación del asiento contable...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        If ComprobarDatos(oForm) = False Then
                            Exit Function
                        End If

                        'recorrer para meterlo en un datatable de VBNET
                        If oForm.DataSources.DataTables.Item("DT_GR").Rows.Count > 0 Then
                            Dim DT As System.Data.DataTable
                            Dim sGrupo As String = ""

                            DT = ConvertirDataTableSAP(oForm.DataSources.DataTables.Item("DT_GR"))

                            'creo una estructura para preparar las lineas de asiento a crear
                            Dim Asientos As New Dictionary(Of String, Decimal)

                            Dim AsientosH As New Dictionary(Of String, Decimal)

                            Dim Asientos2 As New Dictionary(Of String, Decimal)



                            For Each row As DataRow In DT.Rows

                                sGrupo = row("Proyecto").ToString & "|" & row("CECO").ToString & "|" & row("Provisiones").ToString & "|" & row("Departamento").ToString
                                Dim testArray() As String = sGrupo.Split(CType("|", Char()))
                                sProyecto = ""
                                sCECO = ""
                                SDime4 = ""
                                sDime5 = ""
                                For i As Integer = 0 To testArray.Length - 1
                                    If testArray(i) <> "" Then
                                        Select Case i
                                            Case 0
                                                sProyecto = testArray(i)
                                            Case 1
                                                sCECO = testArray(i)
                                            Case 2
                                                SDime4 = testArray(i)
                                            Case 3
                                                sDime5 = testArray(i)
                                        End Select
                                    End If
                                Next

                                Dim dblImpProv As Double = CDbl((DT.Compute("SUM(ImporteProvision)", "Proyecto='" & sProyecto & "' and CECO='" & sCECO & "' and Provisiones='" & SDime4 & "' and Departamento='" & sDime5 & "'")))
                                If Asientos.ContainsKey(sGrupo) Then
                                Else
                                    Asientos.Add(sGrupo, CDec(dblImpProv))
                                End If

                                If row("Clase").ToString <> "REFACTURACION" Then
                                    Dim dblImpProvMax As Double = CDbl((DT.Compute("SUM(Asiento)", "Proyecto='" & sProyecto & "' and CECO='" & sCECO & "' and Provisiones='" & SDime4 & "' and Departamento='" & sDime5 & "'")))
                                    'Dim dblImpProvMax As Double = CDbl((DT.Compute("SUM(ImporteProvisionMax)", "Proyecto='" & sProyecto & "' and CECO='" & sCECO & "' and Provisiones='" & SDime4 & "' and Departamento='" & sDime5 & "'")))
                                    If Asientos2.ContainsKey(sGrupo) Then
                                    Else
                                        Asientos2.Add(sGrupo, CDec(dblImpProvMax))
                                    End If
                                End If
                            Next

                            Try
                                If objGlobal.compañia.InTransaction = True Then
                                    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                End If
                                objGlobal.compañia.StartTransaction()

                                'generar el asiento
                                sComentarios = "PROVISIONES GRUPO " & DT.Rows.Item(0).Item("Grupo").ToString & " - CLASE: " & DT.Rows.Item(0).Item("Clase").ToString
                                AddOJDT(DT, Asientos, DT.Rows.Item(0).Item("ExpensesAc").ToString, DT.Rows.Item(0).Item("TransferAc").ToString, CDbl((DT.Compute("SUM(ImporteProvision)", ""))), sComentarios, CType(oForm.Items.Item("EXO_003").Specific, SAPbouiCOM.EditText).Value.ToString)

                                'para el segundo asiento
                                If Asientos2.Count > 0 Then
                                    sComentarios = "PROVISIONES MÁXIMA GRUPO " & DT.Rows.Item(0).Item("Grupo").ToString & " - CLASE: " & DT.Rows.Item(0).Item("Clase").ToString
                                    AddOJDT(DT, Asientos2, DT.Rows.Item(0).Item("U_EXO_CTAGASTO").ToString, DT.Rows.Item(0).Item("U_EXO_CTAPROV").ToString, CDbl((DT.Compute("SUM(Asiento)", ""))), sComentarios, CType(oForm.Items.Item("EXO_003").Specific, SAPbouiCOM.EditText).Value.ToString)
                                End If


                                If objGlobal.compañia.InTransaction = True Then
                                    objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                End If

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

                            End Try
                            CargarGrid(objGlobal, oForm)
                        End If

                    End If
                End If
            End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Private Function EventHandler_ComboSelect_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim Valor As String = ""
        EventHandler_ComboSelect_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)


            Select Case pVal.ItemUID

                Case "Cmb_0"
                    'Habilitamos o no combo tipo contenedor
                    If pVal.ActionSuccess = True And pVal.ItemChanged = True Then
                        CargarGrid(objGlobal, oForm)
                    End If

            End Select

            EventHandler_ComboSelect_After = True

        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
#End Region
#Region "Metodos auxiliares"
    Public Shared Function OpenFormOINVEDI(ByRef OGlobal As EXO_UIAPI.EXO_UIAPI, ByRef Type As Type) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        OpenFormOINVEDI = False
        Try

            'abrir formulario
            oFP = CType(OGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = OGlobal.leerEmbebido(Type.GetType(), "EXO_SELDOCUS.srf")

            oRs = CType(OGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Try
                oForm = OGlobal.SBOApp.Forms.AddEx(oFP)

            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    OGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try


            'cargar combo grupo de artículos
            sSQL = "SELECT ItmsGrpCod,ItmsGrpNam FROM OITB WHERE Locked='N'"
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                OGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("Cmb_0").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            End If
            oForm.Visible = True
            CType(oForm.Items.Item("EXO_001").Specific, SAPbouiCOM.EditText).Active = True

        Catch ex As Exception
            Throw ex
        Finally

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

        End Try
    End Function

    Public Shared Sub CargarGrid(ByRef OGlobal As EXO_UIAPI.EXO_UIAPI, ByRef oForm As SAPbouiCOM.Form)

        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        'Dim oForm As SAPbouiCOMº.Form = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim strFechaD As String = ""
        Dim strFechaH As String = ""
        Dim strMarcar As String = ""
        Dim strFacTralix As String = ""
        Dim strPagosTralix As String = ""
        Dim strEnvTralix As String = ""
        Dim strClaseG As String = "COMPRAVENTA,REDENCION,REFACTURACION"

        Try
            OGlobal.SBOApp.StatusBar.SetText("Por favor, espere: cargando datos de selección", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oForm.Freeze(True)

            'cargar consulta datos formulario edi
            oRs = CType(OGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

            strFechaD = CType(oForm.Items.Item("EXO_001").Specific, SAPbouiCOM.EditText).Value
            strFechaH = CType(oForm.Items.Item("EXO_002").Specific, SAPbouiCOM.EditText).Value
            If CType(oForm.Items.Item("Check_0").Specific, SAPbouiCOM.CheckBox).Checked = True Then
                strMarcar = "Y"
            Else
                strMarcar = "N"
            End If
            '    & " WHEN 'REDENCION' THEN  ( (t2.Quantity * COALESCE(t8.Price,0)) * CASE WHEN t7.U_EXO_PORCEN IS NULL  THEN COALESCE(t6.U_EXO_PORCEN,0)  ELSE  COALESCE(t7.U_EXO_PORCEN,0) END /100) " _
            '& " Case When t9.U_EXO_PORCEN Is Not NULL  Then COALESCE(t9.U_EXO_PORCEN,0) " _
            '   & "   WHEN t6.U_EXO_PORCEN Is Not NULL  THEN COALESCE(t6.U_EXO_PORCEN,0)" _
            '   & "   WHEN t7.U_EXO_PORCEN Is Not NULL  THEN COALESCE(t7.U_EXO_PORCEN,0)  " _
            '   & "   End /100) " _
            sSQL = "SELECT '" & strMarcar & "' Sel, 'FAC' Tipo,t0.DocEntry,t1.SeriesName, t0.DocNum,t0.ObjType,t2.ItemCode Artículo, t2.Dscription Descripcion,t3.ItmsGrpCod Grupo, t4.U_EXO_SWW6 Clase," _
                & " t2.OcrCode CECO, COALESCE(t2.OcrCode4,'') Provisiones,COALESCE(t2.OcrCode5,'') Departamento, t2.Project Proyecto , COALESCE(t2.LineTotal,0)  TotalLinea ," _
                & " t2.Quantity  Cantidad,  COALESCE(t6.U_EXO_PORCEN,0) PocenGrupo, COALESCE(t7.U_EXO_PORCEN,0) PorcenArt, ROUND(COALESCE(t8.Price,0),2) PrecioLista , " _
                & " CASE t4.U_EXO_SWW6 " _
                & " When 'COMPRAVENTA' THEN ( (t2.Quantity * ROUND(COALESCE(t8.Price,0),2)) *  Case When t6.U_EXO_PORCEN Is Not NULL  Then COALESCE(t6.U_EXO_PORCEN,0) " _
                & " WHEN   t7.U_EXO_PORCEN Is Not NULL  THEN COALESCE(t7.U_EXO_PORCEN,0) End /100) " _
                & " When 'REDENCION' THEN  ( (t2.Quantity * ROUND(COALESCE(t8.Price,0),2)) *  Case When t6.U_EXO_PORCEN Is Not NULL  Then COALESCE(t6.U_EXO_PORCEN,0) " _
                & " WHEN   t7.U_EXO_PORCEN Is Not NULL  THEN COALESCE(t7.U_EXO_PORCEN,0) End /100) " _
                & " When 'REFACTURACION' THEN   (t2.LineTotal *100 /100) " _
                & " END  ImporteProvision, " _
                & " CASE t4.U_EXO_SWW6 " _
                & " When 'COMPRAVENTA' THEN t2.Quantity * ROUND(COALESCE(t8.Price,0),2) " _
                & " WHEN 'REDENCION' THEN  t2.Quantity *ROUND(COALESCE(t8.Price,0),2)  " _
                & " END  ImporteProvisionMax, " _
                & " CASE t4.U_EXO_SWW6  When 'COMPRAVENTA' THEN t2.Quantity * ROUND(COALESCE(t8.Price,0),2) " _
                & " When 'REDENCION' THEN  t2.Quantity * ROUND(COALESCE(t8.Price,0),2)   END - " _
                & " CASE t4.U_EXO_SWW6 " _
                & " When 'COMPRAVENTA' THEN  ( (t2.Quantity * ROUND(COALESCE(t8.Price,0),2)) *  Case When t6.U_EXO_PORCEN Is Not NULL  Then COALESCE(t6.U_EXO_PORCEN,0) " _
                & " When   t7.U_EXO_PORCEN Is Not NULL  Then COALESCE(t7.U_EXO_PORCEN,0) End /100)  " _
                & " When 'REDENCION' THEN  ( (t2.Quantity * ROUND(COALESCE(t8.Price,0),2)) *  Case When t6.U_EXO_PORCEN Is Not NULL  Then COALESCE(t6.U_EXO_PORCEN,0) " _
                & " When   t7.U_EXO_PORCEN Is Not NULL  Then COALESCE(t7.U_EXO_PORCEN,0) End /100)  " _
                & "  When 'REFACTURACION' THEN   (t2.LineTotal *100 /100)  END  Asiento," _
                & " COALESCE(t2.U_EXO_GENPROV,'N') Estado,'1' Orden,t2.VisOrder ,t4.ExpensesAc,t4.TransferAc,t2.LineNum,t4.U_EXO_CTAGASTO, T4.U_EXO_CTAPROV" _
                & " FROM " _
                & " OINV t0" _
                & " LEFT OUTER JOIN NNM1 t1 on t0.Series = t1.Series" _
                & " INNER JOIN INV1 t2 on t2.DocEntry= t0.DocEntry" _
                & " INNER JOIN OITM t3 On t3.ItemCode= t2.ItemCode" _
                & " INNER JOIN OITB t4 on t4.ItmsGrpCod = t3.ItmsGrpCod" _
                & " LEFT OUTER JOIN [@EXO_PROVPOR] t5 on t0.TaxDate between t5.U_EXO_FDESDE And t5.U_EXO_FHASTA" _
                & " LEFT OUTER JOIN [@EXO_PROVPOR1] t6 on t6.DocEntry=t5.DocEntry and t6.U_EXO_GRUPO = t3.ItmsGrpCod and t6.U_EXO_OCRCODE = t2.OcrCode" _
                & " LEFT OUTER JOIN [@EXO_PROVPOR2] t7 on t7.DocEntry=t5.DocEntry And t7.U_EXO_CODART = t2.ItemCode  and t7.U_EXO_OCRCODE = t2.OcrCode" _
                & " LEFT OUTER JOIN ITM1 t8 on t2.ItemCode = t8.ItemCode and t8.PriceList=2" _
                & " where COALESCE(t2.U_EXO_GENPROV,'N') = 'N'  and t0.TaxDate between '" & strFechaD & "' and '" & strFechaH & "' and (t0.CANCELED='N' OR t0.CANCELED='C') " _
                & " And t3.ItmsGrpCod ='" & CType(oForm.Items.Item("Cmb_0").Specific, SAPbouiCOM.ComboBox).Value & "' AND T4.U_EXO_SWW6 IN ('COMPRAVENTA','REDENCION','REFACTURACION')"
            If sSQL <> "" Then
                sSQL = sSQL & "   UNION ALL "
            End If
            sSQL = sSQL & " Select '" & strMarcar & "' Sel, 'ABO' Tipo,t0.DocEntry,t1.SeriesName, t0.DocNum,t0.ObjType, t2.ItemCode Artículo, t2.Dscription Descripcion,t3.ItmsGrpCod Grupo , t4.U_EXO_SWW6 Clase," _
               & " t2.OcrCode CECO, COALESCE(t2.OcrCode4,'') Provisiones,COALESCE(t2.OcrCode5,'') Departamento, t2.Project Proyecto , COALESCE(t2.LineTotal,0)  * -1 TotalLinea ," _
                & " t2.Quantity  Cantidad,  COALESCE(t6.U_EXO_PORCEN,0) PocenGrupo, COALESCE(t7.U_EXO_PORCEN,0) PorcenArt, ROUND(COALESCE(t8.Price,0),2) PrecioLista , " _
                & " CASE t4.U_EXO_SWW6 " _
                & " When 'COMPRAVENTA' THEN ( (t2.Quantity * ROUND(COALESCE(t8.Price,0),2)) *  Case When t6.U_EXO_PORCEN Is Not NULL  Then COALESCE(t6.U_EXO_PORCEN,0) " _
                & "  WHEN   t7.U_EXO_PORCEN Is Not NULL  THEN COALESCE(t7.U_EXO_PORCEN,0) End /100)   * -1  " _
                 & " When 'REDENCION' THEN  ( (t2.Quantity * ROUND(COALESCE(t8.Price,0),2)) *  Case When t6.U_EXO_PORCEN Is Not NULL  Then COALESCE(t6.U_EXO_PORCEN,0) " _
                & "  WHEN   t7.U_EXO_PORCEN Is Not NULL  THEN COALESCE(t7.U_EXO_PORCEN,0) End /100)   * -1  " _
                & " When 'REFACTURACION' THEN  t2.LineTotal + (t2.LineTotal *100 /100)  * -1" _
                & " END  ImporteProvision, " _
                & " CASE t4.U_EXO_SWW6 " _
                & " When 'COMPRAVENTA' THEN (t2.Quantity * ROUND(COALESCE(t8.Price,0),2)) * -1 " _
                & " WHEN 'REDENCION' THEN  (t2.Quantity * ROUND(COALESCE(t8.Price,0),2)) * -1  " _
                & " END  ImporteProvisionMax, " _
                & " (CASE t4.U_EXO_SWW6  When 'COMPRAVENTA' THEN t2.Quantity * ROUND(COALESCE(t8.Price,0),2) " _
                & " When 'REDENCION' THEN  t2.Quantity * ROUND(COALESCE(t8.Price,0),2)   END - " _
                & " Case t4.U_EXO_SWW6" _
                & " When 'COMPRAVENTA' THEN ( (t2.Quantity * ROUND(COALESCE(t8.Price,0),2)) *  Case When t6.U_EXO_PORCEN Is Not NULL  Then COALESCE(t6.U_EXO_PORCEN,0) " _
                & " When t7.U_EXO_PORCEN Is Not NULL  Then COALESCE(t7.U_EXO_PORCEN,0) End /100)  " _
                & " When 'REDENCION' THEN  ( (t2.Quantity * ROUND(COALESCE(t8.Price,0),2)) *  Case When t6.U_EXO_PORCEN Is Not NULL  Then COALESCE(t6.U_EXO_PORCEN,0) " _
                & " When t7.U_EXO_PORCEN Is Not NULL  Then COALESCE(t7.U_EXO_PORCEN,0) End /100)  " _
                & "  When 'REFACTURACION' THEN   (t2.LineTotal *100 /100)  END) *(-1)     Asiento ," _
                & " COALESCE(t2.U_EXO_GENPROV,'N') Estado, '2' Orden,t2.VisOrder,t4.ExpensesAc,t4.TransferAc,t2.LineNum,t4.U_EXO_CTAGASTO, T4.U_EXO_CTAPROV " _
                & " FROM " _
                & " ORIN t0" _
                & " LEFT OUTER JOIN NNM1 t1 on t0.Series = t1.Series" _
                & " INNER JOIN RIN1 t2 on t2.DocEntry= t0.DocEntry" _
                & " INNER JOIN OITM t3 On t3.ItemCode= t2.ItemCode" _
                & " INNER JOIN OITB t4 on t4.ItmsGrpCod = t3.ItmsGrpCod" _
                & " LEFT OUTER JOIN [@EXO_PROVPOR] t5 on t0.TaxDate between t5.U_EXO_FDESDE And t5.U_EXO_FHASTA" _
               & " LEFT OUTER JOIN [@EXO_PROVPOR1] t6 on t6.DocEntry=t5.DocEntry and t6.U_EXO_GRUPO = t3.ItmsGrpCod and t6.U_EXO_OCRCODE = t2.OcrCode" _
                 & " LEFT OUTER JOIN [@EXO_PROVPOR2] t7 on t7.DocEntry=t5.DocEntry And t7.U_EXO_CODART = t2.ItemCode  and t7.U_EXO_OCRCODE = t2.OcrCode" _
                & " LEFT OUTER JOIN ITM1 t8 on t2.ItemCode = t8.ItemCode and t8.PriceList=2" _
                & " where COALESCE(t2.U_EXO_GENPROV,'N') = 'N' and t0.TaxDate between '" & strFechaD & "' and '" & strFechaH & "' and (t0.CANCELED='N' OR t0.CANCELED='C')  " _
                & " And t3.ItmsGrpCod ='" & CType(oForm.Items.Item("Cmb_0").Specific, SAPbouiCOM.ComboBox).Value & "' AND T4.U_EXO_SWW6 IN ('COMPRAVENTA','REDENCION','REFACTURACION')"
            sSQL = sSQL & "ORDER BY Orden,DocNum,VisOrder"

            oRs.DoQuery(sSQL)

            'columnas
            'Permitimos ordenación por columnas
            'oForm = OGlobal.conexionSAP.SBOApp.Forms.AddEx(oFP)

            oForm.DataSources.DataTables.Item("DT_GR").ExecuteQuery(sSQL)

            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(1).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(2).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(3).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(4).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(5).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(6).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(7).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(8).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(9).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(10).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(11).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(12).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(13).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(14).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(15).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(16).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(17).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(18).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(19).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(20).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(21).TitleObject.Sortable = True
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(22).TitleObject.Sortable = True
            'CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(23).Visible = False
            'CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(24).Visible = False
            'CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(25).Visible = False
            'CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(26).Visible = False
            'CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(27).Visible = False

            'formato columnas
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(0).AffectsFormMode = False
            oColumnChk = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(1), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Tipo"


            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.EditTextColumn)
            oColumnTxt.LinkedObjectType = "13"
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Num. int. documento"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(3), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Serie"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(4), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Num. documento"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(5), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Visible = False
            oColumnTxt.TitleObject.Caption = "ObjType"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(6), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.LinkedObjectType = "4"
            oColumnTxt.TitleObject.Caption = "Artículo"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(7), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Descripción"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(8), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Grupo"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(9), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Clase Grupo"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(10), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "CECO"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(11), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Provisiones"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(12), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Departamento"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(13), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Proyecto"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(14), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Total Linea"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(15), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Cantidad"


            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(16), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "% Grupo"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(17), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "% Art."

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(18), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Precio Lista"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(19), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Importe Provisión"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(20), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Importe Prov. Máxima"

            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(21), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Importe Asiento Prov. Máxima"


            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(22), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.Visible = False
            oColumnTxt.TitleObject.Caption = "Estado"

            'Orden
            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(23), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Visible = False

            'VisOrder
            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(24), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Visible = False

            'ExpensesAc
            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(25), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Cta.Gasto"

            'TransferAc
            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(26), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Cta.Prov"

            'LineNum
            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(27), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Visible = False
            oForm.Freeze(False)

            'ExpensesAc
            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(28), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Cta.Gasto Max."

            'TransferAc
            oColumnTxt = CType(CType(oForm.Items.Item("EXO_GR").Specific, SAPbouiCOM.Grid).Columns.Item(29), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.TitleObject.Caption = "Cta.Prov. Max."

            't2.VisOrder,t4.ExpensesAc,t4.TransferAc,t2.LineNum,t4.U_EXO_CTAGASTO, T4.U_EXO_CTAPROV 

        Catch ex As Exception
            oForm.Freeze(False)
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation)
        Finally
            oForm.Freeze(False)
        End Try

    End Sub

    Public Function ComprobarDatos(ByRef oForm As SAPbouiCOM.Form) As Boolean
        Try
            ComprobarDatos = False
            If oForm.DataSources.DataTables.Item("DT_GR").Rows.Count > 0 Then
                If CType(oForm.Items.Item("EXO_003").Specific, SAPbouiCOM.EditText).Value = "" Then
                    objGlobal.SBOApp.MessageBox("Asigne la fecha de generación del asiento contable")
                    Exit Function
                End If

                For i As Integer = 0 To oForm.DataSources.DataTables.Item("DT_GR").Rows.Count - 1
                    If oForm.DataSources.DataTables.Item("DT_GR").GetValue("ExpensesAc", i).ToString = "" Then
                        'es obligatoria la cuenta
                        objGlobal.SBOApp.MessageBox("Introduza una cuenta de Gastos para ese Grupo de Artículos")
                        Exit Function
                    End If

                    If oForm.DataSources.DataTables.Item("DT_GR").GetValue("TransferAc", i).ToString = "" Then
                        'es obligatoria la cuenta
                        objGlobal.SBOApp.MessageBox("Introduza una cuenta de dotación para ese Grupo de Artículos")
                        Exit Function

                    End If

                    'si la clase es compraventa o redencion, no dejo continuar
                    If oForm.DataSources.DataTables.Item("DT_GR").GetValue("Clase", i).ToString = "COMPRAVENTA" OrElse oForm.DataSources.DataTables.Item("DT_GR").GetValue("Clase", i).ToString = "REDENCION" Then
                        If oForm.DataSources.DataTables.Item("DT_GR").GetValue("U_EXO_CTAGASTO", i).ToString = "" Then
                            'es obligatoria la cuenta
                            objGlobal.SBOApp.MessageBox("Introduza una cuenta de Gasto Máximo para ese Grupo de Artículos")
                            Exit Function
                        End If

                        If oForm.DataSources.DataTables.Item("DT_GR").GetValue("U_EXO_CTAPROV", i).ToString = "" Then
                            'es obligatoria la cuenta
                            objGlobal.SBOApp.MessageBox("Introduza una cuenta de Provisión Máxima para ese Grupo de Artículos")
                            Exit Function
                        End If
                    End If
                Next
                ComprobarDatos = True
            End If
        Catch ex As Exception
            Throw ex
        Finally


        End Try
    End Function
    Public Shared Function SelectDataTable(ByVal dt As System.Data.DataTable, ByVal filter As String, ByVal sort As String) As Data.DataTable

        Dim rows As DataRow()

        Dim dtNew As System.Data.DataTable

        ' copy table structure
        dtNew = dt.Clone()

        ' sort and filter data
        rows = dt.Select(filter, sort)

        ' fill dtNew with selected rows

        For Each dr As DataRow In rows
            dtNew.ImportRow(dr)

        Next

        ' return filtered dt
        Return dtNew

    End Function

    Public Function ConvertirDataTableSAP(ByVal SAPDataTable As SAPbouiCOM.DataTable) As Data.DataTable

        '\ This function will take an SAP DataTable from the SAPbouiCOM library and convert it to a more
        '\ easily used ADO.NET datatable which can be used for data binding much easier.

        Dim dtTable As New Data.DataTable
        Dim NewCol As Data.DataColumn
        Dim NewRow As DataRow
        Dim ColCount As Integer
        Dim bolAdd As Boolean = False


        Try
            objGlobal.SBOApp.StatusBar.SetText("...Preparando datos para la generación del asiento contable...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            For ColCount = 0 To SAPDataTable.Columns.Count - 1
                NewCol = New Data.DataColumn(SAPDataTable.Columns.Item(ColCount).Name)
                'If SAPDataTable.Columns.Item(ColCount).Name = "TotalLinea" OrElse SAPDataTable.Columns.Item(ColCount).Name = "ImporteProvision" OrElse SAPDataTable.Columns.Item(ColCount).Name = "ImporteProvisionMax" OrElse SAPDataTable.Columns.Item(ColCount).Name = "Asiento" Then
                If SAPDataTable.Columns.Item(ColCount).Name = "TotalLinea" OrElse SAPDataTable.Columns.Item(ColCount).Name = "ImporteProvision" OrElse SAPDataTable.Columns.Item(ColCount).Name = "ImporteProvisionMax" OrElse SAPDataTable.Columns.Item(ColCount).Name = "Asiento" Then
                    With NewCol
                        .DataType = System.Type.GetType("System.Double")
                        .AllowDBNull = False

                    End With
                End If
                dtTable.Columns.Add(NewCol)
            Next

            For i = 0 To SAPDataTable.Rows.Count - 1
                'populate each column in the row we're creating
                For ColCount = 0 To SAPDataTable.Columns.Count - 1
                    If SAPDataTable.GetValue(0, i).ToString = "Y" Then
                        If ColCount = 0 Then
                            NewRow = dtTable.NewRow
                            bolAdd = True
                        End If
                        NewRow.Item(SAPDataTable.Columns.Item(ColCount).Name) = SAPDataTable.GetValue(ColCount, i)
                    End If
                Next

                'Add the row to the datatable
                If bolAdd = True Then
                    dtTable.Rows.Add(NewRow)
                End If
                bolAdd = False
            Next

            Return dtTable

        Catch ex As Exception
            MsgBox(ex.ToString & Chr(10) & "Error converting SAP DataTable to DataTable .Net", MsgBoxStyle.Exclamation)
            ConvertirDataTableSAP = Nothing
            Exit Function
        End Try

    End Function


#End Region
#Region "Objetos SAP"
    Public Sub AddOJDT(ByVal dtDatos As System.Data.DataTable, ByVal Asientos As Dictionary(Of String, Decimal), ByVal sCtaDebe As String, ByVal SCtaHaber As String, ByVal dblImpHaber As Double, ByVal sComentarios As String, ByVal sFecha As String)
        Dim oOJDT As SAPbobsCOM.JournalEntries = Nothing
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sTransId As String = "0"
        'Dim sFecha As String = ""
        Dim i As Integer = 0
        Dim sGrupo As String = ""
        Dim dblImpDebe As Double = 0
        'Dim dblImpHaber As Double = 0
        Dim sCECO, SDime4, sDime5, sProyecto As String
        Dim bolAsiento As Boolean = False

        Dim sSql As String = ""

        Try

            oRs = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oOJDT = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries), SAPbobsCOM.JournalEntries)
            'If dblImpHaber > 0 Then

            'sFecha = Date.Now.ToShortDateString.ToString
            sFecha = Mid(sFecha, 1, 4) & "/" & Mid(sFecha, 5, 2) & "/" & Mid(sFecha, 7, 2)
            oOJDT.ReferenceDate = CDate(sFecha)
            oOJDT.TaxDate = CDate(sFecha)
            oOJDT.DueDate = CDate(sFecha)
            oOJDT.Memo = Left(sComentarios.ToString, 50)
            'oOJDT.Memo = "PROVISIONES GRUPO " & dtDatos.Rows.Item(0).Item("Grupo").ToString & " - CLASE: " & dtDatos.Rows.Item(0).Item("Clase").ToString
            oOJDT.AutoVAT = SAPbobsCOM.BoYesNoEnum.tNO

            For Each Par In Asientos

                objGlobal.SBOApp.StatusBar.SetText("Creando detalle de líneas de asiento - " & i & " de " & Asientos.Count, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                sGrupo = Par.Key
                dblImpDebe = Par.Value
                'divido
                Dim testArray() As String = sGrupo.Split(CType("|", Char()))
                sProyecto = ""
                sCECO = ""
                SDime4 = ""
                sDime5 = ""
                For j As Integer = 0 To testArray.Length - 1
                    If testArray(j) <> "" Then
                        Select Case j
                            Case 0
                                sProyecto = testArray(j)
                            Case 1
                                sCECO = testArray(j)
                            Case 2
                                SDime4 = testArray(j)
                            Case 3
                                sDime5 = testArray(j)
                        End Select
                    End If
                Next

                If i <> 0 Then
                    oOJDT.Lines.Add()
                End If

                oOJDT.Lines.AccountCode = sCtaDebe
                If dblImpDebe <> 0 Then
                    bolAsiento = True
                End If
                oOJDT.Lines.Debit = dblImpDebe
                oOJDT.Lines.Credit = 0
                oOJDT.Lines.CostingCode = sCECO
                oOJDT.Lines.CostingCode2 = ""
                oOJDT.Lines.CostingCode3 = ""
                oOJDT.Lines.CostingCode4 = SDime4
                oOJDT.Lines.CostingCode5 = sDime5
                oOJDT.Lines.ProjectCode = sProyecto

                i = i + 1
            Next

            'linea haber
            'dblImpHaber = CDbl((dtDatos.Compute("SUM(ImporteProvision)", "")))

            'haber
            For Each Par In Asientos
                sGrupo = Par.Key
                dblImpHaber = Par.Value
                'divido
                Dim testArray() As String = sGrupo.Split(CType("|", Char()))
                sProyecto = ""
                sCECO = ""
                SDime4 = ""
                sDime5 = ""
                For j As Integer = 0 To testArray.Length - 1
                    If testArray(j) <> "" Then
                        Select Case j
                            Case 0
                                sProyecto = testArray(j)
                            Case 1
                                sCECO = testArray(j)
                            Case 2
                                SDime4 = testArray(j)
                            Case 3
                                sDime5 = testArray(j)
                        End Select
                    End If
                Next

                oOJDT.Lines.Add()
                oOJDT.Lines.AccountCode = SCtaHaber
                If dblImpHaber <> 0 Then
                    bolAsiento = True
                End If
                oOJDT.Lines.Credit = dblImpHaber
                oOJDT.Lines.Debit = 0
                oOJDT.Lines.CostingCode = sCECO
                oOJDT.Lines.CostingCode2 = ""
                oOJDT.Lines.CostingCode3 = ""
                oOJDT.Lines.CostingCode4 = SDime4
                oOJDT.Lines.CostingCode5 = sDime5
                oOJDT.Lines.ProjectCode = sProyecto
            Next



            'oOJDT.Lines.Add()
            'oOJDT.Lines.AccountCode = SCtaHaber
            'oOJDT.Lines.Credit = dblImpHaber
            'oOJDT.Lines.Debit = 0
            'oOJDT.Lines.CostingCode = sCECO
            'oOJDT.Lines.CostingCode2 = ""
            'oOJDT.Lines.CostingCode3 = ""
            'oOJDT.Lines.CostingCode4 = SDime4
            'oOJDT.Lines.CostingCode5 = sDime5
            'oOJDT.Lines.ProjectCode = sProyecto
            If bolAsiento = True Then


                If oOJDT.Add() <> 0 Then
                    Throw New Exception(objGlobal.compañia.GetLastErrorCode & " / " & objGlobal.compañia.GetLastErrorDescription)
                End If
                sTransId = objGlobal.compañia.GetNewObjectKey
                'End If
            End If
            'update de la linea de documentos para pasarlos a tratados 
            For Each row As DataRow In dtDatos.Rows
                sSql = "UPDATE INV1 SET  U_EXO_GENPROV ='Y' WHERE DocEntry='" & row("DocEntry").ToString & "' And LineNum ='" & row("LineNum").ToString & "' "
                oRs.DoQuery(sSql)
            Next

            If objGlobal.compañia.InTransaction = True Then
                objGlobal.compañia.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If

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

            If oOJDT IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oOJDT)
            If oRs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRs)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
#End Region
End Class
