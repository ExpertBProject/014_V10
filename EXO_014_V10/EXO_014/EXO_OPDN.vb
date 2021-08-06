Imports SAPbouiCOM

Public Class EXO_OPDN

    'Private refDI As EXO_DIAPI.EXO_DIAPI
    'Private refUI As EXO_UIAPI.EXO_UIAPI
    Private objGlobal As EXO_UIAPI.EXO_UIAPI

    Public Sub New(objGlobal As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objGlobal
        'refDI = objGlobal.conexionSAP.refCompañia
        'refUI = objGlobal.conexionSAP.refSBOApp

    End Sub

    Public Function SBOApp_FormDataEvent(ByVal infoEvento As BusinessObjectInfo) As Boolean

        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(infoEvento.FormUID)
        Try

            If infoEvento.EventType = BoEventTypes.et_FORM_DATA_ADD Or infoEvento.EventType = BoEventTypes.et_FORM_DATA_UPDATE Then
                If objGlobal.SBOApp.MessageBox("¿Desea generar la entrega del pedido de venta asociado a la compra?", 1, "Aceptar", "Cancelar") = 1 Then
                    'buscamos el pedido enlazado al documento
                    GenerarEntrega(infoEvento)
                End If
            End If

        Catch ex As Exception
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
        End Try

        Return True

    End Function

    Private Sub GenerarEntrega(infoEvento As BusinessObjectInfo)

        Dim sql As String = ""
        Dim oXml As New Xml.XmlDocument
        Dim sDocEntry As String = ""
        Dim tablaCabecera As System.Data.DataTable = Nothing
        Dim tablaLineas As System.Data.DataTable = Nothing

        oXml.LoadXml(infoEvento.ObjectKey)
        sDocEntry = oXml.SelectSingleNode("DocumentParams/DocEntry").InnerText

        sql = " select t3.CardCode,t4.DocDueDate,t4.U_EXO_FENT" +
            " from  pdn1 t0 INNER join por1 t1 on t0.BaseEntry=t1.DocEntry and t0.BaseLine=t1.linenum and t0.BaseType=t1.ObjType " +
            " INNER join ordr t3 on  t1.BaseEntry=t3.DocEntry " +
            " INNER JOIN RDR1 T2 ON T2.DocEntry=T3.DocEntry " +
            " inner join opdn t4 on t0.docentry=t4.docentry " +
            " where t2.LineStatus='O' and t0.DocEntry=" + sDocEntry + ""

        tablaCabecera = objGlobal.refDi.SQL.sqlComoDataTable(sql)

        If tablaCabecera.Rows.Count > 0 Then

            Dim oODLN As SAPbobsCOM.Documents = Nothing
            Dim bPrimeraLinea As Boolean = True
            Dim bEsPrimero As Boolean = True
            Dim tablaLotesSeries As System.Data.DataTable = Nothing

            oODLN = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes), SAPbobsCOM.Documents)

            'necesitamos el cliente
            'la fecha de entrega del pedido

            oODLN.CardCode = tablaCabecera.Rows(0).Item("CardCode").ToString()

            If tablaCabecera.Rows(0).Item("U_EXO_FENT").ToString() = "" Then
                oODLN.DocDueDate = tablaCabecera.Rows(0).Item("DocDueDate").ToString()
                oODLN.DocDate = tablaCabecera.Rows(0).Item("DocDueDate").ToString()
                oODLN.TaxDate = tablaCabecera.Rows(0).Item("DocDueDate").ToString()
            Else
                oODLN.DocDueDate = tablaCabecera.Rows(0).Item("U_EXO_FENT").ToString()
                oODLN.DocDate = tablaCabecera.Rows(0).Item("U_EXO_FENT").ToString()
                oODLN.TaxDate = tablaCabecera.Rows(0).Item("U_EXO_FENT").ToString()
            End If


            'oODLN.DocDate = Date.Now()

            '''VERSION SIN PROYECTO
            sql = " select t0.ItemCode,case when t0.Quantity > t2.OpeNQty then t2.OpeNQty else t0.quantity end as Quantity,t2.docentry,t2.linenum,17 as 'objtype', t0.linenum as 'LineaEntrada'," +
                " CASE WHEN T4.InvntItem = 'N' THEN 'NoInv' when t4.ManBtchNum='Y' then 'Lote' when t4.ManSerNum='Y' then 'Serie' else 'Normal' end as 'TipoArt' " +
                " from pdn1 t0 inner join por1 t1 on t0.BaseEntry=t1.DocEntry and t0.BaseLine=t1.linenum and t0.BaseType=t1.ObjType " +
                " inner join rdr1 t2 on t1.BaseEntry=t2.DocEntry and t1.baseline=t2.linenum and t1.basetype=t2.ObjType " +
                " inner join ordr t3 on t2.DocEntry=t3.DocEntry " +
                " inner join OITM t4 on t2.itemcode=t4.itemcode " +
                " where t2.LineStatus='O' and t4.PrchseItem='Y' and t0.DocEntry='" + sDocEntry + "' " +
                " UNION ALL " +
                " select t3.ItemCode,T3.OpeNQty as quantity,t2.docentry,t3.linenum,17 as 'objtype',0 as 'LineaEntrada' , " +
                " CASE WHEN T4.InvntItem = 'N' THEN 'NoInv' when t4.ManBtchNum='Y' then 'Lote' when t4.ManSerNum='Y' then 'Serie' else 'Normal' end as 'TipoArt'  " +
                " from ordr t2 " +
                " inner join rdr1 t3 on t2.docentry=t3.docentry " +
                " inner join oitm t4 on t4.itemcode=t3.itemcode  and t4.PrchseItem='N' and t4.InvntItem='N' " +
                " where  t2.docentry in (select t1.baseentry from pdn1 t0 inner join por1 t1 on t0.BaseEntry=t1.DocEntry and t0.BaseLine=t1.linenum and t0.BaseType=t1.ObjType where t0.docentry='" + sDocEntry + "' ) " +
                " order by linenum  "

            ''''VERSION CON PROYECTO EN LINEA DE PEDIDO Y PROYECTO EN CABECERA EN ENTRADA DE MERCANCIAS
            'sql = " select t0.ItemCode,t0.Quantity,t2.docentry,t2.linenum,17 as 'objtype', t0.linenum as 'LineaEntrada'," +
            '   " CASE WHEN T4.InvntItem = 'N' THEN 'NoInv' when t4.ManBtchNum='Y' then 'Lote' when t4.ManSerNum='Y' then 'Serie' else 'Normal' end as 'TipoArt' " +
            '   " from opdn t00 inner join pdn1 t0 on t00.docentry=t0.docentry  inner join por1 t1 on t0.BaseEntry=t1.DocEntry and t0.BaseLine=t1.linenum and t0.BaseType=t1.ObjType " +
            '   " inner join rdr1 t2 on t1.BaseEntry=t2.DocEntry and t1.baseline=t2.linenum and t1.basetype=t2.ObjType and t00.Project=t2.Project" +
            '   " inner join ordr t3 on t2.DocEntry=t3.DocEntry " +
            '   " inner join OITM t4 on t2.itemcode=t4.itemcode " +
            '   " where t2.LineStatus='O' and t0.DocEntry='" + sDocEntry + "' " +
            '   " UNION ALL " +
            '   " select t3.ItemCode,T3.OpeNQty as quantity,t2.docentry,t3.linenum,17 as 'objtype',0 as 'LineaEntrada' , " +
            '   " CASE WHEN T4.InvntItem = 'N' THEN 'NoInv' when t4.ManBtchNum='Y' then 'Lote' when t4.ManSerNum='Y' then 'Serie' else 'Normal' end as 'TipoArt'  " +
            '   " from ordr t2 " +
            '   " inner join rdr1 t3 on t2.docentry=t3.docentry " +
            '   " inner join oitm t4 on t4.itemcode=t3.itemcode and t4.InvntItem='N' " +
            '   " where  t2.docentry in (select t1.baseentry from pdn1 t0 inner join por1 t1 on t0.BaseEntry=t1.DocEntry and t0.BaseLine=t1.linenum and t0.BaseType=t1.ObjType where t0.docentry='" + sDocEntry + "' ) " +
            '   " order by linenum  "


            tablaLineas = objGlobal.refDi.SQL.sqlComoDataTable(sql)

            If tablaLineas.Rows.Count > 0 Then
                For Each row As DataRow In tablaLineas.Rows

                    If Convert.ToDouble(row.Item("Quantity").ToString()) > 0 Then

                        If Not bPrimeraLinea Then
                            oODLN.Lines.Add()
                        End If

                        bPrimeraLinea = False

                        oODLN.Lines.ItemCode = row.Item("ItemCode").ToString()
                        oODLN.Lines.Quantity = Convert.ToDouble(row.Item("Quantity").ToString())
                        oODLN.Lines.BaseEntry = Convert.ToInt32(row.Item("DocEntry").ToString())
                        oODLN.Lines.BaseLine = Convert.ToInt32(row.Item("linenum").ToString())
                        oODLN.Lines.BaseType = Convert.ToInt32(row.Item("objtype").ToString())

                        tablaLotesSeries = Nothing
                        bEsPrimero = True

                        'recepcionamos la cantidad del pedido, esto nos sirve para controlar los bucles de los lotes y las series y solo coger los necesarios de la recepcion.
                        'Dim cantidadTotal As Double = Convert.ToDouble(row.Item("Quantity").ToString())
                        Dim cantidadRestante As Double = Convert.ToDouble(row.Item("Quantity").ToString())
                        'Dim contcantidad As Double = 0

                        If row.Item("TipoArt").ToString() = "Lote" Then

                            sql = "select t4.DistNumber,t3.Quantity from opdn t0 " +
                                " inner join pdn1 t1 on t0.docentry=t1.docentry " +
                                " inner join oitl t2 on t1.docentry=t2.DocEntry and t2.DocLine=t1.LineNum and t2.DocType=t1.ObjType " +
                                " inner join itl1 t3 on t2.LogEntry=t3.LogEntry " +
                                " inner join obtn t4 on t3.ItemCode=t4.ItemCode and t3.SysNumber=t4.SysNumber " +
                                " where t1.docentry = '" + sDocEntry + "' and t1.linenum ='" + row.Item("LineaEntrada").ToString() + "'"
                            tablaLotesSeries = objGlobal.refDi.SQL.sqlComoDataTable(sql)

                            For Each rowLotesSeries As DataRow In tablaLotesSeries.Rows

                                If cantidadRestante = 0 Then
                                    Exit For
                                End If

                                If Not bEsPrimero Then
                                    oODLN.Lines.BatchNumbers.Add()
                                End If

                                bEsPrimero = False

                                oODLN.Lines.BatchNumbers.BatchNumber = rowLotesSeries.Item("Distnumber").ToString()

                                If Convert.ToDouble(rowLotesSeries.Item("Quantity").ToString()) > cantidadRestante Then

                                    oODLN.Lines.BatchNumbers.Quantity = cantidadRestante
                                    cantidadRestante = 0

                                Else

                                    oODLN.Lines.BatchNumbers.Quantity = Convert.ToDouble(rowLotesSeries.Item("Quantity").ToString())
                                    cantidadRestante = cantidadRestante - Convert.ToDouble(rowLotesSeries.Item("Quantity").ToString())

                                End If

                            Next
                        End If

                        If row.Item("TipoArt").ToString() = "Serie" Then

                            sql = "select t4.MnfSerial,t3.Quantity,t4.sysnumber from opdn t0 " +
                                " inner join pdn1 t1 on t0.docentry=t1.docentry " +
                                " inner join oitl t2 on t1.docentry=t2.DocEntry and t2.DocLine=t1.LineNum and t2.DocType=t1.ObjType " +
                                " inner join itl1 t3 on t2.LogEntry=t3.LogEntry " +
                                " inner join osrn t4 on t3.ItemCode=t4.ItemCode and t3.SysNumber=t4.SysNumber " +
                                " where t1.docentry = '" + sDocEntry + "' and t1.linenum ='" + row.Item("LineaEntrada").ToString() + "'"
                            tablaLotesSeries = objGlobal.refDi.SQL.sqlComoDataTable(sql)

                            For Each rowLotesSeries As DataRow In tablaLotesSeries.Rows

                                If cantidadRestante = 0 Then
                                    Exit For
                                End If

                                If Not bEsPrimero Then
                                    oODLN.Lines.SerialNumbers.Add()
                                End If

                                bEsPrimero = False

                                oODLN.Lines.SerialNumbers.SystemSerialNumber = Convert.ToInt16(rowLotesSeries.Item("sysnumber").ToString())
                                oODLN.Lines.SerialNumbers.ManufacturerSerialNumber = rowLotesSeries.Item("MnfSerial").ToString()
                                oODLN.Lines.SerialNumbers.Quantity = Convert.ToDouble(rowLotesSeries.Item("Quantity").ToString())

                                cantidadRestante = cantidadRestante - 1
                            Next

                        End If
                    End If
                Next

                'busco el valor de la ultima entrega, si al crearla y no da error es el mismo entonces estoy en un proceso de autorizacion
                'debo de decir que se ha creado pero no mostrarlo.

                If oODLN.Add <> 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("Error al generar la entrega: " + objGlobal.compañia.GetLastErrorCode.ToString + "/" + objGlobal.compañia.GetLastErrorDescription & ".", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Else

                    If objGlobal.SBOApp.MessageBox("Documento generado correctamente ¿Desea abrir la entrega?", 1, "Aceptar", "Cancelar") = 1 Then

                        objGlobal.SBOApp.ActivateMenuItem("2051")
                        objGlobal.SBOApp.ActivateMenuItem("1289")
                        'objGlobal.conexionSAP.SBOApp.Forms.ActiveForm.Freeze(False)

                    End If

                    objGlobal.SBOApp.StatusBar.SetText("Entrega realizada correctamente", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                End If



            Else
                'mensaje El documento no esta trazado
            End If
        End If

        'EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(formulario, Object))
        'EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCompras, Object))
    End Sub

End Class

