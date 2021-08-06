using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;


public enum DH { Debe, Haber }

struct ApuntesProvision
{
    public string Cuenta;
    public DH DebeHaber;
    public double Importe;
    public string CeCo1;
    public string CeCo2;
    public string CeCo3;
    public string Proyecto;
    public int ClavePed;
    public int DocNumPed; 
}

namespace Cliente
{
    public class EXO_Genprov
    {
        public EXO_Genprov(bool lCreacion = false)
        {
            if (lCreacion)
            {
                SAPbouiCOM.Form oForm = null;

                #region CargoScreen
                SAPbouiCOM.FormCreationParams oParametrosCreacion = (SAPbouiCOM.FormCreationParams)(Matriz.oGlobal.conexionSAP.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams));
                
                //EXOGENPROV                
                string strXML = Utilidades.LeoFichEmbebido("Formularios.xEXO_ProviFactTrans.srf");
                oParametrosCreacion.XmlData = strXML;
                oParametrosCreacion.UniqueID = "";

                try
                {
                    oForm = Matriz.oGlobal.conexionSAP.SBOApp.Forms.AddEx(oParametrosCreacion);
                }
                catch (Exception ex)
                {
                    Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
                }
                #endregion

                oForm.DataSources.UserDataSources.Item("dsText").ValueEx = "Provision Art No Inv";
                
                
                SAPbouiCOM.Matrix oMatLin = (SAPbouiCOM.Matrix)oForm.Items.Item("matLin").Specific;
                oMatLin.Columns.Item("Col_1").Width = 15;
              
                oForm.Visible = true;

            }
        }

        public bool ItemEvent(EXO_Generales.EXO_infoItemEvent infoEvento)
        {
            SAPbouiCOM.Form oForm = Matriz.oGlobal.conexionSAP.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);

            switch (infoEvento.EventType)
            {

                case BoEventTypes.et_VALIDATE:
                    #region Refesco la matriz al cambiar de HASTA fecha
                    if (!infoEvento.InnerEvent && infoEvento.BeforeAction && infoEvento.ItemUID == "txtHFecha" && infoEvento.ItemChanged)
                    {
                        string cHastaFecha = oForm.DataSources.UserDataSources.Item("dsHFecha").ValueEx;
                        #region Valido
                        if (cHastaFecha == "")
                        {
                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("Ha de introducir 'Hasta Fecha'", 1, "Ok", "", "");
                            return false;
                        }
                        #endregion

                        CargoMatriz(ref oForm);
                    }
                    #endregion
                    break;


                case BoEventTypes.et_ITEM_PRESSED:
                    if (!infoEvento.BeforeAction && infoEvento.ItemUID == "btnGenerar")
                    {
                        if (Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("¿ Generar asiento de provision ?", 1, "Si", "No", "") != 1) return true;

                        string cFechaAsiento = oForm.DataSources.UserDataSources.Item("dsFecha").ValueEx;
                        #region Valido
                        if (cFechaAsiento == "")
                        {
                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("Fecha Asiento no válida", 1, "Ok", "", "");
                            return true;
                        }
                        #endregion

                        //Avanzado o no
                        bool lAvanzada = ( Matriz.oGlobal.SQL.sqlStringB1("SELECT isnull(NewAcctDe, 'Y') AS 'DeterAva' FROM OADM") == "Y" );

                        
                        string cTextoApun = oForm.DataSources.UserDataSources.Item("dsText").ValueEx;

                        SAPbouiCOM.DataTable oTabla = oForm.DataSources.DataTables.Item("TablaDat");
                        DateTime dFechasiento;
                        DateTime.TryParseExact(cFechaAsiento, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dFechasiento);
                                              
                        string sqlBase = Utilidades.LeoFichEmbebido("qDetaProvi.sql");
                        sqlBase = sqlBase.Replace("##AVANZADA", lAvanzada ? "Y" : "N");

                        List<ApuntesProvision> ListaApuntes = new List<ApuntesProvision>();
                        ApuntesProvision AuxApun;                        
                        for (int j = 0; j < oTabla.Rows.Count; j++)
                        {
                            #region Relleno la lista con los futuros apuntes a crear
                            int nNumPed = oTabla.GetValue("Clave", j);
                            string sql = sqlBase.Replace("##CLAVEPEDIDO", nNumPed.ToString());
                            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.SQL.sqlComoRsB1(sql);
                            while (!oRec.EoF)
                            {
                                AuxApun = new ApuntesProvision();
                                //Apunte gasto
                                AuxApun.Cuenta = oRec.Fields.Item("CuentaGasto").Value;
                                AuxApun.DebeHaber = DH.Debe;
                                AuxApun.Importe = oRec.Fields.Item("TotalProvisionar").Value;
                                AuxApun.ClavePed = oRec.Fields.Item("ClavePed").Value;
                                AuxApun.CeCo1 = oRec.Fields.Item("CeCo1").Value;
                                AuxApun.CeCo2 = oRec.Fields.Item("CeCo2").Value;
                                AuxApun.CeCo3 = oRec.Fields.Item("CeCo3").Value;
                                AuxApun.Proyecto = oRec.Fields.Item("Projecto").Value;
                                AuxApun.DocNumPed = oRec.Fields.Item("DocNum").Value;
                                ListaApuntes.Add(AuxApun);

                                //Apunte provision
                                AuxApun = new ApuntesProvision();
                                AuxApun.Cuenta = oRec.Fields.Item("CuentaProvi").Value;
                                AuxApun.DebeHaber = DH.Haber;
                                AuxApun.Importe = oRec.Fields.Item("TotalProvisionar").Value;
                                AuxApun.ClavePed = oRec.Fields.Item("ClavePed").Value;
                                AuxApun.CeCo1 = "";
                                AuxApun.CeCo2 = "";
                                AuxApun.CeCo3 = "";
                                AuxApun.Proyecto = "";
                                AuxApun.DocNumPed = oRec.Fields.Item("DocNum").Value;
                                ListaApuntes.Add(AuxApun);

                                oRec.MoveNext();
                            }
                            #endregion
                        }


                        if (ListaApuntes.Count == 0)
                        {
                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("No hay apuntes a crear", 1, "Ok", "", "");
                            return true;
                        }
                        
                        SAPbobsCOM.JournalEntries oAsiento = (SAPbobsCOM.JournalEntries) Matriz.oGlobal.conexionSAP.compañia.GetBusinessObject(BoObjectTypes.oJournalEntries);

                        oAsiento.ReferenceDate = dFechasiento;
                        oAsiento.TaxDate = dFechasiento;
                        oAsiento.DueDate = dFechasiento;
                        oAsiento.UserFields.Fields.Item("U_EXO_EsProvi").Value = "Y";
                        oAsiento.Memo = cTextoApun;

                        bool lPrimera = true;
                        List<int> ListaPedidos = new List<int>();

                        foreach (ApuntesProvision AuxApunProv in ListaApuntes)
                        {
                            if (!lPrimera) oAsiento.Lines.Add();

                            oAsiento.Lines.AccountCode = AuxApunProv.Cuenta;
                            if (AuxApunProv.CeCo1 != "") oAsiento.Lines.CostingCode = AuxApunProv.CeCo1;
                            if (AuxApunProv.CeCo2 != "") oAsiento.Lines.CostingCode2 = AuxApunProv.CeCo2;
                            if (AuxApunProv.CeCo3 != "") oAsiento.Lines.CostingCode3 = AuxApunProv.CeCo3;
                            if (AuxApunProv.Proyecto != "") oAsiento.Lines.ProjectCode = AuxApunProv.Proyecto;
                            if (AuxApunProv.DebeHaber == DH.Debe) oAsiento.Lines.Debit = AuxApunProv.Importe;
                            if (AuxApunProv.DebeHaber == DH.Haber) oAsiento.Lines.Credit = AuxApunProv.Importe;
                            oAsiento.Lines.Reference1 = AuxApunProv.DocNumPed.ToString();

                            if (!ListaPedidos.Exists(x => x == AuxApunProv.ClavePed)) ListaPedidos.Add(AuxApunProv.ClavePed);
                            lPrimera = false;
                        }

                        string cMenError = "";
                        string cNuevApun = "";
                        bool lTransaccionOK = false;
                        try
                        {          
                            #region INICIO TRANSACCION
                            if (Matriz.oGlobal.conexionSAP.compañia.InTransaction)
                            {
                                Matriz.oGlobal.conexionSAP.compañia.EndTransaction(BoWfTransOpt.wf_RollBack);
                            }
                            Matriz.oGlobal.conexionSAP.compañia.StartTransaction();
                            #endregion

                            if (oAsiento.Add() == 0)
                            {
                                cNuevApun = Matriz.oGlobal.conexionSAP.compañia.GetNewObjectKey();
                                #region Genero la lista
                                string cCadenaUp = "";
                                foreach (int x in ListaPedidos)
                                {
                                    cCadenaUp += x.ToString() + ",";
                                }
                                cCadenaUp = cCadenaUp.Substring(0, cCadenaUp.Length - 1);
                                #endregion

                                string sqlUp = "UPDATE OPOR SET U_EXO_AsiProvComp = "  + cNuevApun + " WHERE DocEntry IN (" + cCadenaUp + ")";
                                Utilidades.ActualizarBBDD(sqlUp, ref cMenError);

                                if (cMenError == "")
                                {
                                    #region COMPLETO TRANSACCION
                                    if (Matriz.oGlobal.conexionSAP.compañia.InTransaction)
                                    {
                                        Matriz.oGlobal.conexionSAP.compañia.EndTransaction(BoWfTransOpt.wf_Commit);
                                        lTransaccionOK = true;
                                    }
                                    #endregion
                                }
                            }
                            else
                            {
                                cMenError = Matriz.oGlobal.conexionSAP.compañia.GetLastErrorDescription();
                            }
                        }
                        catch(Exception ex)
                        {
                            #region RECHAZO TRANSACCION
                            if (Matriz.oGlobal.conexionSAP.compañia.InTransaction)
                            {
                                Matriz.oGlobal.conexionSAP.compañia.EndTransaction(BoWfTransOpt.wf_RollBack );
                            }
                            #endregion

                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("ERROR en la generacion del apunte de provision\n" + ex.Message, 1, "Ok", "", "");                                                        
                        }
                        finally
                        {
                            #region RECHAZO TRANSACCION
                            if (Matriz.oGlobal.conexionSAP.compañia.InTransaction)
                            {
                                Matriz.oGlobal.conexionSAP.compañia.EndTransaction(BoWfTransOpt.wf_RollBack );
                                if (cMenError == "") cMenError = "No se completo la transaccion";
                                                                                                       
                                Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("No se completo la transaccion\n" + cMenError, 1, "Ok", "", "");                            
                            }
                            #endregion                           
                        }

                        if (lTransaccionOK)
                        {
                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("Proceso terminado con exito\nCreado apunte " + cNuevApun, 1, "Ok", "", "");
                            CargoMatriz(ref oForm);
                        }                        
                    }
                    break;


            }

            return true;
        }
       
        private static void CargoMatriz(ref SAPbouiCOM.Form oForm)
        {
            string cHastaFecha = oForm.DataSources.UserDataSources.Item("dsHFecha").ValueEx;

            string sql = Utilidades.LeoFichEmbebido("qProviCompra.sql");
            sql = sql.Replace("##HASTAFECHA", cHastaFecha);
            SAPbouiCOM.DataTable oTabla = oForm.DataSources.DataTables.Item("TablaDat");
            oTabla.ExecuteQuery(sql);

            SAPbouiCOM.Matrix oMatLin = (SAPbouiCOM.Matrix)oForm.Items.Item("matLin").Specific;

            #region Bindeo - Lo hace mal sap studio - Aunque aqui no hacia falta
            oMatLin.Columns.Item("Col_1").DataBind.Bind("TablaDat", "Clave");
            oMatLin.Columns.Item("Col_3").DataBind.Bind("TablaDat", "NumDoc");
            oMatLin.Columns.Item("Col_4").DataBind.Bind("TablaDat", "Fecha Doc");
            oMatLin.Columns.Item("Col_2").DataBind.Bind("TablaDat", "Ref Prov");
            oMatLin.Columns.Item("Col_5").DataBind.Bind("TablaDat", "Proveedor");
            oMatLin.Columns.Item("Col_7").DataBind.Bind("TablaDat", "Nombre");
            oMatLin.Columns.Item("Col_0").DataBind.Bind("TablaDat", "Total Doc");
            oMatLin.Columns.Item("Col_6").DataBind.Bind("TablaDat", "Total Provisionar");
            #endregion            
            oForm.Freeze(true);
            try
            {                
                oMatLin.LoadFromDataSource();
                EXO_CleanCOM.CLiberaCOM.FormMatrix(ref oMatLin);
            }
            catch (Exception ex)
            { }
            oForm.Freeze(false);                       
        }
    }
}

