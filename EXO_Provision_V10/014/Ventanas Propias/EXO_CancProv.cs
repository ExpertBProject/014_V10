using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{
    public class EXO_CancProv
    {
        public EXO_CancProv(bool lCreacion = false, string cFormUID = "",  int nTransId = 0)
        {
            if (lCreacion)
            {
                SAPbouiCOM.Form oForm = null;

                #region CargoScreen
                SAPbouiCOM.FormCreationParams oParametrosCreacion = (SAPbouiCOM.FormCreationParams)(Matriz.oGlobal.conexionSAP.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams));
                
                //EXOCANCPPROV                
                string strXML = Utilidades.LeoFichEmbebido("Formularios.xEXO_CanceProv.srf");
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

                oForm.DataSources.UserDataSources.Item("dsFecCan").ValueEx = DateTime.Now.ToString("yyyyMMdd");
                oForm.DataSources.UserDataSources.Item("dsAsien").ValueEx = nTransId.ToString();
                oForm.DataSources.UserDataSources.Item("dsFormU").ValueEx = cFormUID;

                oForm.Visible = true;
            }           
       }

        public bool ItemEvent(EXO_Generales.EXO_infoItemEvent infoEvento)
        {
            SAPbouiCOM.Form oForm = Matriz.oGlobal.conexionSAP.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);

            switch (infoEvento.EventType)
            {

                case BoEventTypes.et_ITEM_PRESSED:
                    if (infoEvento.ItemUID == "btnOk" && !infoEvento.BeforeAction)
                    {                        

                        int nNumAsi = Convert.ToInt32(oForm.DataSources.UserDataSources.Item("dsAsien").ValueEx);
                        #region Pido confirmacion
                        if (Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("¿ Cancelar asiento " + nNumAsi.ToString() + " ?", 2, "Si", "No", "") != 1)
                        {
                            return true;
                        }

                        string cFechaAsiento = oForm.DataSources.UserDataSources.Item("dsFecCan").ValueEx;
                        if (cFechaAsiento == "")
                        {
                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("Ha de introducir Fecha para el asiento de cancelacion", 1, "Ok", "", "");
                            return true;
                        }

                        DateTime dFechasiento;
                        DateTime.TryParseExact(cFechaAsiento, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dFechasiento);


                        #endregion

                        bool lTransaccionOK = false;
                        string cNuevoAsiento = "";
                        string cMenError = "";

                        try
                        {
                            
                            #region INICIO TRANSACCION
                            if (Matriz.oGlobal.conexionSAP.compañia.InTransaction)
                            {
                                Matriz.oGlobal.conexionSAP.compañia.EndTransaction(BoWfTransOpt.wf_RollBack);
                            }
                            Matriz.oGlobal.conexionSAP.compañia.StartTransaction();
                            #endregion

                            #region Creo el asiento 'al reves' y 'desmarco' los pedidos
                            SAPbobsCOM.JournalEntries oAsiento = (SAPbobsCOM.JournalEntries)Matriz.oGlobal.conexionSAP.compañia.GetBusinessObject(BoObjectTypes.oJournalEntries);
                            oAsiento.ReferenceDate = dFechasiento;
                            oAsiento.TaxDate = dFechasiento;
                            oAsiento.DueDate = dFechasiento;
                            oAsiento.UserFields.Fields.Item("U_EXO_CancProv").Value = "Y";
                            oAsiento.UserFields.Fields.Item("U_EXO_EsProvi").Value = "Y";   
                            //oAsiento.UserFields.Fields.Item("U_EXO_EsProvi").Value = "Y";

                            bool lPrimera = true;
                            

                            string sql = "SELECT T0.TransId, isnull(T0.Memo, '') as 'Memo', T1.Ref1 AS 'Ref1', T1.Account as 'Cuenta', T1.Credit as 'Haber', T1.Debit as 'Debe', T1.Project as 'Proyecto', ";
                            sql += " T1.ProfitCode as 'CeCo1', T1.OcrCode2 as 'Ceco2', T1.OcrCode3 as 'Ceco3', T1.Line_ID FROM OJDT T0 INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId ";
                            sql += " WHERE T0.TransId = " + nNumAsi.ToString();
                            sql += " ORDER BY T1.Line_ID ";
                            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.SQL.sqlComoRsB1(sql);
                            while (!oRec.EoF)
                            {
                                if (lPrimera)
                                {
                                    string cTexto = "CANCELACION - " + nNumAsi.ToString() + " " + oRec.Fields.Item("Memo").Value;
                                    oAsiento.Memo = cTexto.Substring(0, Math.Min(cTexto.Length, 49));
                                    oAsiento.Reference = nNumAsi.ToString();
                                }
                                else
                                {
                                    oAsiento.Lines.Add();
                                }
                                oAsiento.Lines.AccountCode = oRec.Fields.Item("Cuenta").Value;
                                oAsiento.Lines.Reference1 = oRec.Fields.Item("Ref1").Value;

                                if (cMenError == "")
                                {
                                    #region Norma 1
                                    if (oRec.Fields.Item("CeCo1").Value != "")
                                    {
                                        double nImporte = (oRec.Fields.Item("Debe").Value == 0) ? oRec.Fields.Item("Haber").Value : oRec.Fields.Item("Debe").Value;
                                        string cNorma = AsignoNorma(1, oRec.Fields.Item("CeCo1").Value, ref cMenError);
                                        if (cMenError == "")
                                        {
                                            oAsiento.Lines.CostingCode =  cNorma;
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                    #endregion
                                }

                                if (cMenError == "")
                                {
                                    #region Norma 2
                                    if (oRec.Fields.Item("CeCo2").Value != "")
                                    {
                                        double nImporte = (oRec.Fields.Item("Debe").Value == 0) ? oRec.Fields.Item("Haber").Value : oRec.Fields.Item("Debe").Value;
                                        string cNorma = AsignoNorma(2, oRec.Fields.Item("CeCo2").Value, ref cMenError);
                                        if (cMenError == "")
                                        {
                                            oAsiento.Lines.CostingCode2 = cNorma;
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                    #endregion
                                }

                                if (cMenError == "")
                                {
                                    #region Norma 3
                                    if (oRec.Fields.Item("CeCo3").Value != "")
                                    {
                                        double nImporte = (oRec.Fields.Item("Debe").Value == 0) ? oRec.Fields.Item("Haber").Value : oRec.Fields.Item("Debe").Value;
                                        string cNorma = AsignoNorma(3, oRec.Fields.Item("CeCo3").Value, ref cMenError);
                                        if (cMenError == "")
                                        {
                                            oAsiento.Lines.CostingCode3 = cNorma;
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                    #endregion
                                }

                                if (cMenError != "") break;

                                if (oRec.Fields.Item("Proyecto").Value != "") oAsiento.Lines.ProjectCode = oRec.Fields.Item("Proyecto").Value;
                                    
                                if (oRec.Fields.Item("Debe").Value != 0) oAsiento.Lines.Debit = -1 * oRec.Fields.Item("Debe").Value;
                                if (oRec.Fields.Item("Haber").Value != 0) oAsiento.Lines.Credit = -1 * oRec.Fields.Item("Haber").Value;
                            
                                lPrimera = false;                            
                                oRec.MoveNext();                                
                            }

                            if (cMenError == "")
                            {                                
                                if (oAsiento.Add() == 0)
                                {
                                    string sqlUp = "UPDATE OPOR SET U_EXO_AsiProvComp = 0 WHERE U_EXO_AsiProvComp = " + nNumAsi.ToString();
                                    Utilidades.ActualizarBBDD(sqlUp, ref cMenError);
                                    if (cMenError == "")
                                    {

                                        sqlUp = "UPDATE OJDT SET U_EXO_CancProv = 'Y' WHERE TransId = " + nNumAsi.ToString();
                                        Utilidades.ActualizarBBDD(sqlUp, ref cMenError);
                                        if (cMenError == "")
                                        {

                                            #region COMPLETO TRANSACCION - UNICO PUNTO
                                            if (Matriz.oGlobal.conexionSAP.compañia.InTransaction)
                                            {
                                                Matriz.oGlobal.conexionSAP.compañia.EndTransaction(BoWfTransOpt.wf_Commit);
                                                lTransaccionOK = true;
                                                cNuevoAsiento = Matriz.oGlobal.conexionSAP.compañia.GetNewObjectKey();
                                            }
                                            #endregion
                                        }
                                    }
                                }
                                else
                                {
                                    cMenError = Matriz.oGlobal.conexionSAP.compañia.GetLastErrorDescription();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            #region RECHAZO TRANSACCION
                            if (Matriz.oGlobal.conexionSAP.compañia.InTransaction)
                            {
                                Matriz.oGlobal.conexionSAP.compañia.EndTransaction(BoWfTransOpt.wf_RollBack);
                            }
                            #endregion

                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("ERROR en la generacion del apunte de provision\n" + ex.Message, 1, "Ok", "", "");                                                        
          
                        }
                        finally
                        {
                            #region RECHAZO TRANSACCION
                            if (Matriz.oGlobal.conexionSAP.compañia.InTransaction)
                            {
                                Matriz.oGlobal.conexionSAP.compañia.EndTransaction(BoWfTransOpt.wf_RollBack);
                                if (cMenError == "") cMenError = "No se completo la transaccion";                                
                            }
                            #endregion                           
                        }

                        if (lTransaccionOK)
                        {
                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox("Proceso terminado con exito\nAsiento creado " + cNuevoAsiento, 1, "Ok", "", "");
                            string cUnqFormIDOQUT = oForm.DataSources.UserDataSources.Item("dsFormU").ValueEx;
                            oForm.Close();

                            if (Utilidades.BuscoFormLanzado(cUnqFormIDOQUT) != null)
                            {
                                Utilidades.Reposiciono(cUnqFormIDOQUT, Convert.ToInt32(cNuevoAsiento));
                            }
                            return true;
                        }
                        else
                        {
                            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(cMenError, 1, "Ok", "", "");
                        }

                        #endregion
                    }
                    break;
            }

            return true;
        }

        public string AsignoNorma(int nDimension, string cNormaReparto, ref string cMenError) 
        {
            string cRetorno = "";

            try
            {
                if (Matriz.oGlobal.SQL.sqlStringB1("SELECT T0.OcrCode FROM OOCR T0 WHERE T0.OcrCode = '" + cNormaReparto + "'") != "")
                {
                    cRetorno = cNormaReparto;
                }
                else
                {
                    List<RepartosCeCos> ListaCeCosDim = new List<RepartosCeCos>();

                    string sql1 = "SELECT T0.OcrCode as 'CodNorma', T0.PrcCode AS 'CeCo', T0.PrcAmount as 'Total' FROM MDR1 T0 WHERE T0.OcrCode = '" + cNormaReparto + "'";
                    SAPbobsCOM.Recordset oRecNorma = Matriz.oGlobal.SQL.sqlComoRsB1(sql1);
                    while (!oRecNorma.EoF)
                    {
                        double nImpCeCo = oRecNorma.Fields.Item("Total").Value;
                        //if (Math.Abs(nImpCeCo) >= Matriz.ngValorMinimo)
                        //{
                        RepartosCeCos AuxCeCo;
                        AuxCeCo.CeCo = oRecNorma.Fields.Item("CeCo").Value;
                        AuxCeCo.Importe = -1 * (decimal)nImpCeCo;
                        ListaCeCosDim.Add(AuxCeCo);
                        //}
                        oRecNorma.MoveNext();
                    }
                    Object Ob = (Object)oRecNorma;
                    EXO_CleanCOM.CLiberaCOM.liberaCOM(ref Ob);

                    cRetorno = Utilidades.CreoNormaRepartoManual(ref cMenError, 1, ListaCeCosDim);
                    if (cMenError != "")
                    {
                        cRetorno = "";
                    }
                }
            }
            catch (Exception ex)
            {
                cMenError = ex.Message;
            }

            return cRetorno;                  
        }
       
    }
}
