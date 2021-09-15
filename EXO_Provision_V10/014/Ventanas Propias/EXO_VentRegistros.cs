using System;
using System.Collections.Generic;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{
    class EXO_VENREG
    {

        public EXO_VENREG()
        { }

        public EXO_VENREG(bool lCreacion,  SAPbouiCOM.DataTable oTabla = null, string Titulo = "")
        {
            
            if (lCreacion)
            {
                SAPbouiCOM.Form oForm = null;
                SAPbouiCOM.Matrix oMatrix = null;
                #region CargoScreen

                SAPbouiCOM.FormCreationParams oParametrosCreacion = (SAPbouiCOM.FormCreationParams)(Matriz.oGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams));
                string strXML = "";
                //string strXML = Utilidades.LeoFichEmbebido("xEXO_VenReg.srf");                
                oParametrosCreacion.XmlData = strXML;
                oParametrosCreacion.FormType = "VENREGFACTRANS";
                                
                try
                {
                    oForm = Matriz.oGlobal.SBOApp.Forms.AddEx(oParametrosCreacion);
                }
                catch (Exception ex)
                {
                    Matriz.oGlobal.SBOApp.MessageBox(Matriz.oGlobal.compañia.GetLastErrorDescription(), 1, "Ok", "", "");
                }

                #endregion
                oForm.Title = Titulo;

                #region binding
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("matLin").Specific;
                oMatrix.Columns.Item("V_1").DataBind.Bind("TablaReg", "Clave");
                oMatrix.Columns.Item("V_0").DataBind.Bind("TablaReg", "Mensaje");
                oMatrix.Columns.Item("V_2").DataBind.Bind("TablaReg", "Objeto");
                ((SAPbouiCOM.LinkedButton)oMatrix.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_Order;
                #endregion

                oMatrix.Columns.Item("V_2").Visible = false;

                //lleno rejilla
                SAPbouiCOM.DataTable oTablaMATRIX = oForm.DataSources.DataTables.Item("TablaReg");
                for (int j = 0; j <= oTabla.Rows.Count - 1; j++)
                {
                    oTablaMATRIX.Rows.Add();
                    oTablaMATRIX.SetValue("Clave", oTablaMATRIX.Rows.Count - 1, oTabla.GetValue("Clave", j));
                    oTablaMATRIX.SetValue("Mensaje", oTablaMATRIX.Rows.Count - 1, oTabla.GetValue("Mensaje", j));
                    oTablaMATRIX.SetValue("Objeto", oTablaMATRIX.Rows.Count - 1, oTabla.GetValue("Objeto", j));
                }
                
                //oTablaMATRIX.CopyFrom(oTabla);
                oMatrix.LoadFromDataSource();                               
                
            }
        }

        public bool ItemEvent(ItemEvent infoEvento)
        {

            switch (infoEvento.EventType)
            {
                case BoEventTypes.et_MATRIX_LINK_PRESSED:
                    {
                        if (infoEvento.ItemUID == "matLin" && infoEvento.BeforeAction)
                        {
                            SAPbouiCOM.Form oForm = Matriz.oGlobal.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);

                            //infoEvento
                            SAPbouiCOM.Matrix oMatLin = (SAPbouiCOM.Matrix) oForm.Items.Item("matLin").Specific;

                            string cObjeto = ((SAPbouiCOM.EditText)oMatLin.GetCellSpecific("V_2", infoEvento.Row)).Value;
                            switch (cObjeto)
                            {
                                //Factura
                                case "13":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_Invoice;                                    
                                    break;
                                //Abono
                                case "14":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_InvoiceCreditMemo;
                                    break;
                                //Albaran de compra
                                case "20":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_GoodsReceiptPO;
                                    break;     
                                //Asiento
                                case "30":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_JournalPosting;
                                    break;
                                //Documentos preliminares
                                case "112":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_Drafts;
                                    break;
                                //Traslados
                                case "67":
                                    ((SAPbouiCOM.LinkedButton)oMatLin.Columns.Item("V_1").ExtendedObject).LinkedObject = BoLinkedObject.lf_StockTransfers;
                                    break;                                                   
                                default:

                                    return false;
                                    //break;
                            }                            
                        }                        
                    }
                    break;

            }
            return true;
        }
    }
}
