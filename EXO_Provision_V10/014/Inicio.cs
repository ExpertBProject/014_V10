using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Xml;

namespace Cliente
{
    
    public class Matriz:EXO_UIAPI.EXO_DLLBase
    {
        public static EXO_UIAPI.EXO_UIAPI oGlobal;
        public static Type TypeMatriz;
        public static bool lgProvArtNoInv;
      //  public static double ngValorMinimo;


        public static int SumDec;
        public static int PriceDec;
        public static int RateDec;
        public static int QtyDec;
        public static int PercentDec;
        public static int MeasureDec;
        public static string SepMill;
        public static string SepDec;


        public Matriz(EXO_UIAPI.EXO_UIAPI gen, Boolean act, Boolean usalicencia, int idAddon)
             : base(gen, act, usalicencia, idAddon)
        {
            oGlobal = this.objGlobal;
            TypeMatriz = this.GetType();

            if (act)
            {
                #region Creo tablas y UDO
                if (objGlobal.refDi.comunes.esAdministrador())
                {
                    string fBD = "", cMen = "", cUDO = "";
                    EXO_Generales.EXO_UDO fUDO = null;

                                                          
                    //#region  UDO EXO_LOGINSTALA
                    //cUDO = Utilidades.LeoFichEmbebido("xUDO_EXO_LOGINSTALA.xml");
                    //fUDO = new EXO_Generales.EXO_UDO("UDO_EXO_LOGINSTALA", ref Matriz.oGlobal);
                    //fUDO.validaObjeto(cUDO);
                    //this.SboApp.SetStatusBarMessage("Validado UDO EXO_LOGINSTALA", BoMessageTime.bmt_Short, false);
                    //#endregion
                  
                    
                    #region Campos de usuario y tablas no UDO
                    cMen = "";
                    fBD = Utilidades.LeoFichEmbebido("db_014Provi.xml");                                        
                    if (!objGlobal.refDi.comunes.LoadBDFromXML( fBD, cMen))
                    {
                        objGlobal.SBOApp.MessageBox(cMen, 1, "Ok", "", "");
                        objGlobal.SBOApp.MessageBox("Error en creacion de campos db_014Provi.xml", 1, "Ok", "", "");
                    }
                    else
                    {
                        objGlobal.SBOApp.MessageBox("Actualizacion de campos db_014Provi.xml realizada", 1, "Ok", "", "");
                    }
                    #endregion
                                       
                }
                else
                {
                    objGlobal.SBOApp.MessageBox("Necesita permisos de administrador para actualizar la base de datos.\nCampos no creados", 1, "Ok", "", "");
                }
                #endregion
               
                //cFormSRF = Utilidades.LeoFichEmbebido("EXO_Instaladores.srf");
                //lError = false;
                //cMen1 = Utilidades.FormUDO(cFormSRF, "EXO_INSTALADORES", ref lError);
                //this.SboApp.SetStatusBarMessage(cMen1, BoMessageTime.bmt_Short, lError);          
            }
                     
            #region Decimales de la aplicacion y provisionar o no art no inve
            string sql = "SELECT T0.SumDec as 'SumDec', T0.PriceDec as 'PriceDec', T0.RateDec as 'RateDec', T0.QtyDec as 'QtyDec', T0.PercentDec as 'PercentDec', T0.MeasureDec as 'MeasureDec', ";
            sql += " T0.ThousSep as 'ThousSep', T0.DecSep as 'DecSep', T0.U_EXO_ProviComp as 'ProviComp' FROM OADM T0";
            SAPbobsCOM.Recordset oRec = oGlobal.refDi.SQL.sqlComoRsB1(sql);
            SumDec =  Convert.ToInt32(oRec.Fields.Item("SumDec").Value);
            PriceDec = Convert.ToInt32(oRec.Fields.Item("PriceDec").Value);
            RateDec = Convert.ToInt32(oRec.Fields.Item("RateDec").Value);
            QtyDec = Convert.ToInt32(oRec.Fields.Item("QtyDec").Value);
            PercentDec = Convert.ToInt32(oRec.Fields.Item("PercentDec").Value);
            MeasureDec = Convert.ToInt32(oRec.Fields.Item("MeasureDec").Value);
            SepMill = Convert.ToString(oRec.Fields.Item("ThousSep").Value);
            SepDec = Convert.ToString(oRec.Fields.Item("DecSep").Value);
            lgProvArtNoInv = (Convert.ToString(oRec.Fields.Item("ProviComp").Value) == "Y");
            
 
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            oRec = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            #endregion

            #region Variables generales
            //EXO_DIAPI.EXO_OGEN fGen = new EXO_DIAPI.EXO_OGEN(Matriz.oGlobal.conexionSAP.refCompañia, true, System.Reflection.ProcessorArchitecture.X86);
            //gAlmacenObsoleto = fGen.valorVariable("ALMAOBSOLETO");
            //gAlmacenTransito = fGen.valorVariable("ALMATRANSITO");              
            //fGen = null;
            #endregion           

        }

        public override SAPbouiCOM.EventFilters filtros()
        {
            SAPbouiCOM.EventFilters oFilter = new SAPbouiCOM.EventFilters();
            
            #region Mando filtros
            try
            {                
                string fXML = Utilidades.LeoFichEmbebido("xFiltros014.xml");                                        
                oFilter.LoadFromXML(fXML);
            }
            catch (Exception ex)
            {
                objGlobal.SBOApp.MessageBox("Error en carga de filtros 014 Provision", 1, "Ok", "", "");
                oFilter = null;
            }
            #endregion            
            return oFilter;
        }

        public override XmlDocument menus()
        {

            XmlDocument oXML = new XmlDocument();
            Type MyType = this.GetType();
            string mXML = Utilidades.LeoFichEmbebido(lgProvArtNoInv ? "xMenu014.xml" : "xMenu014Sin.xml");                        
            oXML.LoadXml(mXML);
            return oXML;
        }

        public override bool SBOApp_ItemEvent(ItemEvent infoEvento)
        {
            bool lRetorno = true;


            if (infoEvento.FormTypeEx == "138")
            {
                EXO_138 f138 = new EXO_138();
                lRetorno = f138.ItemEvent(infoEvento);
                f138 = null;
            }

            if (lgProvArtNoInv && infoEvento.FormTypeEx == "142")
            {
                EXO_142 f142 = new EXO_142();
                lRetorno = f142.ItemEvent(infoEvento);
                f142 = null;
            }

            if (lgProvArtNoInv && infoEvento.FormTypeEx == "EXOGENPROV")
            {
                EXO_Genprov fGenProv = new EXO_Genprov();
                lRetorno = fGenProv.ItemEvent(infoEvento);
                fGenProv = null;
            }

            if (lgProvArtNoInv &&  infoEvento.FormTypeEx == "EXOCANCPPROV")
            {
                EXO_CancProv fCancProv = new EXO_CancProv();
                lRetorno = fCancProv.ItemEvent(infoEvento);
                fCancProv = null;
            }

             return lRetorno;            
        }

        public override bool SBOApp_FormDataEvent(BusinessObjectInfo infoDataEvent)
        {
            bool lRetorno = true;

            //Pedido de compras
            if (lgProvArtNoInv && infoDataEvent.FormTypeEx == "142")
            {
                EXO_142 f142 = new EXO_142();
                lRetorno = f142.DataEvent(infoDataEvent);
                f142 = null;
            }


            return lRetorno;
        }

        public override bool SBOApp_MenuEvent(MenuEvent infoMenuEvent)
        {
            bool lRetorno = true;
            
            switch (infoMenuEvent.MenuUID)
            {
                case "1284": //cANCELAR
                    switch (Matriz.oGlobal.SBOApp.Forms.ActiveForm.TypeEx)
                        {

                            case "392":
                                if (lgProvArtNoInv)
                                {
                                    EXO_392 f392 = new EXO_392();
                                    lRetorno = f392.MenuEvent(infoMenuEvent);
                                    f392 = null;
                                }
                                break;
                                
                        }
                    break;

                case "mProvComp":
                    //Produccion
                    if (lgProvArtNoInv && !infoMenuEvent.BeforeAction)
                    {
                        EXO_Genprov fGenProvi = new EXO_Genprov(true);
                        fGenProvi = null;
                    }
                    break;            
            }
                                        
            return lRetorno;
        }

        public override bool SBOApp_RightClickEvent(ContextMenuInfo infoMenu)
        {
            bool lRetorno = true;
            string cTypeEx = "";
            try
            {
                cTypeEx = Matriz.oGlobal.SBOApp.Forms.Item(infoMenu.FormUID).TypeEx;
            }
            catch (Exception ex)
            {                
                return lRetorno;
            }


            //if (cTypeEx == "134")
            //{
            //    #region Menu Equipos de cliente
            //    if (infoMenu.BeforeAction)
            //    {
            //          try
            //          {
            //              SAPbouiCOM.MenuCreationParams oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Matriz.oGlobal.conexionSAP.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            //              oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            //              oCreationPackage.UniqueID = "mMovCli";
            //              oCreationPackage.String = "Equipos del cliente";
            //              oCreationPackage.Enabled = true;

            //              SAPbouiCOM.MenuItem oMenuItem = Matriz.oGlobal.conexionSAP.SBOApp.Menus.Item("1280"); // Data'
            //              SAPbouiCOM.Menus oMenus = oMenuItem.SubMenus;
            //              oMenus.AddEx(oCreationPackage);
            //          }
            //          catch (Exception ex)
            //          {
            //              Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
            //          }
            //    }
            //    else
            //    {                               
            //        try
            //        {
            //            Matriz.oGlobal.conexionSAP.SBOApp.Menus.RemoveEx("mMovCli");
            //        }
            //        catch (Exception ex)
            //        {
            //            //csVariablesGlobales.SboApp.MessageBox(ex.Message, 1, "Ok", "", "");
            //        }
            //    }
            //    #endregion
            //}

                      
            return lRetorno;

        }

        
        //public static List<ColumnasConsulta> ColumnasMoviNumSerie()
        //{

        //    List<ColumnasConsulta> ListaColumnas = new List<ColumnasConsulta>();
        //    ColumnasConsulta AuxColumn;

        //    #region Objtype
        //    AuxColumn.Width = 0;
        //    AuxColumn.Nombre = "";
        //    AuxColumn.CampoVinculado = "";
        //    AuxColumn.ObjetoVinculado = "";
        //    ListaColumnas.Add(AuxColumn);
        //    #endregion

        //    #region Clave
        //    AuxColumn.Width = 15;
        //    AuxColumn.Nombre = "x";
        //    AuxColumn.CampoVinculado = "ObjType";
        //    AuxColumn.ObjetoVinculado = "";
        //    ListaColumnas.Add(AuxColumn);
        //    #endregion

        //    #region Tipo
        //    AuxColumn.Width = 110;
        //    AuxColumn.Nombre = "";
        //    AuxColumn.CampoVinculado = "";
        //    AuxColumn.ObjetoVinculado = "";
        //    ListaColumnas.Add(AuxColumn);
        //    #endregion

        //    #region Fecha
        //    AuxColumn.Width = 80;
        //    AuxColumn.Nombre = "";
        //    AuxColumn.CampoVinculado = "";
        //    AuxColumn.ObjetoVinculado = "";
        //    ListaColumnas.Add(AuxColumn);
        //    #endregion
            
        //    #region Documento
        //    AuxColumn.Width = 80;
        //    AuxColumn.Nombre = "";
        //    AuxColumn.CampoVinculado = "";
        //    AuxColumn.ObjetoVinculado = "";
        //    ListaColumnas.Add(AuxColumn);
        //    #endregion
                                    
        //    #region Cliente
        //    AuxColumn.Width = 80;
        //    AuxColumn.Nombre = "";
        //    AuxColumn.CampoVinculado = "";
        //    AuxColumn.ObjetoVinculado = "2";
        //    ListaColumnas.Add(AuxColumn);
        //    #endregion

        //    #region Nombre
        //    AuxColumn.Width = 200;
        //    AuxColumn.Nombre = "";
        //    AuxColumn.CampoVinculado = "";
        //    AuxColumn.ObjetoVinculado = "";
        //    ListaColumnas.Add(AuxColumn);
        //    #endregion
                        
        //    #region Cancelado
        //    AuxColumn.Width = 80;
        //    AuxColumn.Nombre = "";
        //    AuxColumn.CampoVinculado = "";
        //    AuxColumn.ObjetoVinculado = "";
        //    ListaColumnas.Add(AuxColumn);
        //    #endregion

        //    return ListaColumnas;

        //}

    }
}
