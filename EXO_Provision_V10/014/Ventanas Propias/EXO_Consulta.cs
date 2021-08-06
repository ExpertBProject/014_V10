using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{
    public class EXO_Consulta
    {

        public EXO_Consulta()
        { }

        //public EXO_Consulta(bool lCreacion, string sqlConsulta, List<Matriz.ColumnasConsulta> ListaColumnas, string Titulo = "")
        //{

        //    if (lCreacion)
        //    {
        //        SAPbouiCOM.Form oForm = null;

        //        SAPbouiCOM.FormCreationParams oParametrosCreacion = (SAPbouiCOM.FormCreationParams)(Matriz.oGlobal.conexionSAP.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams));          
        //        string strXML = Utilidades.LeoFichEmbebido("EXO_Consulta.srf");
        //        oParametrosCreacion.XmlData = strXML;
        //        //oParametrosCreacion.UniqueID = "";
        //        //oParametrosCreacion.BorderStyle = BoFormBorderStyle.fbs_Fixed;

        //        try
        //        {
        //            oForm = Matriz.oGlobal.conexionSAP.SBOApp.Forms.AddEx(oParametrosCreacion);
        //        }
        //        catch (Exception ex)
        //        {
        //            Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");

        //        }

        //        SAPbouiCOM.DataTable oTabla = oForm.DataSources.DataTables.Item("TablaReg");
        //        oTabla.ExecuteQuery(sqlConsulta);
        //        SAPbouiCOM.Grid oGrd = (SAPbouiCOM.Grid) oForm.Items.Item("GrdCon").Specific;

        //        int nWidthGrid = 0;
        //        for (int j = 0; j < oGrd.Columns.Count; j++)
        //        {
        //            SAPbouiCOM.EditTextColumn oColumn = (SAPbouiCOM.EditTextColumn)oGrd.Columns.Item(j);
        //            if (ListaColumnas.ElementAt(j).Width == 0)
        //            {
        //                oColumn.Visible = false;
        //            }
        //            else
        //            {
        //                oColumn.Width = ListaColumnas.ElementAt(j).Width;
        //                nWidthGrid += oColumn.Width;
        //            }

        //            if (ListaColumnas.ElementAt(j).Nombre != "") oColumn.TitleObject.Caption = ListaColumnas.ElementAt(j).Nombre;

        //            if (ListaColumnas.ElementAt(j).CampoVinculado != "" || ListaColumnas.ElementAt(j).ObjetoVinculado != "")
        //            {
        //                if (ListaColumnas.ElementAt(j).ObjetoVinculado != "")
        //                {
        //                    oColumn.LinkedObjectType = ListaColumnas.ElementAt(j).ObjetoVinculado;                            
        //                }

        //                if (ListaColumnas.ElementAt(j).CampoVinculado != "")
        //                {
        //                    oColumn.LinkedObjectType = "143"; 
        //                    oColumn.Description = "#" + ListaColumnas.ElementAt(j).CampoVinculado;
        //                }                        
        //            }                                        
        //        }

        //        //Coloco el grid
        //        oForm.Items.Item("GrdCon").Width = nWidthGrid + 25;
        //        oForm.Width = oForm.Items.Item("GrdCon").Width + 60;
        //        oForm.Items.Item("GrdCon").Left = 20;

        //        oForm.Visible = true;
        //        oForm.Title = Titulo;
        //    }
        //}

        //public bool ItemEvent(EXO_Generales.EXO_infoItemEvent infoEvento)
        //{

        //    switch (infoEvento.EventType)
        //    {
        //        case BoEventTypes.et_MATRIX_LINK_PRESSED:
        //            if (infoEvento.ItemUID == "GrdCon" && infoEvento.BeforeAction  )
        //            {

        //                SAPbouiCOM.Form oForm = Matriz.oGlobal.conexionSAP.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);
        //                SAPbouiCOM.Grid oGrd = (SAPbouiCOM.Grid)oForm.Items.Item("GrdCon").Specific;
        //                SAPbouiCOM.EditTextColumn oColumn = (SAPbouiCOM.EditTextColumn)oGrd.Columns.Item(infoEvento.ColUID);
        //                if (oColumn.Description.Length > 0 &&  oColumn.Description.Substring(0, 1) == "#")
        //                {
        //                    string cCampoVinc = oColumn.Description.Substring(1);
        //                    string cTipo = oGrd.DataTable.GetValue(cCampoVinc, oGrd.GetDataTableRowIndex(infoEvento.Row)).ToString();
        //                    oColumn.LinkedObjectType = cTipo;
        //                }                                                                        
        //            }

        //            break;
        //    }
        //    return true;
        //  }

           
    }
}
