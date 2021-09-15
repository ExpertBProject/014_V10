using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{
    public class EXO_138
    {
        public bool ItemEvent(ItemEvent infoEvento)
        {

            SAPbouiCOM.Form oForm = Matriz.oGlobal.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);

            switch (infoEvento.EventType)
            {
                //En la pestaña de Inventario
                case BoEventTypes.et_FORM_LOAD:
                    if (!infoEvento.BeforeAction)
                    {
                        #region Casilla para activar Provi compras
                        SAPbouiCOM.Item oItem;
                        oItem = oForm.Items.Add("chkProCom", BoFormItemTypes.it_CHECK_BOX );
                        oItem.Left = oForm.Items.Item("1320002088").Left;
                        oItem.Top = oForm.Items.Item("1320002088").Top + 40;
                        oItem.Width = 200;
                        oItem.Height = oForm.Items.Item("1320002088").Height;
                        oItem.FromPane = oForm.Items.Item("1320002088").FromPane;
                        oItem.ToPane = oForm.Items.Item("1320002088").ToPane;
                        ((SAPbouiCOM.CheckBox )oItem.Specific).Caption = "Activar provision compra Art. No Inv";
                        ((SAPbouiCOM.CheckBox )oItem.Specific).ValOn = "Y";
                        ((SAPbouiCOM.CheckBox)oItem.Specific).ValOff = "N";

                        ((SAPbouiCOM.CheckBox)oItem.Specific).DataBind.SetBound(true, "OADM", "U_EXO_ProviComp");
                        #endregion
                                              
                    }
                    break;                
            }
            EXO_CleanCOM.CLiberaCOM.Form(oForm);
            return true;
        }
                
    }
}
