using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{
    public class EXO_142
    {
        public bool ItemEvent(ItemEvent infoEvento)
        {
            SAPbouiCOM.Form oForm = Matriz.oGlobal.SBOApp.Forms.GetForm(infoEvento.FormTypeEx, infoEvento.FormTypeCount);

            switch (infoEvento.EventType)
            {
                case BoEventTypes.et_FORM_LOAD:
                    if (infoEvento.BeforeAction)
                    {                     
                        SAPbouiCOM.Item oItem;
                        
                        #region Casilla para no provisionar
                        oItem = oForm.Items.Add("chkProv", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                        oItem.Left = oForm.Items.Item("135").Left;
                        oItem.Width = 180;
                        oItem.Top = oForm.Items.Item("135").Top + oForm.Items.Item("135").Height + 5;
                        oItem.Height = oForm.Items.Item("135").Height;
                        oItem.LinkTo = "157";
                        oItem.FromPane = oForm.Items.Item("157").FromPane;
                        oItem.ToPane = oForm.Items.Item("157").ToPane;
                        ((SAPbouiCOM.CheckBox)oItem.Specific).Caption = "Provisionar Art. no inventariables";
                         ((SAPbouiCOM.CheckBox)oItem.Specific).ValOff = "N";
                        ((SAPbouiCOM.CheckBox)oItem.Specific).ValOn = "Y";
                        ((SAPbouiCOM.CheckBox)oItem.Specific).DataBind.SetBound(true, "OPOR", "U_EXO_ProviComp");                  
                        #endregion
                        
                        #region Enlace con el asiento de transporte
                        oItem = (SAPbouiCOM.Item)oForm.Items.Add("txtAsiTra", BoFormItemTypes.it_EDIT);
                        oItem.Width = oForm.Items.Item("134").Width;
                        oItem.Top = oForm.Items.Item("chkProv").Top + oForm.Items.Item("chkProv").Height + 10;
                        oItem.Left = oForm.Items.Item("157").Left;
                        oItem.Height = oForm.Items.Item("157").Height;
                        oItem.LinkTo = "157";
                        oItem.DisplayDesc = true;
                        oItem.AffectsFormMode = false;
                        oItem.FromPane = oForm.Items.Item("157").FromPane;
                        oItem.ToPane = oForm.Items.Item("157").ToPane;
                        oItem.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, -1, BoModeVisualBehavior.mvb_False);
                        ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "OPOR", "U_EXO_AsiProvComp");


                        oItem = (SAPbouiCOM.Item)oForm.Items.Add("lblAsiTra", BoFormItemTypes.it_STATIC);
                        oItem.Width = oForm.Items.Item("156").Width - 20;
                        oItem.Top = oForm.Items.Item("txtAsiTra").Top;
                        oItem.Left = oForm.Items.Item("156").Left;
                        oItem.Height = oForm.Items.Item("156").Height;
                        oItem.LinkTo = "txtAsiTra";
                        oItem.FromPane = oForm.Items.Item("txtAsiTra").FromPane;
                        oItem.ToPane = oForm.Items.Item("txtAsiTra").ToPane;
                        ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Asiento Provision.";


                        oItem = (SAPbouiCOM.Item)oForm.Items.Add("lkTrans", BoFormItemTypes.it_LINKED_BUTTON);
                        oItem.Width = 15;
                        oItem.Top = oForm.Items.Item("txtAsiTra").Top;
                        oItem.Left = oForm.Items.Item("157").Left - 15;
                        oItem.Height = oForm.Items.Item("157").Height;
                        oItem.LinkTo = "txtAsiTra";
                        oItem.AffectsFormMode = false;
                        oItem.FromPane = oForm.Items.Item("txtAsiTra").FromPane;
                        oItem.ToPane = oForm.Items.Item("txtAsiTra").ToPane;
                        ((SAPbouiCOM.LinkedButton)oItem.Specific).LinkedObject = BoLinkedObject.lf_JournalPosting;
                        #endregion              

                    }
                    break;
            }

            return true;
        }

        public bool DataEvent(BusinessObjectInfo args)
        {
            SAPbouiCOM.Form oForm = Matriz.oGlobal.SBOApp.Forms.Item(args.FormUID);

            if (args.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && args.BeforeAction)
            {
                #region Para limpiar a casilla del asiento de provision                
                try
                {
                        //Limpio lo del asiento de transporte
                        ((SAPbouiCOM.EditText)oForm.Items.Item("txtAsiTra").Specific).Value = "";
                }
                catch (Exception ex)
                { }
                
                #endregion
            }

            return true;
        }
    }
}

