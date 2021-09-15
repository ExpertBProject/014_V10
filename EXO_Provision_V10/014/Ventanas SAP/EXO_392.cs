using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;

namespace Cliente
{
    public class EXO_392
    {

        public bool MenuEvent(MenuEvent args)
        {
            switch (args.MenuUID)
            {

                case "1284": //Menu cancelar para los asientos de provision
                    #region Limpio nombre sales
                    if (args.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = Matriz.oGlobal.SBOApp.Forms.ActiveForm;
                        if (oForm.DataSources.DBDataSources.Item("OJDT").GetValue("U_EXO_EsProvi", 0).Trim() == "Y")
                        {
                            if (oForm.DataSources.DBDataSources.Item("OJDT").GetValue("U_EXO_CancProv", 0).Trim() == "Y")
                            {
                                Matriz.oGlobal.SBOApp.SetStatusBarMessage("Ya ha cancelado este asiento de provision", BoMessageTime.bmt_Short, true);                                
                                return false;
                            }
                            
                            int nAsiento = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OJDT").GetValue("TransId", 0).Trim());
                            EXO_CancProv fCance = new EXO_CancProv(true, oForm.UniqueID, nAsiento);
                            fCance = null;
                            return false;
                        }
                    }
                    #endregion
                    break;
            }
            return true;
        }



    }
}
