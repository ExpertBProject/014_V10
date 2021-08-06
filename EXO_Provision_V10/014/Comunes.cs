using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Reflection;
using Microsoft.VisualBasic.CompilerServices;

namespace Cliente
{
    public struct RepartosCeCos
    {
        public string CeCo;
        public decimal Importe;
    }


    public class Conversiones
    {

        public static double ValueSAPToDoubleSistema(string Texto)
        {
            string Cadena = Texto;

            double Valor = 0.0;
            System.Globalization.NumberFormatInfo nfi = new System.Globalization.NumberFormatInfo();
            string SepDecSistema = nfi.NumberGroupSeparator;
            //string SepMilSistema = nfi.NumberDecimalSeparator;

            //En pantalla el separador decimal es .
            if (SepDecSistema != ".")
            {
                Cadena = Cadena.Replace('.', ',');
            }
            double.TryParse(Cadena, out Valor);
            return Valor;
        }

        public static double StringSAPToDoubleSistema(string Texto)
        {
            double nRetorno = 0;

            //Quito la moneda y el sep miles
            string Cadena = Texto;
            Cadena = Cadena.Replace(Matriz.SepMill, "");

            System.Globalization.NumberFormatInfo nfi = new System.Globalization.NumberFormatInfo();
            string SepDecSistema = nfi.NumberGroupSeparator;
            Cadena = Cadena.Replace(Matriz.SepDec, SepDecSistema);
            double.TryParse(Cadena, out nRetorno);

            return nRetorno;
        }

        public static double StringSAPToDoubleSistema(string Texto, string cMoneda)
        {
            double nRetorno = 0;
            string Cadena = Texto.Replace((cMoneda != "") ? cMoneda : "EUR", "");

            nRetorno = StringSAPToDoubleSistema(Cadena);

            return nRetorno;
        }

      public static string DoubleStringSAP(double Valor, BoFldSubTypes BoTipo)
            {

                string cRetorno = "";

                switch (BoTipo)
                {
                    case BoFldSubTypes.st_Quantity:
                        Valor = Math.Round(Valor, Matriz.QtyDec);
                        break;
                    case BoFldSubTypes.st_Sum:
                        Valor = Math.Round(Valor, Matriz.SumDec);
                        break;
                    case BoFldSubTypes.st_Percentage:
                        Valor = Math.Round(Valor, Matriz.PercentDec);
                        break;
                    case BoFldSubTypes.st_Price:
                        Valor = Math.Round(Valor, Matriz.PriceDec);
                        break;
                    case BoFldSubTypes.st_Measurement:
                        Valor = Math.Round(Valor, Matriz.MeasureDec);
                        break;
                    case BoFldSubTypes.st_Rate:
                        Valor = Math.Round(Valor, Matriz.RateDec);
                        break;
                    default:
                        Valor = Math.Round(Valor, 2);
                        break;
                }

                cRetorno = Valor.ToString();
                cRetorno = cRetorno.Replace(',', '.');
                return cRetorno;

                //string cRetorno = "";
                //string cAux = Valor.ToString();

                //cRetorno = cAux.Replace(',', csVariablesGlobales.cSepDecimal);

                ////System.Globalization.NumberFormatInfo nfi = new System.Globalization.NumberFormatInfo();

                ////string hh = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
                ////string hh1 = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;

                ////string SepDecSAP = DevuelveValor("OADM", "DecSep", "");
                ////string SepMilSAP = DevuelveValor("OADM", "ThousSep", ""); 
                ////string SepDecSistema = nfi.NumberGroupSeparator;
                ////string SepMilSistema = nfi.NumberDecimalSeparator;
                ////string ValorDevuelto = Valor.ToString();
                ////if (SepDecSAP != SepDecSistema)
                ////{
                ////    ValorDevuelto = ValorDevuelto.Replace(SepDecSistema, SepDecSAP);
                ////}


            }
    }

    public class TratamientoFicheros
    {
       public static string EscojoFichero(string cCadenaDefect)
            {
                string Ruta = "";
           
                EXO_SaveFileDialog oFichero = new EXO_SaveFileDialog();
                oFichero.Filter = "All Files (*)|*|Dat (*.dat)|*.dat|Text Files (*.txt)|*.txt";
                oFichero.FileName = cCadenaDefect;
                string DirectorioActual = Environment.CurrentDirectory;
                Thread threadGetFile = new Thread(new ThreadStart(oFichero.GetFileName));
                threadGetFile.TrySetApartmentState(ApartmentState.STA);
                threadGetFile.Start();
                try
                {
                    while (!threadGetFile.IsAlive) ; // Wait for thread to get started
                    Thread.Sleep(1);  // Wait a sec more
                    threadGetFile.Join();    // Wait for thread to end

                    Ruta = oFichero.FileName;
                }
                catch (Exception ex)
                {
                    Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(ex.Message, 1, "OK", "", "");
                }
                threadGetFile = null;
                oFichero = null;

                return Ruta;
            }

       public static string SeleccionoFichero(string cCadenaDefect)
            {
                string cRetorno = "";

                EXO_OpenFileDialog OpenFileDialog = new EXO_OpenFileDialog();
                OpenFileDialog.Filter = "Todos los ficheros|*.*";
                OpenFileDialog.InitialDirectory = "";
                Thread threadGetFile = new Thread(new ThreadStart(OpenFileDialog.GetFileName));
                threadGetFile.TrySetApartmentState(ApartmentState.STA);
                try
                {
                    
                    threadGetFile.Start();
                    while (!threadGetFile.IsAlive) ; // Wait for thread to get started
                    Thread.Sleep(1);  // Wait a sec more
                    threadGetFile.Join();    // Wait for thread to end

                    // Use file name as you will here
                    cRetorno = OpenFileDialog.FileName;
                    threadGetFile.Abort();
                    threadGetFile = null;
                    OpenFileDialog.InitialDirectory = "";
                    OpenFileDialog = null;
                }
                catch (Exception ex)
                {
                    Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(ex.Message, 1, "OK", "", "");
                    threadGetFile.Abort();
                    threadGetFile = null;
                    OpenFileDialog.InitialDirectory = "";
                    OpenFileDialog = null;

                }

                return cRetorno;
            }

       public static bool IsDirectoryWritable(string dirPath, bool throwIfFails = false)
       {
           try
           {
               using (FileStream fs = File.Create(Path.Combine(dirPath, Path.GetRandomFileName()), 1, FileOptions.DeleteOnClose)
               )
               { }
               return true;
           }
           catch
           {
               if (throwIfFails)
                   throw;
               else
                   return false;
           }
       }

    }

    public class WindowWrapper : System.Windows.Forms.IWin32Window
    {
        private IntPtr _hwnd;

        // Property
        public virtual IntPtr Handle
        {
            get { return _hwnd; }
        }

        // Constructor
        public WindowWrapper(IntPtr handle)
        {
            _hwnd = handle;
        }
    }

    public class Utilidades
    {               

        public static string FormUDO(string cFormSRF, string cUDO, ref bool lError)
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            GC.Collect();            
            oUserObjectMD = (SAPbobsCOM.UserObjectsMD)  Matriz.oGlobal.conexionSAP.compañia.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
            int lRetCode;
            string cMensaje = "";

            try
            {
                oUserObjectMD.GetByKey(cUDO);
                oUserObjectMD.EnableEnhancedForm = BoYesNoEnum.tYES;
                oUserObjectMD.RebuildEnhancedForm = BoYesNoEnum.tNO;                
                oUserObjectMD.FormSRF = cFormSRF;
                lRetCode = oUserObjectMD.Update();

                cMensaje = (lRetCode != 0) ? Matriz.oGlobal.conexionSAP.compañia.GetLastErrorDescription() : "Actualizado el UDO " + cUDO;
                lError = (lRetCode != 0);
            }
            catch (Exception ex)
            {
                cMensaje = ex.Message;
                lError = true;
            }

            oUserObjectMD = null;
            GC.Collect(); //Release the handle to the table

            return cMensaje;
        }

        public static void BorroDataTable(ref SAPbouiCOM.DataTable oTablaInf)
        {
            if (!oTablaInf.IsEmpty)
            {
                int nNumReg = oTablaInf.Rows.Count;
                for (int i = 0; i < nNumReg; i++)
                {
                    oTablaInf.Rows.Remove(0);
                }
            }
        }

        public static SAPbouiCOM.Form BuscoFormLanzado(string cFormUID)
        {
            SAPbouiCOM.Form oFORMORET = null;

            try
            {
                oFORMORET = Matriz.oGlobal.conexionSAP.SBOApp.Forms.Item(cFormUID);
            }
            catch (Exception EX)
            { }
            
            return oFORMORET;
        }

        public static bool LanzoMenuUserTable(string cTablaSinArroba, bool lUDO)
        {
            bool lRetorno = false;
            SAPbouiCOM.Menus oMenus = Matriz.oGlobal.conexionSAP.SBOApp.Menus.Item(lUDO ? "47616":"51200").SubMenus;
            for (int i = 0; i <= oMenus.Count - 1; i++)
            {
                if (oMenus.Item(i).String.IndexOf(cTablaSinArroba) == 0)
                {
                    Matriz.oGlobal.conexionSAP.SBOApp.ActivateMenuItem(oMenus.Item(i).UID);
                    lRetorno = true;
                    break;
                }
            }


            EXO_CleanCOM.CLiberaCOM.Menus(ref oMenus);
            return lRetorno;
        }


        public static bool LanzoQueryPorMenu(string cCategoria, string cNomQuery)
        {
            bool lRetorno = false;
            SAPbouiCOM.Menus oMenus = Matriz.oGlobal.conexionSAP.SBOApp.Menus.Item("53248").SubMenus;
            SAPbouiCOM.Menus oMenusCons;
            for (int i = 0; i <= oMenus.Count - 1; i++)
            {                
                if (oMenus.Item(i).String.IndexOf(cCategoria) == 0)
                {
                    oMenusCons = oMenus.Item(i).SubMenus;
                    for (int j = 0; j <= oMenusCons.Count; j++)
                    {
                        if ( oMenusCons.Item(j).String.IndexOf(cNomQuery) == 0)
                        {
                            Matriz.oGlobal.conexionSAP.SBOApp.ActivateMenuItem(oMenusCons.Item(j).UID);
                            lRetorno = true;
                            break;
                        }
                    }               
                }
            }


            EXO_CleanCOM.CLiberaCOM.Menus(ref oMenus);
            return lRetorno;
        }



        //public static string LeoQuerySinEmbebida(string cNomFichLargo)
        //{
        //    string sql = "", cAux = "";
        //    System.IO.StreamReader Fichero = new System.IO.StreamReader(cNomFichLargo);
        //    while (Fichero.Peek() != -1)
        //    {
        //        cAux = Fichero.ReadLine();
        //        if (cAux.Length > 2 && cAux.Substring(0, 2) == "--") continue;

        //        sql += cAux.Replace("\t", " ") + " ";
        //    }
        //    Fichero.Close();

        //    return sql;
        //}

        public static string LeoFichEmbebido(string cFichEmbebido)
        {
            string result = "";
            try
            {
                Type tipo = Matriz.TypeMatriz;
                Assembly assembly = tipo.Assembly;
                StreamReader streamReader = new StreamReader(tipo.Assembly.GetManifestResourceStream(tipo.Namespace + "." + cFichEmbebido));
                result = streamReader.ReadToEnd();
                result = result.Replace("\t", " ").Replace("\n", " ").Replace("\r", " ");
                streamReader.Close();
            }
            catch (Exception expr_40)
            {
                ProjectData.SetProjectError(expr_40);
                ProjectData.ClearProjectError();
            }            

            return result;
        }

        public static void BorroLineaMatrix(ref SAPbouiCOM.Matrix oMatrix, ref SAPbouiCOM.Form oFormulario)
        {

            oFormulario.Freeze(true);
            try
            {
                oMatrix.FlushToDataSource();
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    if (oMatrix.IsRowSelected(i))
                    {
                        oMatrix.DeleteRow(i);
                        if (oFormulario.Mode == BoFormMode.fm_OK_MODE) oFormulario.Mode = BoFormMode.fm_UPDATE_MODE;
                        break;
                    }
                }

                oMatrix.FlushToDataSource();
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
            }

            oFormulario.Freeze(false);
        }

        public static string DevuelvoNombreTabla(string cTipoEx)
        {
            string cTabla = "";

            #region Tabla
            switch (cTipoEx)
            {
                //Albaran
                case "140":
                    cTabla = "ODLN";
                    break;
                //Pedido
                case "139":
                    cTabla = "ORDR";
                    break;
                //Oferta
                case "149":
                    cTabla = "OQUT";
                    break;
                //Oferta
                case "60091":
                    cTabla = "OINV";
                    break;
                //Factura
                case "133":
                    cTabla = "OINV";
                    break;
                //Abono
                case "179":
                    cTabla = "ORIN";
                    break;
                //Devolucion
                case "180":
                    cTabla = "ORDN";
                    break;
            }
            #endregion

            return cTabla;

        }

        public static void LLenoComboGenerico(ref SAPbouiCOM.Item oItemCombo, string cTabla, string cWhere = "")
        {
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oItemCombo.Specific;
            string sql = "SELECT T0.Code, T0.Name FROM [" + cTabla + "] T0 " + cWhere + " ORDER BY T0.Name ";
            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.SQL.sqlComoRsB1(sql);
            while (!oRec.EoF)
            {
                oCombo.ValidValues.Add(oRec.Fields.Item(0).Value, oRec.Fields.Item(1).Value);
                oRec.MoveNext();
            }
            oItemCombo.DisplayDesc = true;
            oCombo.ExpandType = BoExpandType.et_DescriptionOnly;
        }

        public static void LLenoComboGenerico(ref SAPbouiCOM.Column oColumCombo, string cTabla, string cWhere = "")
        {
            string sql = "SELECT T0.Code, T0.Name FROM [" + cTabla + "] T0 " + cWhere + " ORDER BY T0.Name ";
            SAPbobsCOM.Recordset oRec = Matriz.oGlobal.SQL.sqlComoRsB1(sql);
            while (!oRec.EoF)
            {
                oColumCombo.ValidValues.Add(oRec.Fields.Item(0).Value, oRec.Fields.Item(1).Value);
                oRec.MoveNext();
            }
            oColumCombo.DisplayDesc = true;
            oColumCombo.ExpandType = BoExpandType.et_DescriptionOnly;
        }

        public static void ActualizarBBDD(string sqlUPD, ref string cMenError)
        {

            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)Matriz.oGlobal.conexionSAP.compañia.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                oRec.DoQuery(sqlUPD);
            }
            catch (Exception EX)
            {
                cMenError = EX.Message;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
                oRec = null;
            }
        }

        public static void Reposiciono(string cFormUID, int nAbsEntry)
        {

            SAPbouiCOM.Form oForm = Utilidades.BuscoFormLanzado(cFormUID);
            oForm.Select();

            oForm.Freeze(true);
            try
            {
                //string cNumDoc = (nAbsEntry == 0) ? ((SAPbouiCOM.EditText)oForm.Items.Item("8").Specific).Value : nAbsEntry.ToString();
                oForm.Mode = BoFormMode.fm_FIND_MODE;
                switch (oForm.TypeEx)
                {
                    case "392":
                        ((SAPbouiCOM.EditText)oForm.Items.Item("5").Specific).Value = nAbsEntry.ToString();
                        oForm.Items.Item("1").Click(BoCellClickType.ct_Regular);
                        break;
                }                
            }
            catch (Exception ex)
            {
                //Matriz.oGlobal.conexionSAP.SBOApp.MessageBox(ex.Message, 1, "Ok", "", "");
            }
            oForm.Freeze(false);
        }


        //HAY QUE LANZARLA DENTRO DE UNA TRANSACION
        public static string CreoNormaRepartoManual(ref string cMenError, int nDimension, List<RepartosCeCos> ListaDatos)
        {
            //HAY QUE LANZARLA DENTRO DE UNA TRANSACION
            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)Matriz.oGlobal.conexionSAP.compañia.GetBusinessObject(BoObjectTypes.BoRecordset);
            string cRetorno = "";
            bool lErrorTransaction = false;
            string sqlUpd = "";

            try
            {
                #region Busco el contador de repartos manuales
                string sql = "select top 1 AutoKey from onnm where objectcode = '252'";
                int nNumerador = Convert.ToInt32(Matriz.oGlobal.SQL.sqlNumericaB1(sql));
                if (nNumerador == 0)
                {
                    cMenError = "No se pudo recuperar el numerador de repartos manuales";
                    lErrorTransaction = true;
                }
                #endregion

                if (!lErrorTransaction)
                {
                    #region Sumo los totales
                    decimal nTotalReparto = 0;
                    foreach (RepartosCeCos AuxDatos in ListaDatos)
                    {
                        nTotalReparto += AuxDatos.Importe;
                    }
                    #endregion

                    string cCodNorma = "M" + nNumerador.ToString("0000000");

                    sqlUpd = "UPDATE ONNM SET AUTOKEY =" + (nNumerador + 1).ToString() + " WHERE OBJECTCODE = '252';";

                    #region Para la cabecera
                    sqlUpd += "INSERT INTO OMDR (OcrCode,OcrName,OcrTotal,Direct,Locked,DataSource,UserSign,DimCode,AbsEntry,Active)";
                    sqlUpd += " VALUES (";
                    sqlUpd += "'" + cCodNorma + "'";
                    sqlUpd += ",'Norma de reparto manual'";

                    string cAux = nTotalReparto.ToString();
                    cAux = cAux.Replace(".", "");
                    cAux = cAux.Replace(",", ".");
                    sqlUpd += "," + cAux;
                    sqlUpd += ",'N'";
                    sqlUpd += ",'N'";
                    sqlUpd += ",'I'";
                    sqlUpd += "," + Matriz.oGlobal.conexionSAP.compañia.UserSignature;
                    sqlUpd += "," + nDimension.ToString();
                    sqlUpd += "," + nNumerador.ToString();
                    sqlUpd += ",'Y'";
                    sqlUpd += " )";
                    sqlUpd += ";";
                    #endregion

                    #region Para las lineas
                    foreach (RepartosCeCos AuxDato in ListaDatos)
                    {
                        sqlUpd += "INSERT INTO MDR1 (OcrCode,PrcCode,PrcAmount,OcrTotal,Direct,ValidFrom)";
                        sqlUpd += " VALUES (";
                        sqlUpd += "'" + cCodNorma + "'";
                        sqlUpd += ",'" + AuxDato.CeCo + "'";

                        cAux = AuxDato.Importe.ToString();
                        cAux = cAux.Replace(".", "");
                        cAux = cAux.Replace(",", ".");
                        sqlUpd += "," + cAux;

                        cAux = nTotalReparto.ToString();
                        cAux = cAux.Replace(".", "");
                        cAux = cAux.Replace(",", ".");
                        sqlUpd += "," + cAux;

                        sqlUpd += ",'N'";
                        sqlUpd += ",'1900-01-01 00:00:00.000'";
                        sqlUpd += " )";
                        sqlUpd += ";";
                    }
                    #endregion

                    //Todo de un tiron
                    oRec.DoQuery(sqlUpd);

                    //Si llego, es que esta bien
                    cRetorno = cCodNorma;
                }
            }
            catch (Exception ex)
            {
                cMenError = ex.Message;
            }
            finally
            { }
            return cRetorno;
        }



    }



}
