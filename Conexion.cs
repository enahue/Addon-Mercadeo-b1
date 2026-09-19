using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mercadeo
{
    public class Conexion
    {
        public static SAPbobsCOM.Company oCompany;
        public static SAPbouiCOM.Application SBOApplication;

        public static bool Open()
        {
            bool result = false;

            try
            {
                SAPbouiCOM.SboGuiApi SboGuiApi = new SAPbouiCOM.SboGuiApi();
                SboGuiApi.Connect(Environment.GetCommandLineArgs().GetValue(1).ToString());
                SBOApplication = SboGuiApi.GetApplication();
                SboGuiApi = null;

                oCompany = (SAPbobsCOM.Company)SBOApplication.Company.GetDICompany();
                if (oCompany.Connected)
                {
                    result = true;
                    SBOApplication.StatusBar.SetText("Conexión exitosa a SAP Business One.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    result = false;
                    SBOApplication.StatusBar.SetText("Error al conectar a SAP Business One.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return result;

        }

    }
}
