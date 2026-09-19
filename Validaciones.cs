using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mercadeo
{
    class Validaciones
    {
        private Form oForms;
        public Validaciones()
        {
            Conexion.Open();

            Conexion.SBOApplication.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(SBOApplication_MenuEvent);
        }
        public void SBOApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            //Llamada al Formulario

            if (pVal.MenuUID == "MENU_OP1" && pVal.BeforeAction == true)
            {
                CreateForm();

            }


        }

        private void CreateForm()
        {
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string directory = System.IO.Path.GetDirectoryName(path);
            directory += "\\Formulario.xml";

            System.Xml.XmlDocument oXmlDoc = new System.Xml.XmlDocument();
            oXmlDoc.Load(directory);

            SAPbouiCOM.FormCreationParams oCreationParams = Conexion.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
            oCreationParams.XmlData = oXmlDoc.InnerXml;

            oForms = Conexion.SBOApplication.Forms.AddEx(oCreationParams);
            oForms.Visible = true;
        }

    }
}