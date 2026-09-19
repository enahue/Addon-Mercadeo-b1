using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mercadeo
{
    /// <summary>
    /// Clase que gestiona eventos de menú y crea el formulario de importación.
    /// </summary>
    class Validaciones
    {
        private Form oForms;
        /// <summary>Inicializa la conexión y registra el manejador de eventos del menú.</summary>
        public Validaciones()
        {
            Conexion.Open();

            Conexion.SBOApplication.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(SBOApplication_MenuEvent);
        }
        /// <summary>Manejador del evento de menú; abre el formulario cuando se pulsa la opción del complemento.</summary>
        public void SBOApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            //Llamada al Formulario

            if (pVal.MenuUID == "MENU_OP1" && pVal.BeforeAction == true)
            {
                CreateForm();

            }


        }

        /// <summary>Carga el XML del formulario y lo muestra en SAP Business One.</summary>
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

            // Configuración de la columna "ItemCode" como un botón vinculado

            SAPbouiCOM.Item MyGrid = oForms.Items.Item("mtx_import");
            SAPbouiCOM.Matrix MyMatrix = MyGrid.Specific;
            SAPbouiCOM.Columns oColumns = MyMatrix.Columns;

            SAPbouiCOM.LinkedButton oLink = null;
            SAPbouiCOM.Column oColumn1 = oColumns.Add("ItemCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn1.TitleObject.Caption = "Código de artículo";
            oColumn1.Width= 150;
            oLink = oColumn1.ExtendedObject;
            oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
            oColumn1.Editable = false;
            oColumn1.Visible = true;


            SAPbouiCOM.Column oColumn2 = oColumns.Add("ItemPrice", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn2.TitleObject.Caption = "Precio de artículo";
            oColumn2.Width= 200;
            oColumn2.Editable = false;
            oColumn2.Visible = true;

            SAPbouiCOM.Column oColumn3 = oColumns.Add("FromDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn3.TitleObject.Caption = "Fecha de inicio";
            oColumn3.Width = 200;
            oColumn3.Editable = false;
            oColumn3.Visible = true;

            SAPbouiCOM.Column oColumn4 = oColumns.Add("ToDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn4.TitleObject.Caption = "Fecha de finalización";
            oColumn4.Width = 200;
            oColumn4.Editable = false;
            oColumn4.Visible = true;




        }

    }
}