using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using UI_API_Csharp;
using System.IO;
using ExcelDataReader;

namespace Mercadeo
{
    
    class Mercadeo
    {
        private Form oForms;
        public System.Data.DataTable Texportar = new System.Data.DataTable();
        public Mercadeo()
        {
            Conexion.Open();

            Conexion.SBOApplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBOApplication_ItemEvent);
            Conexion.SBOApplication.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBOApplication_FormDataEvent);
            Conexion.SBOApplication.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(SBOApplication_MenuEvent);

        }

        public void SBOApplication_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (FormUID == "Mercadeo" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "btn_xls" && pVal.ActionSuccess)
            {

                using (GetFileNameClass oGetFileName = new GetFileNameClass())
                {

                    oGetFileName.Filter = "Excel files (*.xlsx)|*.xlsx";
                    oGetFileName.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                    Thread threadGetExcelFile = new Thread(new ThreadStart(oGetFileName.GetFileName));
                    threadGetExcelFile.SetApartmentState(ApartmentState.STA);
                    Texportar.Clear();
                    Texportar.Columns.Clear();
                    Texportar.Rows.Clear();

                    try
                    {

                        threadGetExcelFile.Start();
                        while (!threadGetExcelFile.IsAlive) ;
                        Thread.Sleep(1);
                        threadGetExcelFile.Join();

                        var fileName = string.Empty;
                        fileName = oGetFileName.FileName;

                        if (!string.IsNullOrEmpty(fileName))
                        {
                            string archivo = fileName;
                            string hoja = "Hoja1"; 

                            using (var stream = File.Open(archivo, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {

                                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                    {
                                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                        {
                                            UseHeaderRow = true
                                        }
                                    });

                                    if (result.Tables.Contains(hoja))
                                    {
                                        Texportar = result.Tables[hoja].Copy();
                                    }
                                    else if (result.Tables.Count > 0)
                                    {
                                        Texportar = result.Tables[0].Copy();
                                    }
                                }
                            }

                            //Conexion.SBOApplication.MessageBox(Texportar.Rows.Count.ToString());

                            DataTable oDataTable = null;
                            if(oForms.DataSources.DataTables.Count.Equals(0))
                            {
                                oForms.DataSources.DataTables.Add("DT_Import");
                            }
                            else
                            {
                               oForms.DataSources.DataTables.Item("DT_Import").Clear();
                            }
                            oDataTable = oForms.DataSources.DataTables.Item("DT_Import");

                            oDataTable.Columns.Add("ItemCode", BoFieldsType.ft_AlphaNumeric, 50);
                            oDataTable.Columns.Add("ItemPrice", BoFieldsType.ft_Price, 50);
                            oDataTable.Columns.Add("FromDate", BoFieldsType.ft_Date, 50);
                            oDataTable.Columns.Add("ToDate", BoFieldsType.ft_Date, 50);

                            oDataTable.Rows.Add(Texportar.Rows.Count);

                            for (int i = 1; i <= Texportar.Rows.Count; i++)
                            {
                                oDataTable.SetValue("ItemCode", i, Texportar.Rows[i][0].ToString());
                                oDataTable.SetValue("ItemPrice", i, Texportar.Rows[i][1].ToString());
                                oDataTable.SetValue("FromDate", i, Texportar.Rows[i][2].ToString());
                                oDataTable.SetValue("ToDate", i, Texportar.Rows[i][3].ToString());
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        Conexion.SBOApplication.MessageBox("Error: " + ex.Message);
                    }

                }
            }
        }

        public void SBOApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

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
            oColumn1.Width = 150;
            oLink = oColumn1.ExtendedObject;
            oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
            oColumn1.Editable = false;
            oColumn1.Visible = true;


            SAPbouiCOM.Column oColumn2 = oColumns.Add("ItemPrice", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn2.TitleObject.Caption = "Precio de artículo";
            oColumn2.Width = 200;
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
