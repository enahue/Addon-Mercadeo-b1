using SAPbouiCOM;
using System;
using System.Collections.Generic;
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
                        while (threadGetExcelFile.IsAlive)
                        {
                            System.Windows.Forms.Application.DoEvents();
                            Thread.Sleep(50);
                        }

                        var fileName = oGetFileName.FileName;

                        if (!string.IsNullOrEmpty(fileName))
                        {
                            // 1. Leer Excel
                            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = ExcelReaderFactory.CreateReader(stream))
                                {
                                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                    {
                                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                                    });

                                    string hoja = "Hoja1";
                                    if (result.Tables.Contains(hoja))
                                        Texportar = result.Tables[hoja].Copy();
                                    else if (result.Tables.Count > 0)
                                        Texportar = result.Tables[0].Copy();
                                }
                            }

                            // 2. Limpiar y llenar el DataTable de SAPbouiCOM
                            // 1. Obtener el formulario activo directamente desde el evento (más seguro que la variable global oForms)
                            SAPbouiCOM.Form oForm = Conexion.SBOApplication.Forms.Item(FormUID);

                            // 2. Obtener o crear el DataTable
                            SAPbouiCOM.DataTable oDataTable;
                            try
                            {
                                oDataTable = oForm.DataSources.DataTables.Item("DT_Import");
                            }
                            catch
                            {
                                oDataTable = oForm.DataSources.DataTables.Add("DT_Import");
                            }

                            // 3. Forzar la creación de las columnas si no existen (ESTO EVITA EL ERROR "No Fields in Table")
                            if (oDataTable.Columns.Count == 0)
                            {
                                oDataTable.Columns.Add("ItemCode", BoFieldsType.ft_AlphaNumeric, 50);
                                oDataTable.Columns.Add("ItemPrice", BoFieldsType.ft_Price, 50);
                                oDataTable.Columns.Add("FromDate", BoFieldsType.ft_Date, 50);
                                oDataTable.Columns.Add("ToDate", BoFieldsType.ft_Date, 50);
                            }

                            oDataTable.Clear();

                            // Asegurar que existan las columnas antes de agregar filas (evita "No Fields in Table")
                            if (oDataTable.Columns.Count == 0)
                            {
                                try
                                {
                                    oDataTable.Columns.Add("ItemCode", BoFieldsType.ft_AlphaNumeric, 50);
                                    oDataTable.Columns.Add("ItemPrice", BoFieldsType.ft_Price, 50);
                                    oDataTable.Columns.Add("FromDate", BoFieldsType.ft_Date, 50);
                                    oDataTable.Columns.Add("ToDate", BoFieldsType.ft_Date, 50);
                                }
                                catch { }
                            }

                            
                            Conexion.SBOApplication.SetStatusBarMessage("Se han importado " + Texportar.Rows.Count + " registros desde el archivo Excel.", BoMessageTime.bmt_Short, false);

                            if (Texportar.Rows.Count > 0)
                            {
                                // Añadir una fila por cada registro y luego volcar los valores
                                for (int i = 0; i < Texportar.Rows.Count; i++)
                                {
                                    // Añadir una fila vacía
                                    oDataTable.Rows.Add(1);

                                    // Escribir valores en la fila recién añadida (índice i)
                                    try
                                    {
                                        oDataTable.SetValue("ItemCode", i, Texportar.Rows[i].ItemArray.Length > 0 ? (Texportar.Rows[i][0]?.ToString() ?? "") : "");
                                        oDataTable.SetValue("ItemPrice", i, Texportar.Rows[i].ItemArray.Length > 1 ? (Texportar.Rows[i][1]?.ToString() ?? "0") : "0");
                                        oDataTable.SetValue("FromDate", i, Texportar.Rows[i].ItemArray.Length > 2 ? FormatearFechaParaSAP(Texportar.Rows[i][2]) : "");
                                        oDataTable.SetValue("ToDate", i, Texportar.Rows[i].ItemArray.Length > 3 ? FormatearFechaParaSAP(Texportar.Rows[i][3]) : "");
                                    }
                                    catch (Exception exRow)
                                    {
                                        // Registrar y continuar con la siguiente fila
                                        Conexion.SBOApplication.StatusBar.SetText("Fila " + (i + 1) + ": error al asignar valores - " + exRow.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    }
                                }

                                SAPbouiCOM.Item grid = oForm.Items.Item("mtx_import");
                                SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)grid.Specific;

                                // Vincular columnas del UI al DataTable (proteger con try/catch para mensajes claros)
                                try { matrix.Columns.Item("ItemCode").DataBind.Bind("DT_Import", "ItemCode"); } catch { }
                                try { matrix.Columns.Item("ItemPrice").DataBind.Bind("DT_Import", "ItemPrice"); } catch { }
                                try { matrix.Columns.Item("FromDate").DataBind.Bind("DT_Import", "FromDate"); } catch { }
                                try { matrix.Columns.Item("ToDate").DataBind.Bind("DT_Import", "ToDate"); } catch { }

                                // Cargar desde el origen de datos
                                try
                                {
                                    matrix.LoadFromDataSource();
                                }
                                catch (Exception exMatrix)
                                {
                                    Conexion.SBOApplication.MessageBox("Error al cargar la matriz: " + exMatrix.Message);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Conexion.SBOApplication.MessageBox("Error al procesar: " + ex.Message);
                    }
                }
            }


            if (FormUID == "Mercadeo" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "btn_crear" && pVal.ActionSuccess)
            {

                Conexion.SBOApplication.MessageBox("Se ha presionado el botón 'Crear'. Aquí se implementaría la lógica para procesar los datos importados.");
            }


        }

        // Método auxiliar para garantizar que SAP acepte la fecha sin causar errores RPC
        // Acepta valores DateTime, cadenas parseables y números OADate que provienen de Excel
        private string FormatearFechaParaSAP(object valorExcel)
        {
            if (valorExcel == null)
                return "";

            // Si ya es DateTime
            if (valorExcel is DateTime dtValue)
            {
                return dtValue.ToString("yyyyMMdd");
            }

            // Si ExcelDataReader devuelve un número (OADate)
            if (valorExcel is double d)
            {
                try
                {
                    var dt = DateTime.FromOADate(d);
                    return dt.ToString("yyyyMMdd");
                }
                catch
                {
                    return "";
                }
            }

            if (valorExcel is int i)
            {
                try
                {
                    var dt = DateTime.FromOADate(i);
                    return dt.ToString("yyyyMMdd");
                }
                catch
                {
                    return "";
                }
            }

            var str = valorExcel.ToString();
            if (string.IsNullOrWhiteSpace(str))
                return "";

            if (DateTime.TryParse(str, out DateTime fechaParsed))
            {
                return fechaParsed.ToString("yyyyMMdd");
            }

            // Intento: si es formato numérico pero en cadena (p.ej. "43831")
            if (double.TryParse(str, out double d2))
            {
                try
                {
                    var dt = DateTime.FromOADate(d2);
                    return dt.ToString("yyyyMMdd");
                }
                catch
                {
                    return "";
                }
            }

            return "";
        }

        public void SBOApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public void SBOApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
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

            // Cargar valores en el ComboBox de forma segura

            SAPbouiCOM.Item oCombo = oForms.Items.Item("lst_price");
            SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oCombo.Specific;

            if (oComboBox.ValidValues.Count != 0)
            {
                for (int i = 0; i < oComboBox.ValidValues.Count - 1; i++)
                {
                    oComboBox.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            }

            oComboBox.ValidValues.Add("1", "Precio de lista 1");
            oComboBox.ValidValues.Add("2", "Precio de lista 2");

            //Habilitar radio buttons

            SAPbouiCOM.Item ChkAplicar = oForms.Items.Item("ddb_agr");
            SAPbouiCOM.CheckBox ChkAplicarX = (SAPbouiCOM.CheckBox)ChkAplicar.Specific;
            ChkAplicarX.ValOn = "Y";
            ChkAplicarX.ValOff = "N";
            oForms.DataSources.UserDataSources.Add("Aplicar", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            ChkAplicarX.DataBind.SetBound(true, "", "Aplicar");
            oForms.DataSources.UserDataSources.Item("Aplicar").Value = "N";

            SAPbouiCOM.Item ChkQuitar = oForms.Items.Item("rdb_del");
            SAPbouiCOM.CheckBox ChkQuitarX = (SAPbouiCOM.CheckBox)ChkQuitar.Specific;
            ChkQuitarX.ValOn = "Y";
            ChkQuitarX.ValOff = "N";
            oForms.DataSources.UserDataSources.Add("Quitar", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            ChkQuitarX.DataBind.SetBound(true, "", "Quitar");
            oForms.DataSources.UserDataSources.Item("Quitar").Value = "N";

            SAPbouiCOM.Item ChkPrice = oForms.Items.Item("rdb_lp");
            SAPbouiCOM.CheckBox ChkPriceX = (SAPbouiCOM.CheckBox)ChkPrice.Specific;
            ChkPriceX.ValOn = "Y";
            ChkPriceX.ValOff = "N";
            oForms.DataSources.UserDataSources.Add("Price", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            ChkPriceX.DataBind.SetBound(true, "", "Price");
            oForms.DataSources.UserDataSources.Item("Price").Value = "N";



            // ============================================================
            // 1. Obtener o crear el DataTable de forma segura
            // ============================================================
            SAPbouiCOM.DataTable oDataTable;
            try
            {
                oDataTable = oForms.DataSources.DataTables.Item("DT_Import");
            }
            catch
            {
                oDataTable = oForms.DataSources.DataTables.Add("DT_Import");
            }

            // Agregar columnas (try-catch individual por si el XML ya las define)
            try { oDataTable.Columns.Add("ItemCode", BoFieldsType.ft_AlphaNumeric, 50); } catch { }
            try { oDataTable.Columns.Add("ItemPrice", BoFieldsType.ft_Price, 50); } catch { }
            try { oDataTable.Columns.Add("FromDate", BoFieldsType.ft_Date, 50); } catch { }
            try { oDataTable.Columns.Add("ToDate", BoFieldsType.ft_Date, 50); } catch { }

            // ============================================================
            // 2. Configurar la Matriz de forma segura
            // ============================================================
            SAPbouiCOM.Item MyGrid = oForms.Items.Item("mtx_import");
            SAPbouiCOM.Matrix MyMatrix = MyGrid.Specific;
            SAPbouiCOM.Columns oColumns = MyMatrix.Columns;

            SAPbouiCOM.Column oColumn;

            // Columna 1
            try { oColumn = oColumns.Add("ItemCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON); } catch { oColumn = oColumns.Item("ItemCode"); }
            oColumn.TitleObject.Caption = "Código de artículo";
            oColumn.Width = 150;
            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
            oColumn.Editable = false;
            oColumn.Visible = true;
            oColumn.DataBind.Bind("DT_Import", "ItemCode");

            // Columna 2
            try { oColumn = oColumns.Add("ItemPrice", SAPbouiCOM.BoFormItemTypes.it_EDIT); } catch { oColumn = oColumns.Item("ItemPrice"); }
            oColumn.TitleObject.Caption = "Precio de artículo";
            oColumn.Width = 200;
            oColumn.Editable = false;
            oColumn.Visible = true;
            oColumn.DataBind.Bind("DT_Import", "ItemPrice");

            // Columna 3
            try { oColumn = oColumns.Add("FromDate", SAPbouiCOM.BoFormItemTypes.it_EDIT); } catch { oColumn = oColumns.Item("FromDate"); }
            oColumn.TitleObject.Caption = "Fecha de inicio";
            oColumn.Width = 200;
            oColumn.Editable = false;
            oColumn.Visible = true;
            oColumn.DataBind.Bind("DT_Import", "FromDate");

            // Columna 4
            try { oColumn = oColumns.Add("ToDate", SAPbouiCOM.BoFormItemTypes.it_EDIT); } catch { oColumn = oColumns.Item("ToDate"); }
            oColumn.TitleObject.Caption = "Fecha de finalización";
            oColumn.Width = 200;
            oColumn.Editable = false;
            oColumn.Visible = true;
            oColumn.DataBind.Bind("DT_Import", "ToDate");
        }
    }
}