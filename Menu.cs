using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Mercadeo
{
    /// <summary>Clase para crear el menú "Mercadeo".</summary>
    class Menu
    {
        /// <summary>Abre conexión y añade los elementos del menú.</summary>
        public Menu()
        {
            Conexion.Open();
            AddMenuItems();

            Validaciones vl = new Validaciones();
        }
        /// <summary>Añade los elementos del menú "Mercadeo" y su opción.</summary>
        public static void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenu;
            SAPbouiCOM.MenuItem oMenuItem;

            oMenu = Conexion.SBOApplication.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage;
            oCreationPackage = Conexion.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

            
            //Agrega Menú personalizado
            oMenuItem = Conexion.SBOApplication.Menus.Item("43520");
            oMenu = oMenuItem.SubMenus;

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "MENU_MERC";
            oCreationPackage.String = "Mercadeo";
            oCreationPackage.Position = oMenuItem.SubMenus.Count + 1;


            if (!(oMenu.Exists("MENU_MERC")))
            {
                oMenu.AddEx(oCreationPackage);
            }
            else
            {
                Conexion.SBOApplication.Menus.RemoveEx("MENU_MERC");
                oMenu.AddEx(oCreationPackage);
            }


            oMenuItem = Conexion.SBOApplication.Menus.Item("MENU_MERC");
            oMenu = oMenuItem.SubMenus;

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
            oCreationPackage.UniqueID = "MENU_OP1";
            oCreationPackage.String = "Importar promociones de productos";
            oCreationPackage.Position = oMenuItem.SubMenus.Count + 1;



            if (!(oMenu.Exists("MENU_OP1")))
            {
                oMenu.AddEx(oCreationPackage);
            }
            else
            {
                Conexion.SBOApplication.Menus.RemoveEx("MENU_OP1");
                oMenu.AddEx(oCreationPackage);
            }



        }

    }
}
