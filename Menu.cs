using Addon_Messem.FORMS.UserForms;
using SAPbouiCOM.Framework;
using System;

namespace Addon_Messem
{
    class Menu
    {
        //static SAPbouiCOM.Form MainMenu;
        /// <summary>
        /// Creates the SAPs Menus
        /// </summary>
        public static void AddMenuItems()
        {
            SAPbouiCOM.MenuCreationParams oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SAPbouiCOM.Framework.Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            SAPbouiCOM.MenuItem oMenuItem = SAPbouiCOM.Framework.Application.SBO_Application.Menus.Item("43520"); // moudles'

            SAPbouiCOM.Menus oMenus = oMenuItem.SubMenus;
         
            try
            {
                oMenuItem = SAPbouiCOM.Framework.Application.SBO_Application.Menus.Item("3328");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "IMPORTAÇÃO DE DADOS";
                oCreationPackage.String = "IMPORTAÇÃO DE DADOS";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            { // Menu already exists

                Application.SBO_Application.MessageBox(ex.Message);
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

        }

        public static void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                //get the clicked menu to get the corect menu its necessary to go SAP and use their tool
                // and find the module you wanto to catch.
                if (pVal.BeforeAction)
                {
                }
                else
                {
                    switch (pVal.MenuUID)
                    {
                        case "IMPORTAÇÃO DE DADOS":
                            {
                                new ImportacaoDeDados();
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
               SAPbouiCOM.Framework.Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

        public static void RemoveMenus()
        {
            if (SAPbouiCOM.Framework.Application.SBO_Application.Menus.Exists("IMPORTAÇÃO DE DADOS"))
                SAPbouiCOM.Framework.Application.SBO_Application.Menus.RemoveEx("IMPORTAÇÃO DE DADOS");

        }
    }
}
