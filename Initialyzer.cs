using Addon_Messem.DIAPI;
using Addon_Messem;
using SAPbobsCOM;
using System;
using System.Windows.Forms;

namespace Addon_Messem
{
    internal static class Initialyzer
    {
        public static void Start(string[] args)
        {
            try
            {
                SAPbouiCOM.Framework.Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new SAPbouiCOM.Framework.Application();
                }
                else
                {
                    oApp = new SAPbouiCOM.Framework.Application(args[0]);
                }

                // pass the company object to the class di api control.
                API.Company = (Company)SAPbouiCOM.Framework.Application.SBO_Application.Company.GetDICompany();
                // check the Data Base and creates the fields and tables if necessary.
                Menu.RemoveMenus();
                Menu.AddMenuItems();// add the menus.
                oApp.RegisterMenuEventHandler(Menu.SBO_Application_MenuEvent);
                SAPbouiCOM.Framework.Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Inicialização Concluida", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                Application.Exit();
            }
        }


        /// <summary>
        /// events of the SAP client.
        /// </summary>
        /// <param name="EventType"> SAPclient events </param>
        private static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    {
                        //Exit Add_on
                        Menu.RemoveMenus();
                        Application.Exit();
                    }
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                        Menu.RemoveMenus();
                        Menu.AddMenuItems();// add the menus.
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }

    }
}
