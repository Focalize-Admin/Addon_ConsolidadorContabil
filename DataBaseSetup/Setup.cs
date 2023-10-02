using SAPbouiCOM;
using Addon_Messem.DIAPI;
using SAPbobsCOM;
using Addon_Messem.DataBaseSetup.Base;
using System;
using Addon_Messem.DataBaseSetup.Models;

namespace Addon_Messem.DataBaseSetup
{
    internal static class SetUp
    {
        public static bool StartSetUp()
        {
            try
            {
                Action<string> action = delegate(string value)
                {
                    Console.WriteLine(value);
                };

                //SetupTables(action, API.Company!);
                SetupFields(API.Company!, action);
                return true;
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        /// <summary>
        /// Creates all tables for the add-on
        /// </summary>
        /// <param name="company"> DI Company </param>
        private static void SetupTables(Action<string> action, SAPbobsCOM.Company company)
        {
            //User Tables
            Table tables = new Table(company, action);
            //
            tables.CreateTable("DUO_PARAMCON", "Consolidador - Parametros", BoUTBTableType.bott_NoObject);
        }

        /// <summary>
        /// Set the tables fields.
        /// </summary>
        /// <param name="company"> DI Company </param>
        private static void SetupFields(SAPbobsCOM.Company company, Action<string> action)
        {
            Base.Fields fields = new Base.Fields(company, action);


            fields.CreateFields("DUO_PARAMCON", "ID", "ID Empresa", BoFieldTypes.db_Numeric, 6, String.Empty) ;
            fields.CreateFields("DUO_PARAMCON", "DB", "Banco de Dados Origem", BoFieldTypes.db_Alpha, 100, String.Empty);
            fields.CreateFields("OJDT", "IDDB", "ID Empresa", BoFieldTypes.db_Numeric, 10, String.Empty);
            fields.CreateFields("OJDT", "IDTRAN", "ID Transacao Origem", BoFieldTypes.db_Numeric, 10, String.Empty);
        }
    }
}