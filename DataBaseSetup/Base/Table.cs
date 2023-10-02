using System;
using System.Runtime.InteropServices;

namespace Addon_Messem.DataBaseSetup.Base
{
    public class Table
    {
        private readonly SAPbobsCOM.Company company;
        private readonly Action<string> action;

        public Table(SAPbobsCOM.Company company, Action<String> action)
        {
            this.company = company;
            this.action = action;
        }

        /// <summary>
        /// Creates a table if teh table does not exist.
        /// </summary>
        /// <param name="tabName"> t]Tabel name without "@"</param>
        /// <param name="tabDescription"> Table Description </param>
        /// <param name="tabType"> Table Type. </param>
        public void CreateTable(string tabName, string tabDescription, SAPbobsCOM.BoUTBTableType tabType)
        {
            SAPbobsCOM.UserTablesMD userTablesMD = (SAPbobsCOM.UserTablesMD)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

            try
            {
                if (TableAlreadyExist(tabName))
                {
                    action($"Tabela: [{tabName}] já cadastrada");
                    return;
                }

                userTablesMD.TableName = tabName;
                userTablesMD.TableDescription = tabDescription;
                userTablesMD.TableType = tabType;

                if (userTablesMD.Add() != 0)
                {
                    company.GetLastError(out int errorCode, out string error);
                    throw new Exception($"Erro ao Gerar Tabela: {tabName} erro : [{errorCode} - {error}]");
                }
                else
                {
                    action($"Sucesso ao Gerar Tabela: [{tabName}]");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao Gerar Tabela: {ex.Message}");
            }
            finally
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    Marshal.FinalReleaseComObject(userTablesMD);

                GC.Collect();
            }
        }

        /// <summary>
        /// verify if the table already exist
        /// </summary>
        /// <param name="table"> nome da tabela. </param>
        /// <returns></returns>
        private bool TableAlreadyExist(string table)
        {
            SAPbobsCOM.Recordset record;
            record = (SAPbobsCOM.Recordset)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = $"SELECT * FROM OUTB WHERE \"TableName\" = '{table}'";
                record.DoQuery(query);

                return record.RecordCount > 0;
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao Procurar Tabela: {ex.Message}");
            }
            finally
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    Marshal.FinalReleaseComObject(record);

                GC.Collect();
            }
        }
    }
}
