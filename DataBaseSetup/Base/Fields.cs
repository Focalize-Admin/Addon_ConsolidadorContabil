using Addon_Messem.DataBaseSetup.Models;
using System.Runtime.InteropServices;
using System;


namespace Addon_Messem.DataBaseSetup.Base
{
    public class Fields
    {
        private SAPbobsCOM.Company company;
        private Action<string> action;

        public Fields(SAPbobsCOM.Company company, Action<string> action)
        {
            this.company = company;
            this.action = action;
        }

        /// <summary>
        /// creates the field in the chosen table.
        /// </summary>
        /// <param name="table"> table where field will be created. </param>
        /// <param name="field"> field Name </param>
        /// <param name="fieldDescription"> Field Description. </param>
        /// <param name="fieldType"> Field Type. </param>
        /// <param name="fieldSize"> Field Size. </param>
        /// <param name="defaultValue"> Default value foe the field. </param>
        /// <param name="subType"> Sub Type. </param>
        /// <param name="validvalues"> Valid Values. </param>
        /// <param name="systemTable"> Is System Table. </param>
        public void CreateFields(string table, string field, string fieldDescription, SAPbobsCOM.BoFieldTypes fieldType,
            int fieldSize, string defaultValue, SAPbobsCOM.BoFldSubTypes subType = SAPbobsCOM.BoFldSubTypes.st_None, ValidValue[]? validvalues = null, bool systemTable = false, string linkedTable = "", LinkType type = LinkType.User_Table)
        {
            // objeto de campos de usuário.
            SAPbobsCOM.UserFieldsMD userFieldsMD = (SAPbobsCOM.UserFieldsMD)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            try
            {
                if (this.IsFieldCreated(field, table, systemTable) != -1)
                {
                    action($"Campo: [{field}] já cadastrado a Tabela: [{table}]");
                    return;
                }

                userFieldsMD.TableName = table;
                userFieldsMD.Name = field;
                userFieldsMD.Description = fieldDescription;

                userFieldsMD.Type = fieldType;
                userFieldsMD.SubType = subType;
                if(!string.IsNullOrEmpty(linkedTable))
                {
                    switch (type)
                    {
                        case LinkType.User_Table:
                            userFieldsMD.LinkedTable = linkedTable;
                            break;
                        case LinkType.UDO_Table:
                            userFieldsMD.LinkedUDO = linkedTable;
                            break;
                    }
                }
                //if (subType != SAPbobsCOM.BoFldSubTypes.st_None)
                //    userFieldsMD.SubType = subType;

                if (fieldType != SAPbobsCOM.BoFieldTypes.db_Date && fieldType != SAPbobsCOM.BoFieldTypes.db_Memo && fieldType != SAPbobsCOM.BoFieldTypes.db_Float)
                    userFieldsMD.EditSize = fieldSize;

                if (validvalues != null)
                {
                    foreach (ValidValue validValue in validvalues)
                    {
                        userFieldsMD.ValidValues.Value = validValue.Value;
                        userFieldsMD.ValidValues.Description = validValue.Description;
                        userFieldsMD.ValidValues.Add();
                    }
                }

                if (!string.IsNullOrEmpty(defaultValue))
                    userFieldsMD.DefaultValue = defaultValue;

                if (userFieldsMD.Add() != 0)
                {
                    int errorCode = 0;
                    string error = string.Empty;
                    company.GetLastError(out errorCode, out error);

                    action($"Erro ao Gerar campo : [{errorCode} - {error}]");
                }
                else
                {
                    action($"Campo: [{field}] cadastrado com sucesso a Tabela: [{table}]");
                }
            }
            finally
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    Marshal.FinalReleaseComObject(userFieldsMD);

                GC.Collect(); // call to release memory
            }
        }

        public void UpdateValidValues(string table, string field, ValidValue[] validvalues, bool systemTable)
        {
            // objeto de campos de usuário.
            SAPbobsCOM.UserFieldsMD userFieldsMD = (SAPbobsCOM.UserFieldsMD)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            try
            {
                int fieldID = this.IsFieldCreated(field, table, systemTable);
                if (fieldID == -1)
                    return;

                string arroba = "";
                if (!systemTable)
                    arroba = "@";

                userFieldsMD.GetByKey(arroba + table, fieldID);

                //deleta os valores validos
                for (int row = userFieldsMD.ValidValues.Count - 1; row >= 0; row--)
                {
                    userFieldsMD.ValidValues.SetCurrentLine(row);
                    userFieldsMD.ValidValues.Delete();
                }

                if (userFieldsMD.ValidValues.Count < 1)
                {
                    userFieldsMD.ValidValues.Add();
                }

                foreach (ValidValue validValue in validvalues)
                {
                    userFieldsMD.ValidValues.Value = validValue.Value;
                    userFieldsMD.ValidValues.Description = validValue.Description;
                    userFieldsMD.ValidValues.Add();
                }

                if (userFieldsMD.Update() != 0)
                {
                    int errorCode = 0;
                    string error = string.Empty;
                    company.GetLastError(out errorCode, out error);

                    action($"Erro ao fazer upgrade do campo : [{errorCode} - {error}]");
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                action($"Erro ao pesquisar existencia de campo: {ex.Message}");
            }
            finally
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    Marshal.FinalReleaseComObject(userFieldsMD);

                GC.Collect(); // call to release memory
            }
        }

        /// <summary>
        /// verify if the field already exists
        /// </summary>
        /// <param name="FieldName"> Field Name. </param>
        /// <param name="table"> table. </param>
        /// <param name="sistemTable"> true if System Table. </param>
        /// <returns></returns>
        private int IsFieldCreated(string FieldName, string table, bool sistemTable)
        {
            SAPbobsCOM.Recordset record = (SAPbobsCOM.Recordset)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {

                string arroba = sistemTable ? "" : "@";
                string query = $"SELECT \"FieldID\" FROM CUFD WHERE \"AliasID\" = '{FieldName}' AND \"TableID\" = '{arroba + table}'";
                record.DoQuery(query);
                if (record.RecordCount > 0)
                    return Convert.ToInt32(record.Fields.Item("FieldID").Value.ToString());

                return -1;
            }
            catch (Exception ex)
            {
                action($"Erro ao pesquisar existencia de campo: {ex.Message}");
                return -1;
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
