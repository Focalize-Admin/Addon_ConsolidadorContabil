using System;
using System.Runtime.InteropServices;

namespace Addon_Messem.DataBaseSetup.Base
{
    /// <summary>
    /// Reference Class for the creation of an UDO in the SAP b1
    /// </summary>
    internal class UDO
    {
        private SAPbobsCOM.Company company;
        private Action<string> action;

        /// <summary>
        /// Initialize the object
        /// </summary>
        /// <param name="company"> the company object to acces the di api </param>
        /// <param name="application"> the ui aplication. </param>
        public UDO(SAPbobsCOM.Company company, Action<string> action)
        {
            this.company = company;
            this.action = action;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="UDO"> UDO ID </param>
        /// <param name="name"> Udo Name </param>
        /// <param name="tabName"> the master data table connected to the UDO. </param>
        /// <param name="type"> the table type </param>
        /// <param name="managedSeries"> is managed ? </param>
        /// <param name="delete"> the object can be deleted </param>
        /// <param name="canClose"> if a document can it be closed. </param>
        /// <param name="canCancel"> if a document can it be canceled. </param>
        /// <param name="canFind"> has a find implemented </param>
        /// <param name="canYearTransfer"> can generate a tranfer </param>
        /// <param name="hasDefaultForm"> has a default form created by the B1 Client </param>
        /// <param name="Campos"> the fields </param>
        /// <param name="enchancedForm"> has an enchanced default form </param>
        /// <param name="log"> generate log. </param>
        /// <param name="archive"> is an archive </param>
        /// <param name="defaultForm"> the default form string i never used but i belive is a xml string like the api uses to create the form during runtime</param> 
        public void SetupUDO(string UDO, string name, string tabName, SAPbobsCOM.BoUDOObjType type, SAPbobsCOM.BoYesNoEnum managedSeries, SAPbobsCOM.BoYesNoEnum delete, SAPbobsCOM.BoYesNoEnum canClose
            , SAPbobsCOM.BoYesNoEnum canCancel, SAPbobsCOM.BoYesNoEnum canFind, SAPbobsCOM.BoYesNoEnum canYearTransfer, SAPbobsCOM.BoYesNoEnum hasDefaultForm, string[] Campos, SAPbobsCOM.BoYesNoEnum enchancedForm, SAPbobsCOM.BoYesNoEnum log,
            SAPbobsCOM.BoYesNoEnum archive, string defaultForm)
        {

            SAPbobsCOM.UserObjectsMD userObjectMD = (SAPbobsCOM.UserObjectsMD)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

            try
            {
                if (this.UDOAlreadyExist(UDO))
                {
                    action($"UDO: [{UDO}] já cadastrado");
                    return;
                }

                userObjectMD.Code = UDO;
                userObjectMD.Name = name;
                userObjectMD.TableName = tabName;
                userObjectMD.ObjectType = type;
                userObjectMD.ManageSeries = managedSeries;
                userObjectMD.CanDelete = delete;
                userObjectMD.CanClose = canClose;
                userObjectMD.CanCancel = canCancel;
                userObjectMD.CanFind = canFind;

                if (canFind == SAPbobsCOM.BoYesNoEnum.tYES)
                {
                    if (type == SAPbobsCOM.BoUDOObjType.boud_MasterData)
                    {
                        userObjectMD.FindColumns.ColumnAlias = "Code";
                        userObjectMD.FindColumns.ColumnDescription = "Code";
                        userObjectMD.FindColumns.Add();
                        userObjectMD.FindColumns.ColumnAlias = "Name";
                        userObjectMD.FindColumns.ColumnDescription = "Name";
                    }
                    else
                    {
                        userObjectMD.FindColumns.ColumnAlias = "DocEntry";
                        userObjectMD.FindColumns.ColumnDescription = "Código";
                    }
                }

                userObjectMD.CanYearTransfer = canYearTransfer;

                userObjectMD.CanCreateDefaultForm = hasDefaultForm;

                if (hasDefaultForm == SAPbobsCOM.BoYesNoEnum.tYES)
                {
                    int nextLine = 1;

                    if (type == SAPbobsCOM.BoUDOObjType.boud_MasterData)
                    {
                        userObjectMD.FormColumns.FormColumnAlias = "Code";
                        userObjectMD.FormColumns.FormColumnDescription = "Code";
                        userObjectMD.FormColumns.Add();
                        userObjectMD.FormColumns.FormColumnAlias = "Name";
                        userObjectMD.FormColumns.FormColumnDescription = "Name";

                        nextLine = 2;
                    }
                    else
                    {
                        userObjectMD.FormColumns.FormColumnAlias = "DocEntry";
                        userObjectMD.FormColumns.FormColumnDescription = "Código";
                    }

                    // campos de usuário
                    if (Campos != null)
                    {
                        foreach (string campo in Campos)
                        {
                            userObjectMD.FormColumns.Add();
                            userObjectMD.FormColumns.SetCurrentLine(nextLine);
                            userObjectMD.FormColumns.FormColumnAlias = "U_" + campo;
                            userObjectMD.FormColumns.FormColumnDescription = campo;
                            userObjectMD.FormColumns.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                            userObjectMD.FormColumns.SonNumber = 0;
                            nextLine++;
                        }
                    }

                    userObjectMD.EnableEnhancedForm = enchancedForm;
                }

                userObjectMD.CanLog = log;
                userObjectMD.CanArchive = archive;

                if (!string.IsNullOrEmpty(defaultForm))
                {
                    userObjectMD.RebuildEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO;
                    userObjectMD.FormSRF = defaultForm;
                }

                if (userObjectMD.Add() != 0)
                {
                    int errorCode = 0;
                    string error = string.Empty;
                    company.GetLastError(out errorCode, out error);
                    action($"Erro ao Gerar Tabela: [{tabName}] erro : [{errorCode} - {error}]");
                }
                else
                {
                    action($"Sucesso ao Gerar UDO: [{UDO}]");
                }

            }
            catch (Exception ex)
            {
                action($"Erro ao Gerar UDO: {ex.Message}");
            }
            finally
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    Marshal.FinalReleaseComObject(userObjectMD);
                GC.Collect(); // call to release memory
            }
        }

        /// <summary>
        /// adds the chields tables like the lines of a document
        /// </summary>
        /// <param name="IdUdo"> the UDO parent ID</param>
        /// <param name="childTable"> the child table of the UDO main table like the lines of a document. </param>
        public void AddUDOChild(string IdUdo, string childTable)
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD = (SAPbobsCOM.UserObjectsMD)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

            if (oUserObjectMD.GetByKey(IdUdo))
            {
                bool childAlreadyExists = false;
                for (int i = 0; i < oUserObjectMD.ChildTables.Count; i++)
                {
                    oUserObjectMD.ChildTables.SetCurrentLine(i);
                    if (oUserObjectMD.ChildTables.TableName.Equals(childTable))
                    {
                        childAlreadyExists = true;
                        break;
                    }
                }

                if (!childAlreadyExists)
                {
                    if (oUserObjectMD.ChildTables.Count > 0)
                        oUserObjectMD.ChildTables.Add();

                    oUserObjectMD.ChildTables.SetCurrentLine(oUserObjectMD.ChildTables.Count - 1);
                    oUserObjectMD.ChildTables.TableName = childTable;
                    oUserObjectMD.ChildTables.ObjectName = childTable;

                    int errorCode = oUserObjectMD.Update();
                    // check for errors in the process
                    if (errorCode != 0)
                    {
                        if (errorCode != -2035)
                        {
                            action($"Erro ao Gerar UDO: {company.GetLastErrorDescription()}");
                        }
                    }
                    else
                    {
                        action($"Sucesso ao Adicionar tabela Filha [{childTable}] para o UDO: [{IdUdo}]");
                    }
                }
                else
                {
                    action($"Filho [{childTable}] já Cadastrado ao UDO [{IdUdo}]");
                }

                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    Marshal.FinalReleaseComObject(oUserObjectMD);
                GC.Collect(); // Release the handle to the table
            }
        }

        /// <summary>
        /// A quick select to validade the UDO so it dosent generate errors in a try catch since teh UDO already Exists.
        /// </summary>
        /// <param name="UDOName"> the Udo ID </param>
        /// <returns> true if exist </returns>
        public bool UDOAlreadyExist(string UDOName)
        {
            SAPbobsCOM.Recordset record = (SAPbobsCOM.Recordset)this.company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = $"SELECT * FROM OUDO WHERE \"Code\" = '{UDOName}'";
                record.DoQuery(query);

                return record.RecordCount > 0;
            }
            catch (Exception ex)
            {
                action($"Erro ao Procurar UDO: {ex.Message}");
                return false;
            }
            finally
            {
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    Marshal.FinalReleaseComObject(record);
                GC.Collect(); // call to release memory
            }
        }
    }
}
