using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Addon_Messem.FORMS.Recursos;
using SAPbobsCOM;
using SAPbouiCOM;
using CheckBox = SAPbouiCOM.CheckBox;

namespace Addon_Messem.FORMS.UserForms
{
    internal class ImportacaoDeDados
    {
        private readonly SAPbouiCOM.Form form;
        private readonly string formid;

        private List<Erro> erros = new List<Erro>();


        public ImportacaoDeDados()
        {
            string xmlForm = Formulários.ImportaçãoDados.ToString();
            try
            {
                FormCreationParams cp = ((FormCreationParams)(SAPbouiCOM.Framework.Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams)));
                cp.FormType = "IMPORTAÇÃO DE DADOS";
                cp.XmlData = xmlForm;
                cp.Modality = BoFormModality.fm_None;
                cp.UniqueID = "IMPORTAÇÃO DE DADOS";
                form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.AddEx(cp);
                formid = form.UniqueID;
                CustomInitialize();
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao crair o formulário: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (form != null)
                {
                    form.Visible = true;
                }
            }
        }

        private void CustomInitialize()
        {
            SAPbouiCOM.Framework.Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
            FormLoad();
        }

        private void SBO_Application_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (FormUID != formid)
                return;
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    case BoEventTypes.et_CLICK:
                        {
                            switch (pVal.ItemUID)
                            {
                                case "Item_12":
                                    {
                                        IniciarImportanção();
                                    }
                                    break;
                                case "Item_13":
                                    {
                                        FormClose();
                                    }
                                    break;
                            }
                        }
                        break;

                }

            }
        }

        private void FormClose()
        {
            try
            {
                form.Close();
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.MessageBox($"Erro ao fechar formulário! : {ex.Message}");
            }
        }
        
        
          

        private void IniciarImportanção()
        {
            form.Freeze(true);
            try
            {
                erros.Clear();
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(@$"Iniciando importação... Aguarde!", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);

                Recordset record = (Recordset)DIAPI.API.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                record.DoQuery(Queries.EmpresasConsolidadoras);
                for (int i = 0; i < record.RecordCount; i++)
                { 
                    Recordset record1 = (Recordset)DIAPI.API.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    record1.DoQuery($@"Select * from ""{record.Fields.Item("U_BD").Value}"".dbo.""OACT"" where ""Levels"" > 1 And ""AcctCode"" Not In (Select ""AcctCode"" from oact)");
                    Recordset record2 = (Recordset)DIAPI.API.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    record2.DoQuery(@$"select * from ""{record.Fields.Item("U_BD").Value}"".dbo.""OPRC"" where ""DimCode"" = 1 and ""PrcCode"" not in (Select ""PrcCode"" from oprc)");
                    Recordset record3 = (Recordset)DIAPI.API.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    string query = @$"select * from ""{record.Fields.Item("U_BD").Value}"".dbo.""OJDT"" a where a.""TransId"" not in (select isnull (""U_IDTRAN"", 0) from ""OJDT"" where cast(""U_IDBD"" as int) = cast({record.Fields.Item("U_ID").Value} as int))";
                    record3.DoQuery(@$"select * from ""{record.Fields.Item("U_BD").Value}"".dbo.""OJDT"" a where a.""TransId"" not in (select isnull (""U_IDTRAN"", 0) from ""OJDT"" where cast(""U_IDBD"" as int) = cast({record.Fields.Item("U_ID").Value} as int))");
                    bool checkPC = ((CheckBox)form.Items.Item("Item_3").Specific).Checked;
                    bool checkCC = ((CheckBox)form.Items.Item("Item_4").Specific).Checked;
                    bool checkLC = ((CheckBox)form.Items.Item("Item_5").Specific).Checked;


                    //Plano de Contas
                    if (checkPC)
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(@$"{record.Fields.Item("U_BD").Value} Lendo plano de contas... Aguarde!", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                       
                        while (!record1.EoF)
                        {
                            CriaPC(record1.Fields.Item("acctcode").Value.ToString(), record1.Fields.Item("acctname").Value.ToString(),
                                record1.Fields.Item("accntntcod").Value.ToString(), record1.Fields.Item("LocManTran").Value.ToString(),
                                record1.Fields.Item("acttype").Value.ToString(), record1.Fields.Item("FatherNum").Value.ToString(),
                                record1.Fields.Item("Postable").Value.ToString(), record.Fields.Item("U_BD").Value.ToString());

                            record1.MoveNext();
                        }
                    }

                    //Centro de Custos
                    if (checkCC)
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(@$"{record.Fields.Item("U_BD").Value} Lendo centro de custos... Aguarde!", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);
                       
                        while (!record2.EoF)
                        {

                            CriaCC(record2.Fields.Item("PrcCode").Value.ToString(), record2.Fields.Item("PrcName").Value.ToString(),
                                record2.Fields.Item("GrpCode").Value.ToString(), (DateTime)record2.Fields.Item("ValidFrom").Value,
                                (DateTime)record2.Fields.Item("ValidTo").Value, record.Fields.Item("U_BD").Value.ToString());

                            record2.MoveNext();
                        }
                    }

                    //Lançamentos Contábeis
                    if (checkLC)
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(@$"{record.Fields.Item("U_BD").Value} Lendo lançamento Contábeis... Aguarde!", BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Warning);

                        while (!record3.EoF)
                        {

                            CriaLC((DateTime)record3.Fields.Item("refdate").Value, (DateTime)record3.Fields.Item("duedate").Value,
                                (DateTime)record3.Fields.Item("taxdate").Value, record3.Fields.Item("memo").Value.ToString(),
                                record3.Fields.Item("transid").Value.ToString(), record.Fields.Item("U_ID").Value.ToString(),
                                record.Fields.Item("U_BD").Value.ToString());

                            record3.MoveNext();
                        }
                    }
                    record.MoveNext();
                }
                _ = new FormErros(erros);
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao Gerar Documentos! : {ex.Message}");
            }
            finally
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("Importação concluída!", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Success);
                form.Freeze(false);
            }

        }

        private void FormLoad()
        {
            try
            {
                Recordset record = (Recordset)DIAPI.API.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                record.DoQuery(Queries.BuscaItens_Empresa);
                Recordset record1 = (Recordset)DIAPI.API.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                record1.DoQuery(Queries.BuscaVendas_Empresa);
                Recordset record2 = (Recordset)DIAPI.API.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                record2.DoQuery(Queries.EmpresasConsolidadoras);
                Grid grid = (Grid)form.Items.Item("Item_2").Specific;
                DataTable dt = form.DataSources.DataTables.Item("DT_0");

                if (record.RecordCount > 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("");
                    return;
                }
                if (record1.RecordCount > 0)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText("");
                    return;
                }
                if (record2.RecordCount > 0)
                {
                    dt.ExecuteQuery(Queries.EmpresasConsolidadoras);
                }
                else
                {

                }
                for (int index = 0; index < grid.Columns.Count; index++)
                    grid.Columns.Item(index).Editable = false;
                grid.Columns.Item("Name").Visible = false;
                grid.Columns.Item("U_ID").Visible = false;


                grid.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao carregar form! : {ex.Message}");
            }
        }

        private void CriaPC(string codigo, string nome, string codigoexterno, string is_control, string tipo_conta, string conta_pai, string is_title, string banco)
        {
            try
            {

                ChartOfAccounts pc = (ChartOfAccounts)DIAPI.API.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts);
                pc.Code = codigo;
                pc.Name = nome;
                pc.ExternalCode = codigoexterno;

                // is_control
                if (is_control == "Y")
                    pc.LockManualTransaction = SAPbobsCOM.BoYesNoEnum.tYES;
                else
                    pc.LockManualTransaction = SAPbobsCOM.BoYesNoEnum.tNO;

                // tipo_conta
                if (tipo_conta == "E")
                    pc.AccountType = SAPbobsCOM.BoAccountTypes.at_Expenses;
                else if (tipo_conta == "N")
                    pc.AccountType = SAPbobsCOM.BoAccountTypes.at_Other;
                else
                    pc.AccountType = SAPbobsCOM.BoAccountTypes.at_Revenues;

                pc.FatherAccountKey = conta_pai;

                // it_title
                if (is_title == "Y")
                    pc.ActiveAccount = SAPbobsCOM.BoYesNoEnum.tYES;
                else
                    pc.ActiveAccount = SAPbobsCOM.BoYesNoEnum.tNO;
                if (pc.Add() != 0)
                {
                    erros.Add(new Erro("Plano de Contas", codigo, banco, DIAPI.API.Company.GetLastErrorDescription()));
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao pesquisar dados! : {ex.Message}");

            }
        }

        private void CriaCC(string code, string name, string grp, DateTime validFrom, DateTime validTo, string banco)
        {
            try
            {
                CompanyService companyService = (CompanyService)DIAPI.API.Company.GetCompanyService();
                ProfitCentersService service = (ProfitCentersService)companyService.GetBusinessService(ServiceTypes.ProfitCentersService);
                ProfitCenter profitCenter = (ProfitCenter)service.GetDataInterface(ProfitCentersServiceDataInterfaces.pcsProfitCenter);
                profitCenter.CenterCode = code;
                profitCenter.CenterName = name;
                profitCenter.GroupCode = grp;
                profitCenter.InWhichDimension = 1;
                profitCenter.Effectivefrom = validFrom;
                profitCenter.EffectiveTo = validTo;
                service.AddProfitCenter(profitCenter);
            }
            catch (Exception ex)
            {
                erros.Add(new Erro("Centro de Custos", code, banco, ex.Message));

            }
        }
        private void CriaLC(DateTime date, DateTime datavcto, DateTime dataDoc, string memo, string transid, string idBase, string banco)
        {
            try
            {

                JournalEntries lc = (JournalEntries)DIAPI.API.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                int x = 0;

                lc.ReferenceDate = date;
                lc.DueDate = datavcto;
                lc.TaxDate = dataDoc;
                lc.Memo = memo;
                lc.UserFields.Fields.Item("U_IDTRAN").Value = transid;
                lc.UserFields.Fields.Item("U_IDBD").Value = idBase;
                Recordset record = (Recordset)DIAPI.API.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                record.DoQuery($@"select * from ""{banco}"".dbo.""JDT1"" where ""TransId"" = " + transid);
                for (int index = 0; index < record.RecordCount; index++)
                {
                    if (x > 0)
                    {
                        lc.Lines.Add();
                        lc.Lines.SetCurrentLine(x);
                    }
                    lc.Lines.AccountCode = record.Fields.Item("account").Value.ToString();
                    lc.Lines.ShortName = record.Fields.Item("shortname").Value.ToString();
                    lc.Lines.ContraAccount = record.Fields.Item("contraact").Value.ToString();
                    lc.Lines.Credit = Convert.ToDouble(record.Fields.Item("credit").Value);
                    lc.Lines.Debit = Convert.ToDouble(record.Fields.Item("debit").Value);
                    lc.Lines.LineMemo = record.Fields.Item("linememo").Value.ToString();
                    lc.Lines.Reference1 = record.Fields.Item("ref1").Value.ToString();
                    lc.Lines.Reference2 = record.Fields.Item("ref2").Value.ToString();
                    lc.Lines.AdditionalReference = record.Fields.Item("ref3Line").Value.ToString();
                    lc.Lines.CostingCode = record.Fields.Item("ProfitCode").Value.ToString();

                    x++;
                    record.MoveNext();
                }
                if (lc.Add() != 0)
                {
                    erros.Add(new Erro("Lançamentos Contábeis", transid, banco, DIAPI.API.Company.GetLastErrorDescription()));
                }
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao gerar Lançamento Contábel! : {ex.Message}");
            }
        }

    }
}
