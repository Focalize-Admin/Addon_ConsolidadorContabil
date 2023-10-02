using Addon_Messem.FORMS.Recursos;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Addon_Messem.FORMS.UserForms
{
    internal class FormErros
    {
        private readonly SAPbouiCOM.Form form;
        private readonly string formid;

        public FormErros(List<Erro> erros)
        {
            string xmlForm = Formulários.Erros.ToString();
            try
            {
                FormCreationParams cp = ((FormCreationParams)(SAPbouiCOM.Framework.Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_FormCreationParams)));
                cp.FormType = "ERROS";
                cp.XmlData = xmlForm;
                cp.UniqueID = "ERROS";
                form = SAPbouiCOM.Framework.Application.SBO_Application.Forms.AddEx(cp);
                formid = form.UniqueID;
                CustomInitialize(erros);
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

        private void CustomInitialize(List<Erro> erros)
        {
            PreencheMatrix(erros);
        }
        private void PreencheMatrix(List<Erro> erros)
        {
            try
            {
                DataTable dt = form.DataSources.DataTables.Item("DT_0");
                Matrix matrix = (Matrix)form.Items.Item("Item_0").Specific;
                for (int i = 0; i < erros.Count; i++)
                {

                    matrix.AddRow();
                    ((EditText)matrix.Columns.Item("Col_0").Cells.Item(i + 1).Specific).Value = erros[i].Rotina.ToString();
                    ((EditText)matrix.Columns.Item("Col_1").Cells.Item(i + 1).Specific).Value = erros[i].ID.ToString();
                    ((EditText)matrix.Columns.Item("Col_2").Cells.Item(i + 1).Specific).Value = erros[i].Banco.ToString();
                    ((EditText)matrix.Columns.Item("Col_3").Cells.Item(i + 1).Specific).Value = erros[i].Error.ToString();

                }
                for (int index = 0; index < matrix.Columns.Count; index++)
                    matrix.Columns.Item(index).Editable = false;
                matrix.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Erro ao preencher Matrix: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
    }
}
