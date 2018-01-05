using System;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Word = Microsoft.Office.Interop.Word;

using TDMS.Interop;

namespace WordAddIn
{
    public partial class TDMSControl
    {
        private string StringVariable;
        private string ScriptResult;       

        private void TDMS_Load(object sender, RibbonUIEventArgs e)
        {   
        }

        /// <summary>
        /// Метод для сохранения документа в ТДМС (флаг не снимается)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Application WordApp = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                Word.Document WordDoc = WordApp.Documents.Application.ActiveDocument;

                if (Path.GetDirectoryName(WordDoc.FullName) != string.Empty)
                {
                    if (Condition.CheckPathToTdms(WordDoc.FullName))
                    {
                        if (Condition.CheckTDMSProcess())
                        {
                            //Получение объекта по GUID
                            var tdmsObj = new TDMSApplication().GetObjectByGUID(Condition.ParseGUID(WordDoc.FullName));

                            //Заблокирован ли чертёж текущим пользователем?
                            if (tdmsObj.Permissions.LockOwner)
                            {
                                //Сохранение изменений в текущем документе
                                WordDoc.Save();

                                //Загрузка в базу ТДМС
                                tdmsObj.CheckIn();
                                tdmsObj.Update();
                            }
                            else
                            {
                                MessageBox.Show(@"\n Документ открыт на просмотр, изменения не будут сохранены в TDMS. Сохраните изменения локально.");
                            }
                        }
                        else
                        {
                            MessageBox.Show(@"Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                        }
                    }
                    else
                    {
                        WordDoc.Save();
                    }
                }
                else
                {  
                    WordDoc.SaveAs2(SaveDialog());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private string SaveDialog()
        {
            var sfd = new SaveFileDialog();
            return sfd.ShowDialog() == DialogResult.OK ? sfd.FileName : string.Empty;
        }

        /// <summary>
        /// Метод для сохранения и закрытия документа в ТДМС (флаг снимается)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSaveAndClose_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Application WordApp = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                Word.Document WordDoc = WordApp.Documents.Application.ActiveDocument;

                if (Path.GetDirectoryName(WordDoc.FullName) != string.Empty)
                {
                    if (Condition.CheckPathToTdms(WordDoc.FullName))
                    {
                        if (Condition.CheckTDMSProcess())
                        {
                            //Получение объекта по GUID
                            var tdmsObj = new TDMSApplication().GetObjectByGUID(Condition.ParseGUID(WordDoc.FullName));

                            //Заблокирован ли чертёж текущим пользователем?
                            if (tdmsObj.Permissions.LockOwner)
                            {
                                //Сохранение в базе и разблокировка
                                WordDoc.Save();
                                tdmsObj.UnlockCheckIn(1);
                                tdmsObj.Update();
                                WordDoc.Close();
                            }
                            else
                            {
                                MessageBox.Show(@"\n Документ открыт на просмотр, изменения не будут сохранены в TDMS. Сохраните изменения локально.");
                            }
                        }
                        else
                        {
                            MessageBox.Show(@"Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                        }
                    }
                    else
                    {
                        WordDoc.Save();
                        MessageBox.Show(@"Документ не принадлежит TDMS и сохранён локально");
                        WordDoc.Close();
                    }
                }
                else
                {
                    try
                    {
                        WordDoc.SaveAs2(SaveDialog());
                        WordDoc.Close();
                    }
                    catch
                    {
                        MessageBox.Show(@"\n Команда отменена \n");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Метод для закрытия документа в ТДМС (флаг снимается)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Application WordApp = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                Word.Document WordDoc = WordApp.Documents.Application.ActiveDocument;

                if (Condition.CheckPathToTdms(WordDoc.FullName))
                {
                    if (Condition.CheckTDMSProcess())
                    {
                        //Получение объекта по GUID
                        var tdmsObj = new TDMSApplication().GetObjectByGUID(Condition.ParseGUID(WordDoc.FullName));

                        //Заблокирован ли чертёж текущим пользователем?
                        if (tdmsObj.Permissions.LockOwner)
                        {
                            tdmsObj.UnlockCheckIn(0);
                            tdmsObj.Update();
                            WordDoc.Close();
                        }
                        else
                        {
                            WordDoc.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show(@"Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    }
                }
                else
                {
                    WordDoc.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnHelp_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"V:\_HELP\");
            }
            catch
            {
                MessageBox.Show(@"Путь к каталогу V:\\_HELP\\ в котором расположены справочные материалы не найден.", @"Ошибка пути");
            }
        }

        private void btnSTP_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"V:\_HELP\STANDART.pdf");
            }
            catch
            {
                MessageBox.Show(@"Путь к файлу V:\\_HELP\\STANDART.pdf не найден.", @"Ошибка пути");
            }
        }

        private void btnFAQ_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(@"V:\_HELP\Часто задаваемые вопросы.pdf");
            }
            catch
            {
                MessageBox.Show(@"Путь к файлу V:\\_HELP\\Часто задаваемые вопросы.pdf не найден.", @"Ошибка пути");
            }
        }

        //О программе
        private void btnOprogr_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Интерфейс разработан для ООО Архитектурное бюро Студия 44 в 2017 году.\nРазработчик: Kozhaev L.I.\nemail : kozhaelf@gmail.com", @"О программе.");
        }

        /// <summary>
        /// Происходит обновление полей атрибутов
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnUpdate_Click(object sender, RibbonControlEventArgs e)
        {
            UpdateTDMSVariables();
        }

        public void UpdateTDMSVariables()
        {
            try
            {
                if (Condition.CheckTDMSProcess())
                {

                    Word.Application wordApp = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                    Word.Document activeDocument = wordApp.Documents.Application.ActiveDocument;
                    
                    if (Condition.CheckPathToTdms(activeDocument.Path))
                    {
                        string str = Condition.ParseGUID(activeDocument.Path);
                        
                        var tdmsApp = (TDMSApplication)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("430C37CB-33C4-4754-836A-A3930689D437")));

                        TDMSUser user = null;
                        TDMSObject tdmsObj = tdmsApp.GetObjectByGUID(Condition.ParseGUID(activeDocument.Path));
                        TDMSAttributes parAttrs = tdmsObj.Attributes;
                        var wrdVars = activeDocument.Variables;

                        // ПСД
                        if (string.Equals(tdmsObj.ObjectDefName, "O_PSD", StringComparison.InvariantCultureIgnoreCase))
                        {
                            foreach (Word.Variable variable in wrdVars)
                            {
                                switch (variable.Name)
                                {
                                    case "F_GetOOOName":
                                        {
                                            string GetOOOName = (string)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetOOOName");
                                            if (!string.IsNullOrEmpty(GetOOOName))
                                            {
                                                activeDocument.Variables[variable.Name].Value = GetOOOName;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "F_GetOOOUrAddress":
                                        {
                                            string GetOOOUrAddress = (string)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetOOOUrAddress");
                                            if (!string.IsNullOrEmpty(GetOOOUrAddress))
                                            {
                                                activeDocument.Variables[variable.Name].Value = GetOOOUrAddress;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "F_GetOOOPhones":
                                        {
                                            string GetOOOPhones = (string)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetOOOPhones");
                                            if (!string.IsNullOrEmpty(GetOOOPhones))
                                            {
                                                activeDocument.Variables[variable.Name].Value = GetOOOPhones;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "A_CONTACT_REF":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(variable.Name))
                                        {
                                            if (!string.IsNullOrEmpty(tdmsObj.Attributes[variable.Name].Value))
                                            {
                                                activeDocument.Variables[variable.Name].Value = tdmsObj.Attributes[variable.Name].Value;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_OBOZN_DOC":
                                    case "A_NAME":
                                    case "A_ARCH_SIGN":
                                    case "A_TOM_PAGE_NUM":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(variable.Name))
                                        {
                                            if (!string.IsNullOrEmpty(tdmsObj.Attributes[variable.Name].Value))
                                            {
                                                activeDocument.Variables[variable.Name].Value = tdmsObj.Attributes[variable.Name].Value;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_STAGE":
                                    case "A_STAGE_CLSF":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(variable.Name))
                                        {
                                            if (!string.IsNullOrEmpty(tdmsObj.Attributes[variable.Name].Value))
                                            {
                                                activeDocument.Variables[variable.Name].Value = tdmsObj.Attributes[variable.Name].Classifier.Code;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_OBOZN":
                                    case "A_INSTEAD_OF_NUM":
                                    case "A_LIC_NUM_P":
                                    case "A_LIC_NUM_IZ":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(variable.Name))
                                        {
                                            if (!string.IsNullOrEmpty(parAttrs[variable.Name].Value))
                                            {
                                                activeDocument.Variables[variable.Name].Value = parAttrs[variable.Name].Value;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_YEAR":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(variable.Name))
                                        {
                                            if (!string.IsNullOrEmpty(Convert.ToString(parAttrs[variable.Name].Value)))
                                            {
                                                activeDocument.Variables[variable.Name].Value = Convert.ToString(parAttrs[variable.Name].Value);
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_TOM_NAME_ADD":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_STAGE"))
                                        {
                                            if (!(parAttrs["A_STAGE"].Classifier.Code == "Р"))
                                                return;
                                        }
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_STAGE_CLSF"))
                                        {
                                            if (!(parAttrs["A_STAGE_CLSF"].Classifier.Code == "Р"))
                                                return;
                                        }
                                        string mark_code = parAttrs["A_MARK"].Classifier.Code;
                                        string name_add = tdmsApp.ExecuteScript("O_PSD", "GetTomNameAddForTitul", mark_code);

                                        if (!string.IsNullOrEmpty(name_add))
                                        {
                                            activeDocument.Variables[variable.Name].Value = "Основной комплект рабочих чертежей" + Strings.Chr(13) + Strings.Chr(10) + name_add;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = "Основной комплект рабочих чертежей";
                                        }
                                        continue;
                                    case "A_TOM_NUM":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_TOM_NUMB"))
                                        {
                                            if (!string.IsNullOrEmpty(parAttrs["A_TOM_NUMB"].Value))
                                            {
                                                activeDocument.Variables[variable.Name].Value = parAttrs["A_TOM_NUMB"].Value;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_TOM_NAME":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_TOM_NUMB"))
                                        {
                                            string tom_num = parAttrs["A_TOM_NUMB"].Value;

                                            if (tom_num == "5")
                                            {
                                                activeDocument.Variables[variable.Name].Value = "Сведения об инженерном оборудовании, о сетях инженерно-технического обеспечения, перечень инженерно-технических мероприятий, содержание технических решений";
                                            }
                                            else if (tom_num == "12")
                                            {
                                                activeDocument.Variables[variable.Name].Value = "Иная документация в случаях, предусмотренных федеральными законами";
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = parAttrs["A_NAME"].Value;
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_SUB_TOM_NUM":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_SUBTOM_NUMB"))
                                        {
                                            if (!string.IsNullOrEmpty(parAttrs["A_SUBTOM_NUMB"].Value))
                                            {
                                                activeDocument.Variables[variable.Name].Value = "Подраздел " + parAttrs["A_SUBTOM_NUMB"].Value;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_SUB_TOM_NAME":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_TOM_NUMB"))
                                        {
                                            string tom_num = parAttrs["A_TOM_NUMB"].Value;

                                            if (tom_num == "5" | tom_num == "12")
                                            {
                                                activeDocument.Variables[variable.Name].Value = Strings.Chr(171) + parAttrs["A_NAME"].Value + Strings.Chr(187);
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_BOOK_NUM":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_BOOK_NUMB"))
                                        {
                                            if (!string.IsNullOrEmpty(parAttrs["A_BOOK_NUMB"].Value))
                                            {
                                                activeDocument.Variables[variable.Name].Value = parAttrs["A_BOOK_NUMB"].Value;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_ADDRESS":
                                        string op_address = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetOPAddress", tdmsObj);
                                        if (!string.IsNullOrEmpty(op_address))
                                        {
                                            activeDocument.Variables[variable.Name].Value = op_address;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_REF_PART_WORK":
                                        string op_name = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetOPName", tdmsObj);
                                        if (!string.IsNullOrEmpty(op_name))
                                        {
                                            activeDocument.Variables[variable.Name].Value = op_name;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_TEXT_FORM":
                                        string txt_form = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetTextForm", tdmsObj);
                                        if (!string.IsNullOrEmpty(txt_form))
                                        {
                                            activeDocument.Variables[variable.Name].Value = txt_form;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_TITLE":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_REF_CONTRACT"))
                                        {
                                            TDMSObject tdmsContr = tdmsObj.Attributes["A_REF_CONTRACT"].Object;
                                            if (!string.IsNullOrEmpty(tdmsContr.Attributes["A_COMMENT"].Value))
                                            {
                                                activeDocument.Variables[variable.Name].Value = Strings.UCase(tdmsContr.Attributes["A_COMMENT"].Value);
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            string title = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromStageFolder", tdmsObj, variable.Name);
                                            if (!string.IsNullOrEmpty(title))
                                            {
                                                activeDocument.Variables[variable.Name].Value = Strings.UCase(title);
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        continue;
                                    case "A_TITLE_ZAO":
                                        string title_zao = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromContract", tdmsObj, variable.Name);

                                        if (!string.IsNullOrEmpty(title_zao))
                                        {
                                            activeDocument.Variables[variable.Name].Value = title_zao;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_WRK":
                                    case "A_CHECKED":
                                    case "A_GL_SPEC":
                                    case "A_NORMK":
                                    case "A_DEPARTM_HEAD":
                                    case "A_GIP":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(variable.Name))
                                        {
                                            if (!string.IsNullOrEmpty(tdmsObj.Attributes[variable.Name].Value))
                                            {
                                                activeDocument.Variables[variable.Name].Value = tdmsObj.Attributes[variable.Name].User.LastName;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_ADOC_DATE_ISP":
                                    case "A_DOC_DATE_CHECK":
                                    case "A_GL_SPEC_DATE":
                                    case "A_NORMK_DATE":
                                    case "A_DEPARTM_HEAD_DATE":
                                    case "A_DATE_SIGN_GIP":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(variable.Name))
                                        {
                                            if (!string.IsNullOrEmpty(Convert.ToString(tdmsObj.Attributes[variable.Name].Value)))
                                            {
                                                string str_date = Convert.ToString(tdmsObj.Attributes[variable.Name].Value);
                                                activeDocument.Variables[variable.Name].Value = Strings.Mid(str_date, 4, 3) + Strings.Right(str_date, 2);
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_GIP_FIO":
                                    case "A_DEPARTM_HEAD_FIO":
                                    case "A_NORMK_FIO":
                                    case "A_GL_SPEC_FIO":
                                    case "A_CHECKED_FIO":
                                    case "A_WRK_FIO":
                                        StringVariable = Strings.Left(variable.Name, Strings.Len(variable.Name) - 4);
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(this.StringVariable))
                                        {
                                            if (!string.IsNullOrEmpty(tdmsObj.Attributes[this.StringVariable].Value))
                                            {
                                                user = tdmsObj.Attributes[this.StringVariable].User;
                                                activeDocument.Variables[variable.Name].Value = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;

                                    case "A_DEVELOP":
                                    case "A_CHECK":
                                    case "A_NORMKL":
                                    case "A_GR_HEAD":
                                        ScriptResult = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromActiveRouteTable", tdmsObj, Strings.Mid(variable.Name, 3), "A_User");
                                        if (!string.IsNullOrEmpty(ScriptResult))
                                        {
                                            activeDocument.Variables[variable.Name].Value = tdmsApp.Users[ScriptResult].LastName;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_GAP_":
                                    case "A_GIP_":
                                    case "A_GKP_":
                                        ScriptResult = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromStageFolder", tdmsObj, variable.Name);

                                        if (!string.IsNullOrEmpty(ScriptResult))
                                        {
                                            activeDocument.Variables[variable.Name].Value = tdmsApp.Users[ScriptResult].LastName;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_GKAB_":
                                        user = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromSysProps", variable.Name);

                                        if ((user != null))
                                        {
                                            activeDocument.Variables[variable.Name].Value = user.LastName;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_DATE_SIGN_DEVELOP":
                                    case "A_DATE_SIGN_CHECK":
                                    case "A_DATE_SIGN_NORMKL":
                                    case "A_DATE_SIGN_GR_HEAD":
                                    case "A_DATE_SIGN_GIP_":
                                    case "A_DATE_SIGN_GAP_":
                                    case "A_DATE_SIGN_GKP_":
                                    case "A_DATE_SIGN_GKAB_":
                                        string date_sign = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromActiveRouteTable", tdmsObj, Strings.Mid(variable.Name, 13), "A_DATE");
                                        if (!string.IsNullOrEmpty(date_sign))
                                        {
                                            activeDocument.Variables[variable.Name].Value = date_sign;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_GAP_FIO":
                                    case "A_GKP_FIO":
                                        StringVariable = Strings.Left(variable.Name, Strings.Len(variable.Name) - 3);
                                        ScriptResult = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromStageFolder", tdmsObj, this.StringVariable);

                                        if (!string.IsNullOrEmpty(ScriptResult))
                                        {
                                            user = tdmsApp.Users[ScriptResult];
                                            activeDocument.Variables[variable.Name].Value = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_GIP__FIO":
                                        this.StringVariable = Strings.Left(variable.Name, Strings.Len(variable.Name) - 4);

                                        ScriptResult = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromStageFolder", tdmsObj, this.StringVariable);

                                        if (!string.IsNullOrEmpty(ScriptResult))
                                        {
                                            user = tdmsApp.Users[ScriptResult];
                                            activeDocument.Variables[variable.Name].Value = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_GKAB_FIO":
                                        this.StringVariable = Strings.Left(variable.Name, Strings.Len(variable.Name) - 3);
                                        user = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromSysProps", this.StringVariable);
                                        if ((user != null))
                                        {
                                            activeDocument.Variables[variable.Name].Value = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_GR_HEAD_FIO":
                                        string sys_name = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromActiveRouteTable", tdmsObj, Strings.Mid(variable.Name, 3, 7), "A_User");

                                        if (!string.IsNullOrEmpty(sys_name))
                                        {
                                            user = tdmsApp.Users[sys_name];
                                            activeDocument.Variables[variable.Name].Value = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                }
                            }
                        }
                        else
                        {
                            tdmsApp.ExecuteScript("CMD_SYSLIB", "FillInSndngDocList", tdmsObj, activeDocument);
                            foreach (Word.Variable variable in wrdVars)
                            {
                                switch (variable.Name)
                                {
                                    case "F_GetDay":
                                        {
                                            short num = (short)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDay", (dynamic)tdmsObj.Attributes["A_DATE"].Value);
                                            if (!string.IsNullOrEmpty(num.ToString()))
                                            {
                                                activeDocument.Variables[variable.Name].Value = num.ToString();
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "F_GetMonth":
                                        {
                                            ScriptResult = (string)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetMonth", (dynamic)tdmsObj.Attributes["A_DATE"].Value);
                                            if (!string.IsNullOrEmpty(ScriptResult.ToString()))
                                            {
                                                activeDocument.Variables[variable.Name].Value = this.ScriptResult.ToString();
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "F_GetYear":
                                        {
                                            short num2 = (short)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetYear", (dynamic)tdmsObj.Attributes["A_DATE"].Value);
                                            if (!string.IsNullOrEmpty(num2.ToString()))
                                            {
                                                activeDocument.Variables[variable.Name].Value = num2.ToString();
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "F_GetOPName":
                                        {
                                            ScriptResult = (string)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetOPName", tdmsObj);
                                            if (!string.IsNullOrEmpty(this.ScriptResult))
                                            {
                                                activeDocument.Variables[variable.Name].Value = ScriptResult;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "F_GetContractNum":
                                        {
                                            ScriptResult = (string)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetContractNum", tdmsObj);
                                            if (!string.IsNullOrEmpty(this.ScriptResult))
                                            {
                                                activeDocument.Variables[variable.Name].Value = this.ScriptResult;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "F_GetContractDate":
                                        {
                                            ScriptResult = (string)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetContractDate", tdmsObj);
                                            if (!string.IsNullOrEmpty(this.ScriptResult.ToString()))
                                            {
                                                activeDocument.Variables[variable.Name].Value = this.ScriptResult.ToString();
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "F_GetSendingNum":
                                        {
                                            ScriptResult = (string)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetSendingNum", tdmsObj);
                                            if (!string.IsNullOrEmpty(this.ScriptResult))
                                            {
                                                activeDocument.Variables[variable.Name].Value = this.ScriptResult;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "F_GetCompanyName":
                                        {
                                            ScriptResult = (string)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetCompanyName", tdmsObj);
                                            if (!string.IsNullOrEmpty(ScriptResult))
                                            {
                                                activeDocument.Variables[variable.Name].Value = ScriptResult;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "F_GetDesObjShifr":
                                        {
                                            ScriptResult = (string)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDesObjShifr", tdmsObj);
                                            if (!string.IsNullOrEmpty(ScriptResult))
                                            {
                                                activeDocument.Variables[variable.Name].Value = this.ScriptResult;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "F_GetStageName":
                                        {
                                            ScriptResult = (string)tdmsApp.ExecuteScript("CMD_SYSLIB", "GetStageName", tdmsObj);
                                            if (!string.IsNullOrEmpty(ScriptResult))
                                            {
                                                activeDocument.Variables[variable.Name].Value = this.ScriptResult;
                                            }
                                            else
                                            {
                                                activeDocument.Variables[variable.Name].Value = " ";
                                            }
                                            continue;
                                        }
                                    case "A_User":
                                        if (!string.IsNullOrEmpty(tdmsObj.Attributes[variable.Name].Value))
                                        {
                                            activeDocument.Variables[variable.Name].Value = (string)tdmsObj.Attributes[variable.Name].Value;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_REG_NUM":
                                    case "A_CONTRACT_SHIFR":
                                    case "A_NAME_WORK":
                                        if (!string.IsNullOrEmpty(tdmsObj.Attributes[variable.Name].Value))
                                        {
                                            activeDocument.Variables[variable.Name].Value = tdmsObj.Attributes[variable.Name].Value;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_DEPART_TO":
                                    case "A_STAGE":
                                        if (!string.IsNullOrEmpty(tdmsObj.Attributes[variable.Name].Value))
                                        {
                                            activeDocument.Variables[variable.Name].Value = tdmsObj.Attributes[variable.Name].Classifier.Code;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_DEPART_FROM":
                                        if (!string.IsNullOrEmpty(tdmsObj.Attributes[variable.Name].Value))
                                        {
                                            activeDocument.Variables[variable.Name].Value = Strings.Left(tdmsObj.Attributes["A_DEPART_FROM"].Classifier.Code, 3);
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_GIP":
                                    case "A_DEPARTM_HEAD":
                                    case "A_GL_SPEC":
                                    case "A_GROUP_HEAD":
                                    case "A_WRK":
                                        if (!string.IsNullOrEmpty(tdmsObj.Attributes[variable.Name].Value))
                                        {
                                            activeDocument.Variables[variable.Name].Value = tdmsObj.Attributes[variable.Name].User.LastName;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    case "A_REF_PART_WORK":
                                        TDMSObject tdmsContr = tdmsObj.Attributes["A_REF_CONTRACT"].Object;
                                        if (!string.IsNullOrEmpty(tdmsContr.Attributes[variable.Name].Value))
                                        {
                                            activeDocument.Variables[variable.Name].Value = tdmsContr.Attributes[variable.Name].Value;
                                        }
                                        else
                                        {
                                            activeDocument.Variables[variable.Name].Value = " ";
                                        }
                                        continue;
                                    default:
                                        {
                                            continue;
                                        }
                                }
                            }
                        }
                        activeDocument.Fields.Update();
                        activeDocument.PrintPreview();
                        activeDocument.ClosePrintPreview();
                    }
                }
                else
                {
                    MessageBox.Show("Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    return;
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message + " " + exception.TargetSite);
            }
        }
    }
}