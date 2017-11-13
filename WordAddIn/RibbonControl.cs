using System;
using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Word = Microsoft.Office.Interop.Word;

using TDMS.Interop;

namespace WordAddIn
{
    public partial class TDMS
    {
        private readonly string activeApplication = "Word.Application";
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
            try
            {
                UpdateTDMSVariables();
            }
            catch
            {
                MessageBox.Show(@"Ошибка обновления.");
            }
        }

        public void UpdateTDMSVariables()
        {
            try
            {
                if (Condition.CheckTDMSProcess())
                {
                    Word.Application WordApp = (Microsoft.Office.Interop.Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                    Word.Document wrdDoc = WordApp.Documents.Application.ActiveDocument;

                    if (Condition.CheckPathToTdms(wrdDoc.Path))
                    {
                        TDMSAttributes parAttrs = null;
                        var tdmsObj = new TDMSApplication().GetObjectByGUID(Condition.ParseGUID(wrdDoc.Path));
                        var wrdVars = wrdDoc.Variables;
                        
                        if (tdmsObj.Uplinks.Count == 0)
                        {
                            if (tdmsObj.ObjectDefName != "O_TASK")
                            {
                                return;
                            }
                        }
                        else
                        {
                            tdmsObj = tdmsObj.Uplinks[0];
                            if (!((tdmsObj.ObjectDefName == "O_PSD" & tdmsObj.ObjectDefName == "O_TOM") |
                                  (tdmsObj.ObjectDefName == "O_PSD" & tdmsObj.ObjectDefName == "O_PSD_FOLDER")))
                            {
                                return;
                            }

                            parAttrs = tdmsObj.Attributes;
                        }

                        // ПСД
                        if (tdmsObj.ObjectDefName == "O_PSD")
                        {
                            foreach (Word.Variable wrdVar in wrdVars)
                            {
                                MessageBox.Show(wrdVar.Name);
                                TDMSUser user;
                                switch (wrdVar.Name)
                                {
                                    //Получить полное наименование ООО
                                    case "F_GetOOOName":
                                        string GetOOOName = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetOOOName");
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(GetOOOName) ? GetOOOName : string.Empty;
                                        break;
                                    //Получить юр.адрес ООО "Студия 44"
                                    case "F_GetOOOUrAddress":
                                        string GetOOOUrAddress = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetOOOUrAddress");
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(GetOOOUrAddress) ? GetOOOUrAddress : string.Empty;
                                        break;
                                    //Получить телефоны ООО "Студия 44"
                                    case "F_GetOOOPhones":
                                        string GetOOOPhones = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetOOOPhones");
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(GetOOOPhones) ? GetOOOPhones : string.Empty;
                                        break;
                                    //Получить полное наименование объекта проектирования
                                    case "F_GetOPName":
                                        string op_name2 = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetOPName", tdmsObj);
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(op_name2) ? op_name2 : string.Empty;
                                        break;
                                    //Получить номер договора с объекта отправки
                                    case "F_GetContractNum":
                                        string ContractNum = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetContractNum", tdmsObj);
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(ContractNum) ? ContractNum : string.Empty;
                                        break;
                                    //Получить дату договора с объекта отправки
                                    case "F_GetContractDate":
                                        string ContractDate = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetContractDate", tdmsObj);
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(ContractDate) ? ContractDate : string.Empty;
                                        break;
                                    //Получить номер накладной
                                    case "F_GetSendingNum":
                                        string SendingNum = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetSendingNum", tdmsObj);
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(SendingNum) ? SendingNum : string.Empty;
                                        break;
                                    //Получить полное наименование организации для накладной
                                    case "F_GetCompanyName": string CompanyName = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetCompanyName", tdmsObj);
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(CompanyName) ? CompanyName : string.Empty;
                                        break;
                                    //Получить имя пользователя
                                    case "A_User":
                                        ScriptResult = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetDataFromActiveRouteTable", tdmsObj, Strings.Mid(wrdVar.Name, 3), "A_User");
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(ScriptResult) ? new TDMSApplication().Users[ScriptResult].LastName : string.Empty;
                                        break;
                                    //получить контакты
                                    case "A_CONTACT_REF":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value) ? tdmsObj.Attributes[wrdVar.Name].Value : string.Empty;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;
                                    case "A_OBOZN_DOC":
                                    case "A_NAME":
                                    case "A_ARCH_SIGN":
                                    case "A_TOM_PAGE_NUM":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value) ? tdmsObj.Attributes[wrdVar.Name].Value : string.Empty;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;
                                    case "A_STAGE":
                                    case "A_STAGE_CLSF":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value) ? tdmsObj.Attributes[wrdVar.Name].Classifier.Code : string.Empty;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;
                                    case "A_OBOZN":
                                    case "A_INSTEAD_OF_NUM":
                                    case "A_LIC_NUM_P":
                                    case "A_LIC_NUM_IZ":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(parAttrs[wrdVar.Name].Value) ? parAttrs[wrdVar.Name].Value : string.Empty;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }

                                        break;
                                    case "A_YEAR":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(Convert.ToString(parAttrs[wrdVar.Name].Value)) ? Convert.ToString(parAttrs[wrdVar.Name].Value) : string.Empty;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }

                                        break;
                                    case "A_TOM_NAME_ADD":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_STAGE"))
                                        {
                                            if (parAttrs["A_STAGE"].Classifier.Code != "Р")
                                            {
                                                return;
                                            }
                                        }
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_STAGE_CLSF"))
                                        {
                                            if (parAttrs["A_STAGE_CLSF"].Classifier.Code != "Р")
                                            {
                                                return;
                                            }
                                        }

                                        string mark_code = parAttrs["A_MARK"].Classifier.Code;
                                        string name_add = new TDMSApplication().ExecuteScript("O_PSD", "GetTomNameAddForTitul", mark_code); 

                                        if (!string.IsNullOrEmpty(name_add))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = "Основной комплект рабочих чертежей" + Strings.Chr(13) + Strings.Chr(10) + name_add;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = "Основной комплект рабочих чертежей";
                                        }
                                        break;
                                    case "A_TOM_NUM":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_TOM_NUMB"))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(parAttrs["A_TOM_NUMB"].Value) ? parAttrs["A_TOM_NUMB"].Value : string.Empty;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;
                                    case "A_TOM_NAME":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_TOM_NUMB"))
                                        {
                                            string tom_num = parAttrs["A_TOM_NUMB"].Value;

                                            if (tom_num == "5")
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value =  "Сведения об инженерном оборудовании, о сетях инженерно-технического обеспечения, перечень инженерно-технических мероприятий, содержание технических решений";
                                            }
                                            else if (tom_num == "12")
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = "Иная документация в случаях, предусмотренных федеральными законами";
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = parAttrs["A_NAME"].Value;
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;
                                    case "A_SUB_TOM_NUM":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_SUBTOM_NUMB"))
                                        {
                                            if (!string.IsNullOrEmpty(parAttrs["A_SUBTOM_NUMB"].Value))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = "Подраздел " + parAttrs["A_SUBTOM_NUMB"].Value;
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;
                                    case "A_SUB_TOM_NAME":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_TOM_NUMB"))
                                        {
                                            string tom_num = parAttrs["A_TOM_NUMB"].Value;

                                            if (tom_num == "5" | tom_num == "12")
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = Strings.Chr(171) + parAttrs["A_NAME"].Value + Strings.Chr(187);
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;
                                    case "A_BOOK_NUM":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_BOOK_NUMB"))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(parAttrs["A_BOOK_NUMB"].Value) ? parAttrs["A_BOOK_NUMB"].Value : string.Empty;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;
                                    case "A_ADDRESS":
                                        string op_address = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetOPAddress", tdmsObj);
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(op_address) ? op_address : string.Empty;
                                        break;
                                    case "A_REF_PART_WORK":
                                        string op_name1 = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetOPName", tdmsObj);
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(op_name1) ? op_name1 : string.Empty;
                                        break;
                                    case "A_TEXT_FORM":
                                        string txt_form = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetTextForm", tdmsObj);
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(txt_form) ? txt_form : string.Empty;
                                        break;
                                    case "A_TITLE":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_REF_CONTRACT"))
                                        {
                                            TDMSObject tdmsContr = tdmsObj.Attributes["A_REF_CONTRACT"].Object;
                                            wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(tdmsContr.Attributes["A_COMMENT"].Value) ? Strings.UCase(tdmsContr.Attributes["A_COMMENT"].Value) : string.Empty;
                                        }
                                        else
                                        {
                                            string title = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetDataFromStageFolder", tdmsObj, wrdVar.Name);
                                            wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(title) ? Strings.UCase(title) : string.Empty;
                                        }
                                        break;
                                    case "A_TITLE_ZAO":
                                        string title_zao = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetDataFromContract", tdmsObj, wrdVar.Name);
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(title_zao) ? title_zao : string.Empty;
                                        break;
                                    case "A_WRK":
                                    case "A_CHECKED":
                                    case "A_GL_SPEC":
                                    case "A_NORMK":
                                    case "A_DEPARTM_HEAD":
                                    case "A_GIP":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value) ? tdmsObj.Attributes[wrdVar.Name].User.LastName : string.Empty;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;

                                    case "A_ADOC_DATE_ISP":
                                    case "A_DOC_DATE_CHECK":
                                    case "A_GL_SPEC_DATE":
                                    case "A_NORMK_DATE":
                                    case "A_DEPARTM_HEAD_DATE":
                                    case "A_DATE_SIGN_GIP":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            if (!string.IsNullOrEmpty(Convert.ToString(tdmsObj.Attributes[wrdVar.Name].Value)))
                                            {
                                                string str_date = Convert.ToString(tdmsObj.Attributes[wrdVar.Name].Value);
                                                wrdDoc.Variables[wrdVar.Name].Value = Strings.Mid(str_date, 4, 3) + Strings.Right(str_date, 2);
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;
                                    case "A_GIP_FIO":
                                    case "A_DEPARTM_HEAD_FIO":
                                    case "A_NORMK_FIO":
                                    case "A_GL_SPEC_FIO":
                                    case "A_CHECKED_FIO":
                                    case "A_WRK_FIO":
                                        this.StringVariable = Strings.Left(wrdVar.Name, Strings.Len(wrdVar.Name) - 4);
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(this.StringVariable))
                                        {
                                            if (!string.IsNullOrEmpty(tdmsObj.Attributes[this.StringVariable].Value))
                                            {
                                                user = tdmsObj.Attributes[this.StringVariable].User;
                                                wrdDoc.Variables[wrdVar.Name].Value = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;

                                    case "A_DEVELOP":
                                    case "A_CHECK":
                                    case "A_NORMKL":
                                    case "A_GR_HEAD":
                                        ScriptResult = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetDataFromActiveRouteTable", tdmsObj, Strings.Mid(wrdVar.Name, 3), "A_User");
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(ScriptResult) ? new TDMSApplication().Users[ScriptResult].LastName : string.Empty;
                                        break;

                                    case "A_GAP_":
                                    case "A_GIP_":
                                    case "A_GKP_":
                                        ScriptResult = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetDataFromStageFolder", tdmsObj, wrdVar.Name);
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(ScriptResult) ? new TDMSApplication().Users[ScriptResult].LastName : string.Empty;
                                        break;
                                    case "A_GKAB_":
                                        user = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetDataFromSysProps", wrdVar.Name);
                                        wrdDoc.Variables[wrdVar.Name].Value = user != null ? user.LastName : string.Empty;
                                        break;
                                    case "A_DATE_SIGN_DEVELOP":
                                    case "A_DATE_SIGN_CHECK":
                                    case "A_DATE_SIGN_NORMKL":
                                    case "A_DATE_SIGN_GR_HEAD":
                                    case "A_DATE_SIGN_GIP_":
                                    case "A_DATE_SIGN_GAP_":
                                    case "A_DATE_SIGN_GKP_":
                                    case "A_DATE_SIGN_GKAB_":
                                        string date_sign = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetDataFromActiveRouteTable", tdmsObj, Strings.Mid(wrdVar.Name, 13), "A_DATE");
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(date_sign) ? date_sign : string.Empty;
                                        break;
                                    case "A_GAP_FIO":
                                    case "A_GKP_FIO":
                                        this.StringVariable = Strings.Left(wrdVar.Name, Strings.Len(wrdVar.Name) - 3);
                                        ScriptResult = new TDMSApplication().ExecuteScript( "CMD_SYSLIB", "GetDataFromStageFolder", tdmsObj, this.StringVariable);

                                        if (!string.IsNullOrEmpty(ScriptResult))
                                        {
                                            user = new TDMSApplication().Users[ScriptResult];
                                            wrdDoc.Variables[wrdVar.Name].Value = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;
                                    case "A_GIP__FIO":
                                        this.StringVariable = Strings.Left(wrdVar.Name, Strings.Len(wrdVar.Name) - 4);
                                        ScriptResult = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetDataFromStageFolder", tdmsObj, this.StringVariable);

                                        if (!string.IsNullOrEmpty(ScriptResult))
                                        {
                                            user = new TDMSApplication().Users[ScriptResult];
                                            wrdDoc.Variables[wrdVar.Name].Value = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;
                                    case "A_GKAB_FIO":
                                        this.StringVariable = Strings.Left(wrdVar.Name, Strings.Len(wrdVar.Name) - 3);
                                        user = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetDataFromSysProps", this.StringVariable);
                                        wrdDoc.Variables[wrdVar.Name].Value = user != null ? new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetUserFIO", user) : string.Empty;
                                        break;
                                    case "A_GR_HEAD_FIO":
                                        string sys_name = new TDMSApplication().ExecuteScript( "CMD_SYSLIB", "GetDataFromActiveRouteTable", tdmsObj, Strings.Mid(wrdVar.Name, 3, 7), "A_User");

                                        if (!string.IsNullOrEmpty(sys_name))
                                        {
                                            user = new TDMSApplication().Users[sys_name];
                                            wrdDoc.Variables[wrdVar.Name].Value = new TDMSApplication().ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = string.Empty;
                                        }
                                        break;

                                }
                            }
                        }
                        else
                        {
                            foreach (Word.Variable wrdVar in wrdVars)
                            {
                                switch (wrdVar.Name)
                                {
                                    case "A_REG_NUM":
                                    case "A_CONTRACT_SHIFR":
                                    case "A_NAME_WORK":
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value) ? tdmsObj.Attributes[wrdVar.Name].Value : string.Empty;
                                        break;

                                    case "A_DEPART_TO":
                                    case "A_STAGE":
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value) ? tdmsObj.Attributes[wrdVar.Name].Classifier.Code : string.Empty;
                                        break;

                                    case "A_DEPART_FROM":
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value) ? Strings.Left(tdmsObj.Attributes["A_DEPART_FROM"].Classifier.Code, 3) : string.Empty;
                                        break;

                                    case "A_GIP":
                                    case "A_DEPARTM_HEAD":
                                    case "A_GL_SPEC":
                                    case "A_GROUP_HEAD":
                                    case "A_WRK":
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value) ? tdmsObj.Attributes[wrdVar.Name].User.LastName : string.Empty;
                                        break;

                                    case "A_REF_PART_WORK":
                                        TDMSObject tdmsContr = tdmsObj.Attributes["A_REF_CONTRACT"].Object;
                                        wrdDoc.Variables[wrdVar.Name].Value = !string.IsNullOrEmpty(tdmsContr.Attributes[wrdVar.Name].Value) ? tdmsContr.Attributes[wrdVar.Name].Value : string.Empty;
                                        break;
                                }
                            }
                        }
                        wrdDoc.Fields.Update();
                        wrdDoc.PrintPreview();
                        wrdDoc.ClosePrintPreview();
                    }
                    else
                    {
                        MessageBox.Show(@"Документ не принадлежит TDMS.");
                    }
                }
                else
                {
                    MessageBox.Show(@"Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}