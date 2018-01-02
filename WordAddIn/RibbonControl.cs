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
                    Word.Document wrdDoc = wordApp.Documents.Application.ActiveDocument;
                    
                    if (Condition.CheckPathToTdms(wrdDoc.Path))
                    {
                        string str = Condition.ParseGUID(wrdDoc.Path);
                        
                        var tdmsApp = (TDMSApplication)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("430C37CB-33C4-4754-836A-A3930689D437")));

                        TDMSUser user = null;
                        TDMSObject tdmsObj = tdmsApp.GetObjectByGUID(Condition.ParseGUID(wrdDoc.Path));
                        TDMSAttributes parAttrs = tdmsObj.Attributes;
                        var wrdVars = wrdDoc.Variables;

                        // ПСД
                        if (string.Equals(tdmsObj.ObjectDefName, "O_PSD", StringComparison.InvariantCultureIgnoreCase) ||
                            string.Equals(tdmsObj.ObjectDefName, "O_DOC_SENDING", StringComparison.InvariantCultureIgnoreCase))
                        {
                            foreach (Word.Variable wrdVar in wrdVars)
                            {
                                switch (wrdVar.Name)
                                {
                                    //Получить полное наименование ООО
                                    case "F_GetOOOName":

                                        string GetOOOName = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetOOOName");

                                        if (!string.IsNullOrEmpty(GetOOOName))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = GetOOOName;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;

                                    //Получить юр.адрес ООО "Студия 44"
                                    case "F_GetOOOUrAddress":

                                        string GetOOOUrAddress = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetOOOUrAddress");

                                        if (!string.IsNullOrEmpty(GetOOOUrAddress))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = GetOOOUrAddress;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;

                                    //Получить телефоны ООО "Студия 44"
                                    case "F_GetOOOPhones":
                                        string GetOOOPhones = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetOOOPhones");

                                        if (!string.IsNullOrEmpty(GetOOOPhones))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = GetOOOPhones;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    case "F_GetDay":
                                        var GetDay = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDay");
                                        wrdDoc.Variables[wrdVar.Name].Value = Convert.ToString(GetDay);
                                        break;
                                    case "F_GetMonth":
                                        var GetMonth = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetMonth");
                                        wrdDoc.Variables[wrdVar.Name].Value = Convert.ToString(GetMonth);
                                        break;
                                    case "F_GetYear":
                                        var GetYear = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetYear");
                                        wrdDoc.Variables[wrdVar.Name].Value = Convert.ToString(GetYear);
                                        break;
                                    //Получить полное наименование объекта проектирования
                                    case "F_GetOPName":
                                        string GetOPName = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetOPName", tdmsObj);

                                        if (!string.IsNullOrEmpty(GetOPName))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = GetOPName;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;

                                    //Получить номер договора с объекта отправки
                                    case "F_GetContractNum":
                                        string ContractNum = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetContractNum", tdmsObj);

                                        if (!string.IsNullOrEmpty(ContractNum))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = ContractNum;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    //Получить дату договора с объекта отправки
                                    case "F_GetContractDate":
                                        string ContractDate = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetContractDate", tdmsObj);

                                        if (!string.IsNullOrEmpty(ContractDate))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = ContractDate;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    //Получить номер накладной
                                    case "F_GetSendingNum":
                                        string SendingNum = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetSendingNum", tdmsObj);

                                        if (!string.IsNullOrEmpty(SendingNum))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = SendingNum;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    //Получить полное наименование организации для накладной
                                    case "F_GetCompanyName":
                                        string CompanyName = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetCompanyName", tdmsObj);

                                        if (!string.IsNullOrEmpty(CompanyName))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = CompanyName;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    //Получить имя пользователя
                                    case "A_User":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            if (!string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = tdmsObj.Attributes[wrdVar.Name].Value;
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;

                                    //получить контакты
                                    case "A_CONTACT_REF":
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            if (!string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = tdmsObj.Attributes[wrdVar.Name].Value;
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }

                                        break;

                                    case "A_OBOZN_DOC":
                                    case "A_NAME":
                                    case "A_ARCH_SIGN":
                                    case "A_TOM_PAGE_NUM":
                                        // tdms3,4

                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            if (!string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = tdmsObj.Attributes[wrdVar.Name].Value;
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }

                                        break;
                                    case "A_STAGE":
                                    case "A_STAGE_CLSF":
                                        //   tdms3      tdms4

                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            if (!string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = tdmsObj.Attributes[wrdVar.Name].Classifier.Code;
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }

                                        break;
                                    case "A_OBOZN":
                                    case "A_INSTEAD_OF_NUM":
                                    case "A_LIC_NUM_P":
                                    case "A_LIC_NUM_IZ":
                                        // tdms3,4     

                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            if (!string.IsNullOrEmpty(parAttrs[wrdVar.Name].Value))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = parAttrs[wrdVar.Name].Value;
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }

                                        break;
                                    case "A_YEAR":
                                        // tdms3,4

                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            if (!string.IsNullOrEmpty(Convert.ToString(parAttrs[wrdVar.Name].Value)))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = Convert.ToString(parAttrs[wrdVar.Name].Value);
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }

                                        break;
                                    case "A_TOM_NAME_ADD":
                                        // tdms3,4

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
                                            wrdDoc.Variables[wrdVar.Name].Value = "Основной комплект рабочих чертежей" + Strings.Chr(13) + Strings.Chr(10) + name_add;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = "Основной комплект рабочих чертежей";
                                        }
                                        break;
                                    case "A_TOM_NUM":
                                        // tdms3,4

                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_TOM_NUMB"))
                                        {
                                            if (!string.IsNullOrEmpty(parAttrs["A_TOM_NUMB"].Value))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = parAttrs["A_TOM_NUMB"].Value;
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }

                                        break;
                                    case "A_TOM_NAME":
                                        // tdms3,4

                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_TOM_NUMB"))
                                        {
                                            string tom_num = parAttrs["A_TOM_NUMB"].Value;

                                            if (tom_num == "5")
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = "Сведения об инженерном оборудовании, о сетях инженерно-технического обеспечения, перечень инженерно-технических мероприятий, содержание технических решений";
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
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    case "A_SUB_TOM_NUM":
                                        // tdms3,4
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_SUBTOM_NUMB"))
                                        {
                                            if (!string.IsNullOrEmpty(parAttrs["A_SUBTOM_NUMB"].Value))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = "Подраздел " + parAttrs["A_SUBTOM_NUMB"].Value;
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }

                                        break;
                                    case "A_SUB_TOM_NAME":
                                        // tdms3,4

                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_TOM_NUMB"))
                                        {
                                            string tom_num = parAttrs["A_TOM_NUMB"].Value;

                                            if (tom_num == "5" | tom_num == "12")
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = Strings.Chr(171) + parAttrs["A_NAME"].Value + Strings.Chr(187);
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    case "A_BOOK_NUM":
                                        // tdms3,4
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_BOOK_NUMB"))
                                        {
                                            if (!string.IsNullOrEmpty(parAttrs["A_BOOK_NUMB"].Value))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = parAttrs["A_BOOK_NUMB"].Value;
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    case "A_ADDRESS":
                                        // tdms3,4

                                        //Dim tdmsContr As TDMS.TDMSObject = tdmsObj.Attributes("A_REF_CONTRACT").Object
                                        //If Not tdmsContr.Attributes("A_Comment_Init").Value = "" Then
                                        //    wrdDoc.Variables(wrdVar.Name).Value = tdmsContr.Attributes("A_Comment_Init").Value
                                        //Else
                                        //    wrdDoc.Variables(wrdVar.Name).Value = " "
                                        //End If

                                        string op_address = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetOPAddress", tdmsObj);
                                        if (!string.IsNullOrEmpty(op_address))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = op_address;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }

                                        break;
                                    case "A_REF_PART_WORK":
                                        // tdms3,4

                                        string op_name1 = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetOPName", tdmsObj);
                                        if (!string.IsNullOrEmpty(op_name1))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = op_name1;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    case "A_TEXT_FORM":
                                        // tdms3

                                        string txt_form = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetTextForm", tdmsObj);
                                        if (!string.IsNullOrEmpty(txt_form))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = txt_form;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    case "A_TITLE":
                                        // tdms3,4
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has("A_REF_CONTRACT"))
                                        {
                                            TDMSObject tdmsContr = tdmsObj.Attributes["A_REF_CONTRACT"].Object;
                                            if (!string.IsNullOrEmpty(tdmsContr.Attributes["A_COMMENT"].Value))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = Strings.UCase(tdmsContr.Attributes["A_COMMENT"].Value);
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            string title = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromStageFolder", tdmsObj, wrdVar.Name);
                                            if (!string.IsNullOrEmpty(title))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = Strings.UCase(title);
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        break;
                                    case "A_TITLE_ZAO":
                                        // tdms4
                                        string title_zao = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromContract", tdmsObj, wrdVar.Name);

                                        if (!string.IsNullOrEmpty(title_zao))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = title_zao;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }

                                        break;
                                    case "A_WRK":
                                    case "A_CHECKED":
                                    case "A_GL_SPEC":
                                    case "A_NORMK":
                                    case "A_DEPARTM_HEAD":
                                    case "A_GIP":
                                        // tdms3

                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            if (!string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value))
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = tdmsObj.Attributes[wrdVar.Name].User.LastName;
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;

                                    case "A_ADOC_DATE_ISP":
                                    case "A_DOC_DATE_CHECK":
                                    case "A_GL_SPEC_DATE":
                                    case "A_NORMK_DATE":
                                    case "A_DEPARTM_HEAD_DATE":
                                    case "A_DATE_SIGN_GIP":
                                        // tdms3

                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(wrdVar.Name))
                                        {
                                            if (!string.IsNullOrEmpty(Convert.ToString(tdmsObj.Attributes[wrdVar.Name].Value)))
                                            {
                                                string str_date = Convert.ToString(tdmsObj.Attributes[wrdVar.Name].Value);
                                                wrdDoc.Variables[wrdVar.Name].Value = Strings.Mid(str_date, 4, 3) + Strings.Right(str_date, 2);
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    case "A_GIP_FIO":
                                    case "A_DEPARTM_HEAD_FIO":
                                    case "A_NORMK_FIO":
                                    case "A_GL_SPEC_FIO":
                                    case "A_CHECKED_FIO":
                                    case "A_WRK_FIO":
                                        // tdms3

                                        this.StringVariable = Strings.Left(wrdVar.Name, Strings.Len(wrdVar.Name) - 4);
                                        if (tdmsObj.ObjectDef.AttributeDefs.Has(this.StringVariable))
                                        {
                                            if (!string.IsNullOrEmpty(tdmsObj.Attributes[this.StringVariable].Value))
                                            {
                                                user = tdmsObj.Attributes[this.StringVariable].User;
                                                wrdDoc.Variables[wrdVar.Name].Value = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                            }
                                            else
                                            {
                                                wrdDoc.Variables[wrdVar.Name].Value = " ";
                                            }
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;

                                    case "A_DEVELOP":
                                    case "A_CHECK":
                                    case "A_NORMKL":
                                    case "A_GR_HEAD":
                                        // tdms4

                                        ScriptResult = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromActiveRouteTable", tdmsObj, Strings.Mid(wrdVar.Name, 3), "A_User");

                                        if (!string.IsNullOrEmpty(ScriptResult))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = tdmsApp.Users[ScriptResult].LastName;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;

                                    case "A_GAP_":
                                    case "A_GIP_":
                                    case "A_GKP_":
                                        // tdms4
                                        ScriptResult = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromStageFolder", tdmsObj, wrdVar.Name);

                                        if (!string.IsNullOrEmpty(ScriptResult))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = tdmsApp.Users[ScriptResult].LastName;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    case "A_GKAB_":
                                        // tdms4

                                        user = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromSysProps", wrdVar.Name);

                                        if ((user != null))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = user.LastName;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    case "A_DATE_SIGN_DEVELOP":
                                    case "A_DATE_SIGN_CHECK":
                                    case "A_DATE_SIGN_NORMKL":
                                    case "A_DATE_SIGN_GR_HEAD":
                                    case "A_DATE_SIGN_GIP_":
                                    case "A_DATE_SIGN_GAP_":
                                    case "A_DATE_SIGN_GKP_":
                                    case "A_DATE_SIGN_GKAB_":
                                        // tdms4

                                        string date_sign = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromActiveRouteTable", tdmsObj, Strings.Mid(wrdVar.Name, 13), "A_DATE");
                                        if (!string.IsNullOrEmpty(date_sign))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = date_sign;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }

                                        break;
                                    case "A_GAP_FIO":
                                    case "A_GKP_FIO":
                                        // tdms4

                                        this.StringVariable = Strings.Left(wrdVar.Name, Strings.Len(wrdVar.Name) - 3);
                                        ScriptResult = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromStageFolder", tdmsObj, this.StringVariable);

                                        if (!string.IsNullOrEmpty(ScriptResult))
                                        {
                                            user = tdmsApp.Users[ScriptResult];
                                            wrdDoc.Variables[wrdVar.Name].Value = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;

                                    case "A_GIP__FIO":
                                        // tdms4

                                        this.StringVariable = Strings.Left(wrdVar.Name, Strings.Len(wrdVar.Name) - 4);

                                        ScriptResult = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromStageFolder", tdmsObj, this.StringVariable);

                                        if (!string.IsNullOrEmpty(ScriptResult))
                                        {
                                            user = tdmsApp.Users[ScriptResult];
                                            wrdDoc.Variables[wrdVar.Name].Value = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    case "A_GKAB_FIO":
                                        // tdms4

                                        this.StringVariable = Strings.Left(wrdVar.Name, Strings.Len(wrdVar.Name) - 3);
                                        user = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromSysProps", this.StringVariable);
                                        if ((user != null))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                    case "A_GR_HEAD_FIO":
                                        // tdms4

                                        string sys_name = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetDataFromActiveRouteTable", tdmsObj, Strings.Mid(wrdVar.Name, 3, 7), "A_User");

                                        if (!string.IsNullOrEmpty(sys_name))
                                        {
                                            user = tdmsApp.Users[sys_name];
                                            wrdDoc.Variables[wrdVar.Name].Value = tdmsApp.ExecuteScript("CMD_SYSLIB", "GetUserFIO", user);
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;
                                }
                            }

                            tdmsApp.ExecuteScript("CMD_SYSLIB", "FillInSndngDocList", tdmsObj, wrdDoc);
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
                                        if (!string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = tdmsObj.Attributes[wrdVar.Name].Value;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;

                                    case "A_DEPART_TO":
                                    case "A_STAGE":
                                        if (!string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = tdmsObj.Attributes[wrdVar.Name].Classifier.Code;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;

                                    case "A_DEPART_FROM":

                                        if (!string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = Strings.Left(tdmsObj.Attributes["A_DEPART_FROM"].Classifier.Code, 3);
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;

                                    case "A_GIP":
                                    case "A_DEPARTM_HEAD":
                                    case "A_GL_SPEC":
                                    case "A_GROUP_HEAD":
                                    case "A_WRK":
                                        if (!string.IsNullOrEmpty(tdmsObj.Attributes[wrdVar.Name].Value))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = tdmsObj.Attributes[wrdVar.Name].User.LastName;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
                                        break;

                                    case "A_REF_PART_WORK":
                                        TDMSObject tdmsContr = tdmsObj.Attributes["A_REF_CONTRACT"].Object;
                                        if (!string.IsNullOrEmpty(tdmsContr.Attributes[wrdVar.Name].Value))
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = tdmsContr.Attributes[wrdVar.Name].Value;
                                        }
                                        else
                                        {
                                            wrdDoc.Variables[wrdVar.Name].Value = " ";
                                        }
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
                        MessageBox.Show("Документ не принадлежит TDMS.");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Невозможно выполнить команду, т.к. TDMS не запущен или количество запущенных приложений TDMS более одного.");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}