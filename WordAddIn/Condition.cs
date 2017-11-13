using System;

namespace WordAddIn
{
    using System.Diagnostics;
    using System.Text.RegularExpressions;
    public static class Condition
    {
        /// <summary>
        /// Реализует проверку принадлежности чертежа базе TDMS. Если путь сохранения файла не совпадает с путём C:\TEMP\ то метод возвращает false, в противном случае true.
        /// Метод в качестве параметра принимает путь к файлу
        /// </summary>
        /// <summary>
        /// Обёртка для метода CheckPath(path) с проверкой на пустой путь, если переданный путь пуст, то false, в противном случае true
        /// </summary>
        public static bool CheckPathToTdms(string path)
        {
            try
            {
                if (path == " " || path == string.Empty)
                {
                    return false;
                }

                path = path.Remove(7);
               
                return string.Equals(path, @"C:\Temp", StringComparison.InvariantCultureIgnoreCase);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Метод проверяет существование процесса с именем TDMS, если такой процесс существует в единственном экземпляре, то возвращает true, в противном случае false
        /// </summary>
        public static bool CheckTDMSProcess()
        {
            return Process.GetProcessesByName("TDMS").Length == 1;
        }

        public static string ParseGUID(string pathName)
        {
            string parseGuidFromFile = null;
            Regex regFF = new Regex("[{](.....................................)", RegexOptions.IgnoreCase);
            MatchCollection mcFF = regFF.Matches(pathName);

            foreach (Match mat in mcFF)
            {
                parseGuidFromFile += mat.Value;
            }
            parseGuidFromFile = parseGuidFromFile?.Remove(0, 38);
            return parseGuidFromFile;
        }
    }
}