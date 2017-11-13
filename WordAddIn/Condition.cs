using System;

namespace WordAddIn
{
    using System.Diagnostics;
    using System.Text.RegularExpressions;
    public static class Condition
    {
        /// <summary>
        /// Реаализует проверку принадлежности чертежа базе TDMS. Если путь сохранения файла не совпадает с путём C:\TEMP\ то метод возвращает false, в противном случае true.
        /// Метод в качестве параметра принимает путь к файлу
        /// </summary>
        private static bool CheckPath(string path)
        {
            return path.Contains(@"C:\Temp");
        }

        /// <summary>
        /// Обёртка для метода CheckPath(path) с проверкой на пустой путь, если переданный путь пуст, то false, в противном случае true
        /// </summary>
        public static bool StartCheckPath(string path)
        {
            return path != string.Empty && CheckPath(path);
        }

        /// <summary>
        /// Метод проверяет существование процесса с именем TDMS, если такой процесс существует в единственном экземпляре, то возвращает true, в противном случае false
        /// </summary>
        public static bool CheckTDMSProcess()
        {
            var process = Process.GetProcessesByName("TDMS");
            return process.Length == 1;
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