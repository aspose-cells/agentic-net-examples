using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DecryptOdsWorkbook
    {
        public static void Run(string inputPath, string password, string outputPath)
        {
            var loadOptions = new LoadOptions();
            loadOptions.Password = password;

            var workbook = new Workbook(inputPath, loadOptions);

            try
            {
                workbook.Unprotect(password);
            }
            catch { }

            workbook.Settings.Password = null;

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.IsProtected)
                {
                    sheet.Unprotect();
                }
            }

            workbook.Save(outputPath);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: <inputPath> <password> <outputPath>");
                return;
            }

            string inputPath = args[0];
            string password = args[1];
            string outputPath = args[2];

            DecryptOdsWorkbook.Run(inputPath, password, outputPath);
        }
    }
}