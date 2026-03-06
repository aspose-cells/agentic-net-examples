using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

            ProtectWorksheetDemo.Run(inputPath, outputPath);
        }
    }

    public class ProtectWorksheetDemo
    {
        public static void Run(string inputPath, string outputPath)
        {
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Protect(ProtectionType.All, "myPassword", string.Empty);
            workbook.Save(outputPath);
        }
    }
}