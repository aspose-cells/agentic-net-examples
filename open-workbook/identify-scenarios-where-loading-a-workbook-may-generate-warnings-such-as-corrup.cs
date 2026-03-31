using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWarningDemo
{
    // Custom warning handler that prints warning details to the console
    public class CustomWarningCallback : IWarningCallback
    {
        public void Warning(WarningInfo warningInfo)
        {
            Console.WriteLine($"[Warning] Type: {warningInfo.Type}, Description: {warningInfo.Description}");
        }
    }

    public class Program
    {
        private static string EnsureWorkbook(string fileName)
        {
            if (!File.Exists(fileName))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Cells["A1"].PutValue($"Sample data for {Path.GetFileNameWithoutExtension(fileName)}");
                wb.Save(fileName);
            }
            return Path.GetFullPath(fileName);
        }

        public static void Main()
        {
            // -----------------------------------------------------------------
            // Scenario 1: Loading a potentially corrupted workbook.
            // -----------------------------------------------------------------
            string corruptedPath = EnsureWorkbook("CorruptedFile.xlsx");
            var corruptedOptions = new LoadOptions { WarningCallback = new CustomWarningCallback() };
            var corruptedWorkbook = new Workbook(corruptedPath, corruptedOptions);
            Console.WriteLine($"Corrupted workbook first cell: {corruptedWorkbook.Worksheets[0].Cells["A1"].StringValue}");

            // -----------------------------------------------------------------
            // Scenario 2: Loading a workbook that uses unsupported features.
            // -----------------------------------------------------------------
            string unsupportedPath = EnsureWorkbook("UnsupportedFeature.xlsx");
            var unsupportedOptions = new LoadOptions { WarningCallback = new CustomWarningCallback() };
            var unsupportedWorkbook = new Workbook(unsupportedPath, unsupportedOptions);
            Console.WriteLine($"Unsupported feature workbook sheet count: {unsupportedWorkbook.Worksheets.Count}");

            // -----------------------------------------------------------------
            // Scenario 3: Loading a workbook with data inconsistencies.
            // -----------------------------------------------------------------
            string inconsistentPath = EnsureWorkbook("DataInconsistency.xlsx");
            var inconsistentOptions = new LoadOptions { WarningCallback = new CustomWarningCallback() };
            var inconsistentWorkbook = new Workbook(inconsistentPath, inconsistentOptions);
            Console.WriteLine($"Inconsistent workbook first defined name count: {inconsistentWorkbook.Worksheets.Names.Count}");

            // -----------------------------------------------------------------
            // Save the loaded workbooks (optional). Saving may also trigger warnings.
            // -----------------------------------------------------------------
            inconsistentWorkbook.Save("DataInconsistency_Output.xlsx");
            unsupportedWorkbook.Save("UnsupportedFeature_Output.xlsx");
            corruptedWorkbook.Save("CorruptedFile_Output.xlsx");
        }
    }
}