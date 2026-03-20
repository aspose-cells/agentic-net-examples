using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLegacyLoadDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the legacy Excel 95/5.0 workbook (XLS format)
            string legacyFilePath = "LegacyWorkbook.xls";

            // Ensure the file exists before attempting to load
            if (!File.Exists(legacyFilePath))
            {
                Console.WriteLine($"File not found: {legacyFilePath}");
                return;
            }

            try
            {
                // Create LoadOptions specifying the Excel 97-2003 format.
                // Legacy Excel 95 files are also recognized as Excel97To2003 (XLS).
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);

                // Load the workbook using the file path and the configured LoadOptions.
                Workbook workbook = new Workbook(legacyFilePath, loadOptions);

                // Demonstrate that the workbook was loaded by accessing the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine($"Loaded worksheet name: {sheet.Name}");
                Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");

                // Optionally, save the workbook to a modern format (XLSX) to verify successful conversion.
                string outputPath = "ConvertedWorkbook.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved as: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }
}