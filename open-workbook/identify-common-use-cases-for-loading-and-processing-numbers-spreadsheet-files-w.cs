using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace NumbersProcessingDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Numbers file
            string numbersFilePath = "Sample.numbers";

            // Ensure the sample file exists; if not, create a simple workbook and save as Numbers
            if (!File.Exists(numbersFilePath))
            {
                var wb = new Workbook();
                var sheet = wb.Worksheets[0];
                sheet.Name = "DemoSheet";
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue(123);
                sheet.Cells["C1"].Formula = "=B1*2";
                wb.Save(numbersFilePath, SaveFormat.Numbers);
                wb.Dispose();
            }

            // -----------------------------------------------------------------
            // 1. Load a Numbers file with default options and read basic info
            // -----------------------------------------------------------------
            using (Workbook defaultWorkbook = new Workbook(numbersFilePath))
            {
                Console.WriteLine("Default load:");
                Console.WriteLine($"  Worksheets count: {defaultWorkbook.Worksheets.Count}");
                Console.WriteLine($"  First sheet name: {defaultWorkbook.Worksheets[0].Name}");

                // Read a sample cell value (A1) from the first worksheet
                var firstCellValue = defaultWorkbook.Worksheets[0].Cells["A1"].Value;
                Console.WriteLine($"  Cell A1 value: {firstCellValue}");
            }

            // ---------------------------------------------------------------
            // 2. Load a Numbers file with custom LoadOptions
            //    - Provide password if the file is protected
            //    - Set regional settings
            // ---------------------------------------------------------------
            var loadOptions = new LoadOptions(LoadFormat.Numbers)
            {
                Password = "mySecretPassword",          // set if the file is password‑protected
                CultureInfo = new CultureInfo("en-US")  // set regional settings
            };

            using (Workbook customWorkbook = new Workbook(numbersFilePath, loadOptions))
            {
                Console.WriteLine("\nCustom load with LoadOptions:");
                Console.WriteLine($"  Worksheets count: {customWorkbook.Worksheets.Count}");

                // Example: iterate all worksheets and print the used range address
                foreach (Worksheet sheet in customWorkbook.Worksheets)
                {
                    var usedRange = sheet.Cells.MaxDisplayRange;
                    Console.WriteLine($"  Sheet \"{sheet.Name}\" used range: {usedRange.RefersTo}");
                }

                // ---------------------------------------------------------------
                // 3. Convert the loaded Numbers workbook to another format (e.g., XLSX)
                // ---------------------------------------------------------------
                string xlsxOutputPath = "ConvertedFromNumbers.xlsx";
                customWorkbook.Save(xlsxOutputPath, SaveFormat.Xlsx);
                Console.WriteLine($"\nConverted Numbers file saved as: {xlsxOutputPath}");

                // ---------------------------------------------------------------
                // 4. Perform a simple data manipulation: calculate formulas
                // ---------------------------------------------------------------
                customWorkbook.CalculateFormula();
                Console.WriteLine("Formulas recalculated.");

                // Save the recalculated workbook
                string recalculatedPath = "Recalculated.xlsx";
                customWorkbook.Save(recalculatedPath, SaveFormat.Xlsx);
                Console.WriteLine($"Recalculated workbook saved as: {recalculatedPath}");
            }

            // ---------------------------------------------------------------
            // 5. (Optional) AI summary placeholder - Aspose.Cells for .NET does not include CellsAI.
            // ---------------------------------------------------------------
            Console.WriteLine("\nAI summary feature is not available in Aspose.Cells for .NET.");
        }
    }
}