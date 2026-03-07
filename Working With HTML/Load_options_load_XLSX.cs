using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLoadXlsxExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string inputPath = "sample.xlsx";

            // Create LoadOptions with default settings
            LoadOptions loadOptions = new LoadOptions();

            // Example of setting some load options
            loadOptions.Password = "myPassword";               // If the file is password protected
            loadOptions.CultureInfo = new CultureInfo("en-US"); // Regional settings
            loadOptions.ParsingFormulaOnOpen = true;           // Parse formulas while loading
            loadOptions.LoadFilter = new LoadFilter
            {
                // Load only cell values and formulas, ignore formatting and drawings
                LoadDataFilterOptions = LoadDataFilterOptions.CellValue | LoadDataFilterOptions.Formula
            };

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example operation: write a message into cell A1
            sheet.Cells["A1"].PutValue("Loaded with custom LoadOptions");

            // Save the workbook to a new file (XLSX format)
            string outputPath = "sample_loaded.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook loaded from '{inputPath}' and saved to '{outputPath}'.");
        }
    }
}