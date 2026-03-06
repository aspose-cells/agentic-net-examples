using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsFormulaEnglishStyle
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file that contains formulas in any locale
            string inputPath = "input.xlsx";

            // Create LoadOptions for XLSX format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Specify English (US) culture to ensure formulas are interpreted in English style
            loadOptions.CultureInfo = new CultureInfo("en-US");

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Example: read a formula from cell A1 to verify it is in English format
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            Console.WriteLine("Formula in A1 (English style): " + cell.Formula);

            // Optionally, calculate all formulas to ensure they are parsed correctly
            workbook.CalculateFormula();

            // Save the workbook (optional, can be saved to a new file)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook loaded and saved with English (US) formula style.");
        }
    }
}