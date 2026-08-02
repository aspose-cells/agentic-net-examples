using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCalculationModeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Set calculation mode to AutomaticExceptTable
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data (including a header row)
                sheet.Cells["A1"].PutValue("Item");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("D");
                sheet.Cells["B5"].PutValue(40);

                // Add a table that covers the data range (including header)
                // Correct overload: Add(string name, string source, bool hasHeaders)
                sheet.ListObjects.Add("Table1", "A1:B5", true);

                // Place a formula outside the table that sums the values column
                sheet.Cells["D1"].Formula = "=SUM(B2:B5)";

                // Initial calculation so the formula has a value
                workbook.CalculateFormula();

                Console.WriteLine("Initial sum (D1): " + sheet.Cells["D1"].IntValue);

                // Change a value inside the table
                sheet.Cells["B3"].PutValue(100); // B3 is inside the table

                // Without recalculation, the formula result stays the same
                Console.WriteLine("After changing B3 (no recalculation) sum (D1): " + sheet.Cells["D1"].IntValue);

                // Manually trigger calculation
                workbook.CalculateFormula();

                // Now the formula reflects the updated table data
                Console.WriteLine("After manual CalculateFormula() sum (D1): " + sheet.Cells["D1"].IntValue);

                // Save the workbook
                string outputPath = "CalculationModeDemo.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}