using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class TreatTextAsZeroDemo
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate cells with numeric values and textual values
                cells["A1"].PutValue(10);          // numeric
                cells["A2"].PutValue("20");        // text that looks like a number
                cells["A3"].PutValue("Thirty");    // non‑numeric text
                cells["A4"].PutValue(40);          // numeric

                // Set a formula that sums the range A1:A4
                cells["B1"].Formula = "=SUM(A1:A4)";

                // Calculate without special options (text is ignored, non‑numeric text treated as 0)
                workbook.CalculateFormula();
                Console.WriteLine("Result without TreatTextAsZero: " + cells["B1"].DoubleValue);
                // Expected output: 50 (10 + 20 + 0 + 40)

                // Create calculation options (TreatTextAsZero is not available in this version,
                // but the options object can still be used for other settings)
                CalculationOptions calcOptions = new CalculationOptions();

                // Re‑calculate the formula using the options
                sheet.CalculateFormula(calcOptions, false);
                Console.WriteLine("Result with default CalculationOptions: " + cells["B1"].DoubleValue);
                // Expected output: 50

                // Save the workbook (optional)
                string outputPath = "TreatTextAsZeroDemo.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}