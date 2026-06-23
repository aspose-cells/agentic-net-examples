using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class TreatEmptyCellsAsZeroDemo
    {
        // Entry point required by the runtime
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate data with some blank cells
                cells["A1"].PutValue(10);   // numeric value
                cells["A2"].PutValue(null); // blank cell
                cells["A3"].PutValue(30);   // numeric value

                // Set a formula that sums the range A1:A3
                cells["B1"].Formula = "=SUM(A1:A3)";

                // Calculate all formulas (empty cells are treated as zero by default)
                workbook.CalculateFormula();

                // Output the result of the formula
                Console.WriteLine("Result of SUM(A1:A3) with empty cells treated as zero: " + cells["B1"].DoubleValue);
                // Expected output: 40 (10 + 0 + 30)

                // Save the workbook
                string outputPath = "TreatEmptyCellsAsZeroDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Runtime safety: report any errors
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}