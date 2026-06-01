using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExtractArrayFormulaTextDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                for (int i = 0; i < 10; i++)
                {
                    worksheet.Cells[i, 0].PutValue(i + 1);          // Column A (A1:A10)
                    worksheet.Cells[i, 1].PutValue((i + 1) * 2);   // Column B (B1:B10)
                }

                // Define an array formula
                string arrayFormula = "SUM(A1:A10*B1:B10)";

                // Apply the array formula to cell B2
                Cell targetCell = worksheet.Cells["B2"];
                targetCell.SetArrayFormula(arrayFormula, 1, 1); // result placed in B2

                // Retrieve the formula text
                string formulaText = targetCell.GetFormula(false, false);
                Console.WriteLine("Extracted array formula text: " + formulaText);

                // Save the workbook
                string outputPath = "ExtractArrayFormulaTextDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ExtractArrayFormulaTextDemo.Run();
        }
    }
}