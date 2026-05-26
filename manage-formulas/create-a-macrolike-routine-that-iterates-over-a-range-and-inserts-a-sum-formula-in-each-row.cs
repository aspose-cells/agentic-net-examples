using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SumFormulaMacro
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample data: fill columns A‑D (0‑3) with some numbers
                int totalRows = 5;
                int dataColumns = 4;
                for (int row = 0; row < totalRows; row++)
                {
                    for (int col = 0; col < dataColumns; col++)
                    {
                        cells[row, col].PutValue(row * dataColumns + col + 1);
                    }
                }

                // Insert a SUM formula in each row that adds the values of A‑D, result in column E
                int targetColumnIndex = 4; // column E
                Cell firstTargetCell = cells[0, targetColumnIndex];
                string sharedFormula = "=SUM(A1:D1)";

                // Apply shared formula using FormulaParseOptions (newer API)
                FormulaParseOptions options = new FormulaParseOptions();
                firstTargetCell.SetSharedFormula(sharedFormula, totalRows, 1, options);

                // Recalculate formulas
                workbook.CalculateFormula();

                // Save the workbook
                string outputPath = "SumRows.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            SumFormulaMacro.Run();
        }
    }
}