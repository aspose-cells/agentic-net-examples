using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMacroStyle
{
    public class ConditionalSumIfRoutine
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample data: Column A = Item, Column B = Amount, Column D = Flag (TRUE/FALSE)
                int totalRows = 10;
                for (int row = 0; row < totalRows; row++)
                {
                    cells[row, 0].PutValue($"Item{row + 1}");
                    cells[row, 1].PutValue((row + 1) * 10);
                    cells[row, 3].PutValue(row % 2 == 0);
                }

                // Insert conditional SUMIF formula into Column C for each row
                // Formula: =IF(D{row}=TRUE, SUMIF($A$1:$A${totalRows}, A{row}, $B$1:$B${totalRows}), 0)
                for (int row = 0; row < totalRows; row++)
                {
                    Cell targetCell = cells[row, 2]; // Column C
                    string formula = $"=IF(D{row + 1}=TRUE, SUMIF($A$1:$A${totalRows}, A{row + 1}, $B$1:$B${totalRows}), 0)";
                    targetCell.Formula = formula; // Set formula using the correct property
                }

                // Calculate all formulas so that results are materialized
                workbook.CalculateFormula();

                // Save the workbook
                string outputPath = "ConditionalSumIfDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
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
            ConditionalSumIfRoutine.Run();
        }
    }
}