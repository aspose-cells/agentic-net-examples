using System;
using Aspose.Cells;

namespace AsposeCellsMultiThreadCalculation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a large data set with formulas to demonstrate performance
            int rows = 5000;
            int cols = 10;

            // Fill cells with numeric values
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    cells[r, c].PutValue(r + c);
                }
            }

            // Add a formula that sums a row – this will be calculated for each row
            for (int r = 0; r < rows; r++)
            {
                // Example: sum of the first 10 columns in the current row
                cells[r, cols].Formula = $"=SUM(A{r + 1}:J{r + 1})";
            }

            // Enable multi‑threaded processing for the cells data model.
            // This allows Aspose.Cells to evaluate formulas using multiple threads,
            // which speeds up calculation on large data sets.
            cells.MultiThreadReading = true;

            // Calculate all formulas in the workbook.
            // The calculation will take advantage of the multi‑thread setting above.
            workbook.CalculateFormula();

            // Save the workbook to verify the results
            workbook.Save("MultiThreadCalculationResult.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved with multi‑threaded calculation enabled.");
        }
    }
}