using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable multi‑threaded reading for the cells collection.
        // This allows the calculation engine to read cell data concurrently,
        // improving performance on large data sets.
        workbook.Worksheets[0].Cells.MultiThreadReading = true;

        // Populate sample data and formulas
        Cells cells = workbook.Worksheets[0].Cells;
        for (int i = 0; i < 1000; i++)
        {
            // Simple numeric values
            cells[i, 0].PutValue(i + 1);
            // Formula that depends on the value in column A
            cells[i, 1].Formula = $"=A{i + 1}*2";
        }

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("MultiThreadedCalculation.xlsx", SaveFormat.Xlsx);
    }
}