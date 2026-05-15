using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable multi‑threaded reading for the cells collection.
        // This allows concurrent read access and can speed up large data processing.
        workbook.Worksheets[0].Cells.MultiThreadReading = true;

        // Enable the calculation chain to improve performance when formulas are
        // recalculated repeatedly after small changes.
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Populate a large data set with sample values and formulas.
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;
        int totalRows = 5000; // example large number of rows

        for (int i = 0; i < totalRows; i++)
        {
            // Simple numeric value
            cells[i, 0].PutValue(i + 1);

            // Formula that depends on the value in column A
            cells[i, 1].Formula = $"=A{i + 1}*2";
        }

        // Calculate all formulas in the workbook.
        // The multi‑threaded reading setting helps the engine read cell data faster.
        workbook.CalculateFormula();

        // Save the workbook to a file.
        workbook.Save("MultiThreadedCalculation.xlsx", SaveFormat.Xlsx);
    }
}