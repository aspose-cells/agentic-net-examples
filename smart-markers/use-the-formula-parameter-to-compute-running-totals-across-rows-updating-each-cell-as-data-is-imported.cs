using System;
using Aspose.Cells;

class RunningTotalExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Sample data to import (e.g., sales amounts)
        double[] data = { 120.5, 250.0, 75.25, 300.0, 180.75 };

        // Header row
        cells[0, 0].PutValue("Amount");
        cells[0, 1].PutValue("Running Total");

        // Import data row by row and set running total formula
        for (int i = 0; i < data.Length; i++)
        {
            int rowIndex = i + 1; // +1 because row 0 is header

            // Put the raw amount value in column A
            cells[rowIndex, 0].PutValue(data[i]);

            // Set formula in column B to calculate running total up to the current row
            // The formula uses an absolute start reference ($A$2) and a relative end reference (A{row})
            string formula = $"=SUM($A$2:A{rowIndex + 1})";
            cells[rowIndex, 1].Formula = formula;
        }

        // Calculate all formulas so that the running totals are materialized
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("RunningTotal.xlsx");
    }
}