using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // ------------------------------------------------------------
        // Set calculation mode to Automatic BEFORE performing any merge.
        // This ensures that formulas are evaluated automatically as
        // cells are changed, preserving their integrity.
        // ------------------------------------------------------------
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Populate some sample data
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);

        // Add a formula that references the cells we will merge
        cells["B1"].Formula = "=SUM(A1:A2)";

        // ------------------------------------------------------------
        // Merge the range A1:A2 (vertical merge). The formula in B1
        // continues to reference the original cells, and because the
        // calculation mode is Automatic, the formula result stays up‑to‑date.
        // ------------------------------------------------------------
        cells.Merge(0, 0, 2, 1); // firstRow=0, firstColumn=0, totalRows=2, totalColumns=1

        // ------------------------------------------------------------
        // Re‑assert the calculation mode after merging to guarantee it
        // remains Automatic (in case any operation altered it).
        // ------------------------------------------------------------
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Save the workbook; formulas remain intact and correctly calculated.
        workbook.Save("MergedFormulas.xlsx");
    }
}