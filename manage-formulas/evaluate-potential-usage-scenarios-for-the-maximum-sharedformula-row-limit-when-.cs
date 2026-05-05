using System;
using Aspose.Cells;

class MaxRowsOfSharedFormulaDemo
{
    static void Main()
    {
        // ------------------------------------------------------------
        // Scenario 1: Create a workbook, set a low MaxRowsOfSharedFormula,
        // and observe that shared formulas beyond this limit are not applied.
        // ------------------------------------------------------------
        Workbook wb1 = new Workbook();

        // Limit shared formulas to the first 100 rows.
        wb1.Settings.MaxRowsOfSharedFormula = 100;

        Worksheet ws1 = wb1.Worksheets[0];
        Cells cells1 = ws1.Cells;

        // Fill column A with values 1..200.
        for (int i = 0; i < 200; i++)
        {
            cells1[i, 0].PutValue(i + 1);
        }

        // Attempt to set a shared formula for 200 rows (exceeds the limit of 100).
        cells1["B1"].SetSharedFormula("=A1*2", 200, 1);

        // Calculate formulas.
        wb1.CalculateFormula();

        // Row 150 is beyond the limit, so its formula/value will be empty or default.
        Console.WriteLine("Scenario 1 - Limit 100 rows:");
        Console.WriteLine($"B150 Formula: {cells1["B150"].Formula}");
        Console.WriteLine($"B150 Value: {cells1["B150"].Value}");

        // ------------------------------------------------------------
        // Scenario 1 (continued): Increase the limit and verify that the
        // same shared formula now works for all rows.
        // ------------------------------------------------------------
        wb1.Settings.MaxRowsOfSharedFormula = 500; // raise limit

        // Add a new sheet to demonstrate the increased limit.
        Worksheet ws2 = wb1.Worksheets.Add("Expanded");
        Cells cells2 = ws2.Cells;

        // Fill column A again.
        for (int i = 0; i < 200; i++)
        {
            cells2[i, 0].PutValue(i + 1);
        }

        // Set the same shared formula for 200 rows.
        cells2["B1"].SetSharedFormula("=A1*2", 200, 1);
        wb1.CalculateFormula();

        Console.WriteLine("Scenario 1 - Limit increased to 500 rows:");
        Console.WriteLine($"B150 Formula: {cells2["B150"].Formula}");
        Console.WriteLine($"B150 Value: {cells2["B150"].Value}");

        // Save the workbook for the next scenario.
        string filePath = "SharedFormulaDemo.xlsx";
        wb1.Save(filePath);

        // ------------------------------------------------------------
        // Scenario 2: Load an existing workbook that contains a large
        // shared‑formula range, then adjust MaxRowsOfSharedFormula before
        // calculation to control how many rows actually receive the formula.
        // ------------------------------------------------------------
        LoadOptions loadOpts = new LoadOptions(); // default load options
        Workbook wb2 = new Workbook(filePath, loadOpts);

        // Initially set a restrictive limit (e.g., 120 rows).
        wb2.Settings.MaxRowsOfSharedFormula = 120;
        wb2.CalculateFormula();

        Console.WriteLine("Scenario 2 - Loaded workbook with limit 120 rows:");
        Console.WriteLine($"Sheet1 B150 Formula after load: {wb2.Worksheets[0].Cells["B150"].Formula}");
        Console.WriteLine($"Sheet1 B150 Value after load: {wb2.Worksheets[0].Cells["B150"].Value}");

        // Increase the limit to cover the full range and recalculate.
        wb2.Settings.MaxRowsOfSharedFormula = 300;
        wb2.CalculateFormula();

        Console.WriteLine("Scenario 2 - Limit increased to 300 rows:");
        Console.WriteLine($"Sheet1 B150 Formula after increase: {wb2.Worksheets[0].Cells["B150"].Formula}");
        Console.WriteLine($"Sheet1 B150 Value after increase: {wb2.Worksheets[0].Cells["B150"].Value}");

        // Save the modified workbook.
        wb2.Save("SharedFormulaDemo_Modified.xlsx");
    }
}