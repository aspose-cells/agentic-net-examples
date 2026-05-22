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

        // Populate three rows of data in column A
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["A3"].PutValue(30);

        // Set a formula in B1 that sums the three values above
        cells["B1"].Formula = "=SUM(A1:A3)";

        // Calculate the initial formula result
        workbook.CalculateFormula();

        Console.WriteLine("Before copying rows:");
        Console.WriteLine($"B1 value = {cells["B1"].Value}"); // Expected 60

        // Insert two blank rows after the original data to make space for the copy
        // Row index is zero‑based; inserting at index 3 creates rows 4 and 5 (A4, A5)
        cells.InsertRows(3, 2);

        // Copy the three source rows (0‑2) to the new destination starting at row index 3
        // This will duplicate the rows and adjust any relative references in formulas
        cells.CopyRows(cells, 0, 3, 3);

        // Recalculate formulas after the copy operation
        workbook.CalculateFormula();

        Console.WriteLine("\nAfter copying rows:");
        // Verify that the formula in the copied row (B4) has been updated to reference the new range
        Console.WriteLine($"B4 formula = {cells["B4"].Formula}"); // Expected "=SUM(A4:A6)"
        Console.WriteLine($"B4 value   = {cells["B4"].Value}");   // Expected 60 (same sum as original)

        // Save the workbook to verify the result in Excel
        workbook.Save("CopyRowsFormulaDemo.xlsx");
    }
}