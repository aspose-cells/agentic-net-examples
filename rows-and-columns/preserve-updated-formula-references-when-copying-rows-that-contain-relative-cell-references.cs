using System;
using Aspose.Cells;

class PreserveFormulaCopyRows
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some data in the first row
        cells["A1"].PutValue(10);          // A1 = 10
        cells["B1"].PutValue(20);          // B1 = 20

        // Set a relative formula in C1 that refers to A1 and B1
        cells["C1"].Formula = "=A1+B1";    // C1 = 30 (after calculation)

        // Copy the first row (row index 0) to row index 2 (third row)
        // This will copy data, formats and formulas.
        // Relative references in the copied formula will be adjusted automatically.
        cells.CopyRows(cells, 0, 2, 1);

        // After copying, the formula in C3 should be adjusted to refer to A3 and B3
        Console.WriteLine("Original formula (C1): " + cells["C1"].Formula);
        Console.WriteLine("Copied formula   (C3): " + cells["C3"].Formula);

        // Save the workbook to verify the result
        workbook.Save("PreserveFormulaCopyRows.xlsx");
    }
}