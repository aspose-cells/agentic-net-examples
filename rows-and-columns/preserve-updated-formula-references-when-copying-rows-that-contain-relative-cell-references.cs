// Title: Copy a worksheet row while preserving relative formula references using Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy row keep relative formula references in C# | how to duplicate a row and update formulas automatically with Aspose.Cells | Cells.CopyRows adjusts relative formulas example C# | preserve Excel formula references when copying rows using Aspose.Cells .NET | copy rows and recalculate formulas Aspose.Cells tutorial
// Tags: Cells.CopyRows relative formula adjustment | preserve Excel formulas during row copy .NET | update formula references after row duplication Aspose.Cells | copy rows with automatic formula recalculation Aspose.Cells | adjust relative cell references when copying rows

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to copy a row in an Excel workbook using Aspose.Cells for .NET while automatically updating relative formulas (e.g., "=A1+B1" becomes "=A4+B4"), recalculate the worksheet, and save the result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // 1. Populate the first two rows with data and relative formulas
            // ------------------------------------------------------------
            // Row 1 (index 0)
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);
            cells["C1"].Formula = "=A1+B1"; // relative reference

            // Row 2 (index 1)
            cells["A2"].PutValue(5);
            cells["B2"].PutValue(7);
            cells["C2"].Formula = "=A2+B2"; // relative reference

            // Calculate the formulas so we can display initial values
            workbook.CalculateFormula();

            Console.WriteLine("=== Before copying rows ===");
            PrintRow(cells, 0); // Row 1
            PrintRow(cells, 1); // Row 2

            // ------------------------------------------------------------
            // 2. Copy the first row (row index 0) to a new location (row index 3)
            //    Cells.CopyRows updates relative formula references automatically.
            // ------------------------------------------------------------
            // Overload requires the source Cells collection as the first argument.
            cells.CopyRows(cells, 0, 3, 1); // copy 1 row from row 1 to row 4

            // After copying, the formula in C4 should be adjusted to "=A4+B4"
            // Set new values in the destination row to see the formula recalculate
            cells["A4"].PutValue(30);
            cells["B4"].PutValue(40);

            // Recalculate formulas after inserting new data
            workbook.CalculateFormula();

            Console.WriteLine("\n=== After copying row 1 to row 4 and updating values ===");
            PrintRow(cells, 3); // Row 4 (the copied row)

            // Show the exact formula text stored in the copied cell
            Console.WriteLine($"Formula in C4: {cells["C4"].Formula}");
            Console.WriteLine($"Calculated value in C4: {cells["C4"].StringValue}");

            // ------------------------------------------------------------
            // 3. Save the workbook (optional – allows visual verification)
            // ------------------------------------------------------------
            string outputPath = "CopyRowsPreserveFormula.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    /// <param name="cells">The Cells collection.</param>
    /// <param name="rowIndex">Zero‑based row index.</param>
    static void PrintRow(Cells cells, int rowIndex)
    {
        string a = cells[rowIndex, 0].StringValue;
        string b = cells[rowIndex, 1].StringValue;
        string c = cells[rowIndex, 2].StringValue;
        string formula = cells[rowIndex, 2].Formula;

        Console.WriteLine($"Row {rowIndex + 1}: A={a}, B={b}, C={c} (Formula={formula})");
    }
}
