// Title: Add a specific cell with a formula to the Formula Watch Window using Worksheet.CellWatches.Add in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add cell B2 to Formula Watch Window after setting formula | how to use Worksheet.CellWatches.Add to monitor a formula in a .NET workbook | C# example for adding a formula watch for a single cell with Aspose.Cells | track changes of a SUM formula in Excel using Aspose.Cells Formula Watch | add cell to Formula Watch Window programmatically with Aspose.Cells for .NET
// Tags: Worksheet.CellWatches.Add C# example | formula watch window Aspose.Cells | monitor cell formula Aspose.Cells .NET | add cell to formula watch Aspose.Cells | track formula evaluation Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new Workbook, sets a SUM formula in cell B2, adds B2 to the Formula Watch Window via Worksheet.CellWatches.Add, and saves the file as FormulaWatchExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a formula in cell B2 (example: sum of A1:A10)
        Cell cell = sheet.Cells["B2"];
        cell.Formula = "=SUM(A1:A10)";

        // Add the cell to the Formula Watch Window
        sheet.CellWatches.Add("B2");

        // Save the workbook (optional)
        workbook.Save("FormulaWatchExample.xlsx");
    }
}
