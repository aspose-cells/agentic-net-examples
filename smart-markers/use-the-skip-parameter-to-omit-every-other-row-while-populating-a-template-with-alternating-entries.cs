// Title: How to import a vertical array into an Excel sheet and skip every other row with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Workbook, uses Worksheet.Cells.ImportObjectArray to import a one‑dimensional object array vertically, and sets the skip argument so a blank row is inserted after each value. | Explain the steps for calling ImportObjectArray with the skip parameter to produce alternating populated and empty rows in an Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells C# import vertical array with blank rows example | ImportObjectArray skip parameter usage .NET | Create Excel file with alternating data rows using Aspose.Cells | Leave an empty row after each record when importing data with Aspose.Cells
// Tags: ImportObjectArray vertical array Aspose.Cells | skip parameter blank rows Aspose.Cells | populate worksheet alternating rows C# | smart markers row skipping Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates creating a Workbook, importing a vertical object array into cell A1 with ImportObjectArray using a skip value of 1 to leave a blank row after each entry, and saving the result as SkipRowsDemo.xlsx.
class SkipRowsDemo
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample data to be placed in the worksheet
        // The array contains six values; they will be written vertically
        object[] data = new object[] { "Alice", "Bob", "Charlie", "David", "Eve", "Frank" };

        // Import the data vertically starting at cell A1 (row 0, column 0)
        // The 'skip' parameter is set to 1, which means after each value a blank row is left,
        // effectively omitting every other row in the output.
        // Parameters: (object[] data, int firstRow, int firstColumn, bool isVertical, int skip)
        worksheet.Cells.ImportObjectArray(data, 0, 0, true, 1);

        // Save the workbook (lifecycle save)
        workbook.Save("SkipRowsDemo.xlsx");
    }
}
