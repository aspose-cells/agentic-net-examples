// Title: How to confirm that a converted Excel workbook has no ListObject tables using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells verify workbook contains no ListObjects after conversion | Check for Excel table objects (ListObject) in a workbook with Aspose.Cells .NET | Throw exception if a worksheet still has ListObjects after converting to .xlsx using Aspose | Iterate all worksheets and ensure ListObject count is zero with Aspose.Cells C# | Validate removal of Excel tables after file conversion using Aspose.Cells
// Tags: Aspose.Cells verify ListObjects removal | C# iterate worksheets ListObjects count | validate Excel tables after conversion Aspose | exception on remaining ListObject Aspose.Cells | check workbook for ListObject tables .NET

using Aspose.Cells;
using System;

// Loads a converted Excel file with Aspose.Cells, loops through each worksheet, and throws an InvalidOperationException if any ListObject tables are found; otherwise it logs a success message confirming no ListObjects remain.
class Program
{
    static void Main()
    {
        // Load the workbook after conversion
        Workbook workbook = new Workbook("ConvertedWorkbook.xlsx");

        // Iterate through all worksheets and verify that no ListObjects exist
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.ListObjects.Count > 0)
            {
                throw new InvalidOperationException(
                    $"Worksheet '{sheet.Name}' contains {sheet.ListObjects.Count} ListObject(s) after conversion.");
            }
        }

        // If no exception was thrown, validation succeeded
        Console.WriteLine("Validation passed: No ListObject objects found in the workbook.");
    }
}
