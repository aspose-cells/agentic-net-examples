// Title: Adjust Column Width with AutoFitColumns After CopyRows in Aspose.Cells for .NET
// Description: C# example that copies rows from a source worksheet to a destination workbook using Cells.CopyRows, then automatically resizes all columns with AutoFitColumns before saving the file.
// Keywords: Aspose.Cells | AutoFitColumns | CopyRows | C# | .NET | adjust column width | copy rows between worksheets | column auto‑size after copy | Excel automation
// Common Searches: Aspose.Cells AutoFitColumns after CopyRows | how to resize columns after copying rows in .NET | C# copy rows and auto‑fit columns Aspose | adjust column widths in destination sheet Aspose.Cells | auto size columns after Cells.CopyRows
// Developer Intent: Automatically resize every column in the target worksheet so that the data copied with Cells.CopyRows fits without truncation.
// Use Cases: Migrate a header row and data rows to a new report workbook while preserving readable column widths. | Create a printable Excel file by copying table rows from a template and applying AutoFitColumns to ensure proper layout. | Generate a summary sheet that aggregates rows from multiple sources and automatically adjusts column sizes for clarity.
// AI Prompts: Show C# code that copies a range of rows to another worksheet and then calls AutoFitColumns only on columns that contain data using Aspose.Cells. | Explain how to preserve original column widths when copying rows and later re‑apply AutoFitColumns for newly added content. | Provide an Aspose.Cells example that copies rows, handles merged cells, and auto‑fits a specific column range in .NET.

using System;
using Aspose.Cells;

// C# example that copies rows from a source worksheet to a destination workbook using Cells.CopyRows, then automatically resizes all columns with AutoFitColumns before saving the file.
class AdjustColumnWidthAfterCopyRows
{
    static void Main()
    {
        // Create a source workbook and populate it with sample data
        Workbook srcWorkbook = new Workbook();
        Worksheet srcSheet = srcWorkbook.Worksheets[0];
        srcSheet.Cells["A1"].PutValue("Header");
        srcSheet.Cells["A2"].PutValue("Short");
        srcSheet.Cells["A3"].PutValue("A very long text that will need column widening");
        srcSheet.Cells["B1"].PutValue("Number");
        srcSheet.Cells["B2"].PutValue(123);
        srcSheet.Cells["B3"].PutValue(456789);

        // Create a destination workbook where rows will be copied to
        Workbook destWorkbook = new Workbook();
        Worksheet destSheet = destWorkbook.Worksheets[0];

        // Copy the first three rows from the source sheet to the destination sheet
        // Parameters: source cells, source start row, destination start row, number of rows to copy
        destSheet.Cells.CopyRows(srcSheet.Cells, 0, 0, 3);

        // After copying, auto‑fit all columns in the destination worksheet
        destSheet.AutoFitColumns();

        // Save the resulting workbook
        destWorkbook.Save("AdjustedColumnsAfterCopyRows.xlsx");
    }
}
