// Title: Auto‑Fit Columns Before Freezing Panes with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills cells with varied text, calls AutoFitColumns to size every column, then freezes the top two rows and left two columns at cell C3, preserving column widths, and saves the file as AutoFitAndFreeze.xlsx.
// Keywords: Aspose.Cells AutoFitColumns C# | FreezePanes preserve column width | auto fit then freeze panes Aspose.Cells | C# spreadsheet column sizing | Aspose.Cells column width retention
// Common Searches: auto fit columns before freeze panes Aspose.Cells | preserve column width when using FreezePanes .NET | C# Aspose.Cells AutoFitColumns example | why call AutoFitColumns before FreezePanes | freeze top rows and columns after auto fit
// Developer Intent: Automatically size all worksheet columns based on content, then lock header rows and columns without altering the calculated widths.
// Use Cases: Generating reports where column widths must stay consistent after freezing header rows/columns. | Exporting data to Excel templates that require optimal column sizing before applying FreezePanes. | Creating reusable spreadsheet layouts that keep column dimensions intact during user navigation.
// AI Prompts: Write C# code using Aspose.Cells to auto‑fit every column, then freeze the first two rows and columns while keeping the column widths unchanged. | Explain the importance of calling AutoFitColumns before FreezePanes in Aspose.Cells and show a short example. | Provide step‑by‑step instructions to adjust column widths, apply FreezePanes, and save the workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a workbook, fills cells with varied text, calls AutoFitColumns to size every column, then freezes the top two rows and left two columns at cell C3, preserving column widths, and saves the file as AutoFitAndFreeze.xlsx.
class AutoFitAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Short");
        sheet.Cells["B1"].PutValue("This is a longer text that should cause column B to expand");
        sheet.Cells["C1"].PutValue("Medium length");
        sheet.Cells["A2"].PutValue("Another row with a very very long text that will affect column A width");
        sheet.Cells["B2"].PutValue(12345);
        sheet.Cells["C2"].PutValue(DateTime.Now);

        // Auto‑fit all columns before freezing panes to preserve column widths
        sheet.AutoFitColumns();

        // Freeze panes at cell C3 (row index 2, column index 2) with 2 rows and 2 columns frozen
        sheet.FreezePanes(2, 2, 2, 2);

        // Save the workbook
        workbook.Save("AutoFitAndFreeze.xlsx");
    }
}
