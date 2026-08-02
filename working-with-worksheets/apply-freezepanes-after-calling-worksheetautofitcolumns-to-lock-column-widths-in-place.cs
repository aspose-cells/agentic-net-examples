// Title: Freeze Panes After AutoFitColumns in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to auto‑fit all columns in an Aspose.Cells worksheet, then apply FreezePanes at cell C3 (2 rows, 2 columns) to lock the adjusted column widths, and save the workbook as an Excel file.
// Keywords: Aspose.Cells C# FreezePanes | AutoFitColumns Aspose.Cells | lock column width Aspose.Cells | freeze panes after autofit | Aspose.Cells worksheet example | C# Excel column auto fit | Excel FreezePanes C# Aspose | Aspose.Cells column width lock | worksheet FreezePanes C#
// Common Searches: Aspose.Cells freeze panes after autofit columns | C# AutoFitColumns then FreezePanes | How to lock column widths in Aspose.Cells | Freeze first two rows and columns in Aspose.Cells | Aspose.Cells C# example freeze panes at C3
// Developer Intent: Apply FreezePanes after AutoFitColumns to preserve column widths.
// Use Cases: Export a data grid to Excel, auto‑size columns, then freeze header rows and key columns for easier navigation. | Create a printable report where column widths are calculated once and remain constant when users scroll. | Build a spreadsheet template that automatically adjusts column widths on generation and then locks the layout with FreezePanes. | Develop an Excel export feature in a .NET web app that needs consistent column sizing across different browsers.
// AI Prompts: Show C# code using Aspose.Cells to auto‑fit columns and then freeze the first two rows and columns at cell C3. | Explain why FreezePanes must be set after calling AutoFitColumns in Aspose.Cells. | Provide a step‑by‑step guide to lock column widths after AutoFitColumns with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to auto‑fit all columns in an Aspose.Cells worksheet, then apply FreezePanes at cell C3 (2 rows, 2 columns) to lock the adjusted column widths, and save the workbook as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to demonstrate column width changes
        sheet.Cells["A1"].PutValue("Short");
        sheet.Cells["B1"].PutValue("This is a longer text that will cause column B to expand significantly");
        sheet.Cells["C1"].PutValue("Medium length");
        sheet.Cells["A2"].PutValue("Another row with some text");
        sheet.Cells["B2"].PutValue("More data in column B");
        sheet.Cells["C2"].PutValue("Data");

        // Auto-fit all columns so their widths match the content
        sheet.AutoFitColumns();

        // Freeze panes after auto-fitting to lock the column widths in place
        // Freeze at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns
        sheet.FreezePanes("C3", 2, 2);

        // Save the workbook to a file
        workbook.Save("FreezeAfterAutoFit.xlsx");
    }
}
