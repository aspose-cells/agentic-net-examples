// Title: C# – Freeze Panes After AutoFitColumns to Preserve Column Widths (Aspose.Cells)
// Description: This example creates a workbook, fills cells with varied text, auto‑fits all columns to the content, then calls Worksheet.FreezePanes to lock the calculated widths while freezing the first two rows and columns. The workbook is saved as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | Worksheet.AutoFitColumns | Worksheet.FreezePanes | lock column width | preserve column width | freeze panes programmatically | Excel column auto fit | freeze top rows and left columns
// Common Searches: Aspose.Cells freeze panes after AutoFitColumns | How to lock column width after auto‑fit in C# | Freeze first two rows and columns with Aspose.Cells | Preserve column widths when freezing panes .NET | AutoFitColumns then FreezePanes example
// Developer Intent: Preserve auto‑fitted column widths while freezing rows and columns.
// Use Cases: Generate a report where columns are auto‑fitted once, then freeze header rows and left columns to keep the layout stable during scrolling. | Create a spreadsheet template that auto‑fits column widths on export and locks them by freezing the top rows and left columns for end‑user editing. | Export a data table to Excel, auto‑fit all columns, freeze the first two rows and columns, and deliver the file to users without layout changes.
// AI Prompts: Show C# code that auto‑fits all columns in a worksheet and then freezes panes at a specific cell using Aspose.Cells. | Provide an Aspose.Cells example where Worksheet.AutoFitColumns is called before Worksheet.FreezePanes to lock column widths. | Explain why FreezePanes should be invoked after AutoFitColumns when you need to preserve column widths in an Excel workbook.

using System;
using Aspose.Cells;

namespace FreezePanesAfterAutoFit
{
    // This example creates a workbook, fills cells with varied text, auto‑fits all columns to the content, then calls Worksheet.FreezePanes to lock the calculated widths while freezing the first two rows and columns. The workbook is saved as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Short");
            sheet.Cells["B1"].PutValue("A much longer text that will cause the column to expand");
            sheet.Cells["C1"].PutValue("Medium length");
            sheet.Cells["A2"].PutValue("Another short");
            sheet.Cells["B2"].PutValue("Another very long piece of text to demonstrate auto‑fit functionality");
            sheet.Cells["C2"].PutValue("Text");

            // Auto‑fit all columns so their widths match the content
            sheet.AutoFitColumns();

            // Freeze panes at row 2, column 2 (C3 cell) with 2 rows and 2 columns frozen
            // This locks the column widths after they have been auto‑fitted
            sheet.FreezePanes(2, 2, 2, 2);

            // Save the workbook to a file
            workbook.Save("FreezePanesAfterAutoFit.xlsx");
        }
    }
}
