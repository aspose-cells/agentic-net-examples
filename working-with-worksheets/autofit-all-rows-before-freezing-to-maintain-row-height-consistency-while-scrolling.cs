// Title: Auto‑Fit Rows Then Freeze Panes in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, writes short and long text (including wrapped multi‑line cells), applies text wrapping, auto‑fits all rows to adjust heights, freezes the first three rows and columns, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | AutoFitRows | FreezePanes | row height | text wrap | Excel export | worksheet formatting
// Common Searches: Aspose.Cells auto fit rows C# | Freeze panes after AutoFitRows Aspose.Cells | Preserve wrapped text height when freezing panes | How to auto fit rows before FreezePanes in .NET | AutoFitRows effect on frozen rows Aspose.Cells
// Developer Intent: Automatically adjust row heights before applying FreezePanes so that the visual layout remains consistent while scrolling.
// Use Cases: Generating reports with wrapped text where header rows and columns are frozen for easy navigation. | Exporting data to Excel with multi‑line cells, ensuring rows are sized correctly before pane freezing. | Creating printable spreadsheets that retain calculated row heights after freezing panes.
// AI Prompts: Provide C# code that wraps text, auto‑fits all rows, then freezes the first three rows and columns using Aspose.Cells. | Explain why AutoFitRows should be called before FreezePanes in Aspose.Cells for .NET and how it preserves row height consistency. | Show a step‑by‑step example of maintaining wrapped‑text row heights while freezing panes in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitAndFreezeDemo
{
    // Creates a workbook, writes short and long text (including wrapped multi‑line cells), applies text wrapping, auto‑fits all rows to adjust heights, freezes the first three rows and columns, and saves the file as an Excel workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data that will affect row heights
            worksheet.Cells["A1"].PutValue("Short text");
            worksheet.Cells["A2"].PutValue("This is a longer piece of text that should cause the row to expand when auto‑fitted.");
            worksheet.Cells["A3"].PutValue("Another line with\nmultiple line breaks\nto demonstrate row height adjustment.");

            // Apply text wrapping to demonstrate multi‑line row height changes
            Style wrapStyle = worksheet.Cells["A3"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            worksheet.Cells["A3"].SetStyle(wrapStyle);

            // Auto‑fit all rows in the worksheet before freezing panes
            worksheet.AutoFitRows();

            // Freeze panes at cell C4 (row index 3, column index 3) with 3 rows and 3 columns frozen
            worksheet.FreezePanes(3, 3, 3, 3);

            // Save the workbook
            workbook.Save("AutoFitRowsAndFreezePanes.xlsx");
        }
    }
}
