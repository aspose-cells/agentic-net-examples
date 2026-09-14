// Title: How to auto-fit an Excel column after inserting wrapped multiline text using Aspose.Cells for .NET
// AI Prompts: Generate C# code that inserts a newline-separated string into a cell, turns on text wrapping, and then invokes Worksheet.AutoFitColumn to size the column for that cell using Aspose.Cells. | Provide a .NET example that applies IsTextWrapped to a cell with multiple lines and automatically resizes the containing column to display all lines. | Write a C# snippet that creates an Excel workbook, adds multiline text with wrapping, auto‑fits the column for the first row, and saves the file with Aspose.Cells.
// Common Searches: Aspose.Cells C# auto fit column after wrapping multiline cell content | how to resize Excel column for wrapped text using Aspose.Cells .NET | Worksheet.AutoFitColumn for specific rows with line breaks example | save workbook with column auto‑sized to show line breaks Aspose.Cells | enable text wrap and adjust column width programmatically in Aspose.Cells
// Tags: column auto sizing after wrap Aspose.Cells | multiline text wrapping in Excel .NET | Worksheet.AutoFitColumn specific rows | Excel column width adjustment for line breaks | Aspose.Cells column auto sizing demo

using System;
using Aspose.Cells;

// // Demonstrates inserting multiline text into cell A1, enabling text wrapping, auto‑fitting column A for the first row, and saving the workbook as AutoFitColumnMultiline.xlsx using Aspose.Cells for .NET.
class AutoFitColumnMultilineDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert multiline text into cell A1
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("First line\nSecond line\nThird line");

        // Enable text wrapping for the cell so the text occupies multiple lines
        Style style = cell.GetStyle();
        style.IsTextWrapped = true;
        cell.SetStyle(style);

        // Auto-fit column A (index 0) for the rows that contain the multiline text
        // Here we autofit for row 0 (first row) only; adjust range as needed
        worksheet.AutoFitColumn(0, 0, 0);

        // Save the workbook
        workbook.Save("AutoFitColumnMultiline.xlsx");
    }
}
