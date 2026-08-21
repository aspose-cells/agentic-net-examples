// Title: Auto‑Fit Excel Column Widths Using Aspose.Cells for .NET (C#)
// Description: Shows how to create a Workbook, add short, medium and long text plus numeric values, invoke Worksheet.AutoFitColumns() to automatically size all columns to the widest entry, and save the result as AutoFitColumnsResult.xlsx.
// Keywords: Aspose.Cells AutoFitColumns | C# Excel column auto fit | adjust Excel column width programmatically | auto resize columns Aspose .NET | worksheet.AutoFitColumns example | Excel column width based on content
// Common Searches: Aspose.Cells auto fit column width C# | How to auto‑fit columns in an Excel file using Aspose.Cells | Worksheet.AutoFitColumns method usage | Resize Excel columns to fit content with .NET | C# code to automatically adjust column width in Aspose.Cells
// Developer Intent: Automatically resize every column in a worksheet so the longest cell value fits without truncation.
// Use Cases: Generating reports with variable‑length strings where each column must display full text. | Exporting numeric datasets that include large numbers requiring wider columns. | Creating a template, populating data via code, applying AutoFitColumns, and delivering a ready‑to‑view workbook.
// AI Prompts: Write C# code that uses Aspose.Cells to auto‑fit all columns after filling a worksheet with mixed text and numbers. | Explain the behavior of Worksheet.AutoFitColumns and how to limit its effect to specific columns. | Show how to combine AutoFitColumns with a maximum column width constraint in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AutoFitColumnDemo
{
    // Shows how to create a Workbook, add short, medium and long text plus numeric values, invoke Worksheet.AutoFitColumns() to automatically size all columns to the widest entry, and save the result as AutoFitColumnsResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells with varying length text to demonstrate auto‑fit
            worksheet.Cells["A1"].PutValue("Short");
            worksheet.Cells["A2"].PutValue("This is a considerably longer piece of text that should cause the column to expand.");
            worksheet.Cells["B1"].PutValue("Medium length text");
            worksheet.Cells["B2"].PutValue("Another long text entry to test column width adjustment automatically.");
            worksheet.Cells["C1"].PutValue(12345);
            worksheet.Cells["C2"].PutValue(9876543210);

            // Auto‑fit all columns based on the content (auto‑fit rule)
            worksheet.AutoFitColumns();

            // Save the workbook (save rule)
            workbook.Save("AutoFitColumnsResult.xlsx");
        }
    }
}
