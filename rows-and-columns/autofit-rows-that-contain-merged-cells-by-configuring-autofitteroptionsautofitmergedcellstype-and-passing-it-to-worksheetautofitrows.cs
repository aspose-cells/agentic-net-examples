// Title: Auto‑Fit Rows with Merged Cells Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to merge cells, enable text wrapping, configure AutoFitterOptions (AutoFitMergedCellsType = EachLine, AutoFitWrappedTextType = Paragraph) and call Worksheet.AutoFitRows to automatically adjust row heights so wrapped text in merged ranges displays correctly, then saves the workbook.
// Keywords: Aspose.Cells AutoFitRows C# | merged cells row height | AutoFitMergedCellsType EachLine example | AutoFitWrappedTextType Paragraph | auto adjust row height .NET | Excel merged cells wrap text | Aspose.Cells tutorial | C# Excel automation
// Common Searches: how to auto‑fit rows with merged cells asp.net | Aspose.Cells AutoFitRows merged cells settings | C# auto adjust row height for wrapped merged cells | AutoFitterOptions example for merged cells | fit each line of merged cell text in Excel using Aspose
// Developer Intent: Resize rows that contain merged cells so that wrapped, multi‑line text fits without truncation.
// Use Cases: Generating a report where a title spans several columns and rows and must expand to show all wrapped lines. | Creating invoices with a merged header cell that holds a lengthy description and needs automatic row‑height adjustment. | Exporting data that includes multi‑line comments in merged cells, requiring rows to resize for full visibility.
// AI Prompts: Provide C# code that uses Aspose.Cells to auto‑fit rows containing merged cells with wrapped text. | Show how to set AutoFitMergedCellsType to EachLine and AutoFitWrappedTextType to Paragraph for merged ranges. | Explain how to modify the example to auto‑fit columns instead of rows while preserving merged‑cell handling.

using System;
using Aspose.Cells;

// Demonstrates how to merge cells, enable text wrapping, configure AutoFitterOptions (AutoFitMergedCellsType = EachLine, AutoFitWrappedTextType = Paragraph) and call Worksheet.AutoFitRows to automatically adjust row heights so wrapped text in merged ranges displays correctly, then saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add long text that will need wrapping and row height adjustment
        worksheet.Cells["A1"].PutValue(
            "This is a sample text for merged cells auto‑fit demonstration. " +
            "It contains enough content to require multiple lines when wrapped.");

        // Merge cells A1:B3 (rows 0‑2, columns 0‑1)
        worksheet.Cells.Merge(0, 0, 3, 2);

        // Enable text wrapping for the merged cell
        Style style = worksheet.Cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        worksheet.Cells["A1"].SetStyle(style);

        // Configure AutoFitterOptions to auto‑fit each line of merged cells
        AutoFitterOptions options = new AutoFitterOptions
        {
            AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
            AutoFitWrappedTextType = AutoFitWrappedTextType.Paragraph
        };

        // Auto‑fit rows using the configured options
        worksheet.AutoFitRows(options);

        // Save the workbook
        workbook.Save("AutoFitMergedCellsDemo.xlsx");
    }
}
