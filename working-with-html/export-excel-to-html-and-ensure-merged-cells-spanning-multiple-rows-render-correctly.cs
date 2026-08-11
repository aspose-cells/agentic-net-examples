// Title: Export Excel to HTML with vertically merged cells – Aspose.Cells for .NET
// Description: This C# sample builds a workbook, merges cells A1:A3 vertically and B1:C2 across rows and columns, configures HtmlSaveOptions.MergeEmptyTdType to MergeForcely, and saves the result as MergedCells.html so that merged cells spanning multiple rows appear correctly in the HTML output.
// Keywords: Aspose.Cells | C# | HTML export | merged cells | vertical merge | MergeEmptyTdType | .NET | Excel to HTML | preserve cell merges | HtmlSaveOptions
// Common Searches: Aspose.Cells export Excel to HTML merged cells | how to keep vertical merged cells in HTML output | HtmlSaveOptions MergeEmptyTdType example | C# save workbook as HTML with merged ranges | merged cells not showing in HTML Aspose.Cells
// Developer Intent: Generate an HTML file from an Excel workbook while ensuring that cells merged across multiple rows are rendered correctly.
// Use Cases: Create web‑ready reports where a header spans several rows | Convert Excel tables with merged blocks into responsive HTML emails | Publish Excel‑based dashboards on a website without losing layout
// AI Prompts: Show how to set HtmlSaveOptions.MergeEmptyTdType to preserve vertical merged cells when exporting to HTML with Aspose.Cells. | Provide C# code that exports a workbook with merged ranges and forces empty TD elements to merge. | Explain why MergeForcely is needed for correct HTML rendering of multi‑row merged cells.

using System;
using Aspose.Cells;

// This C# sample builds a workbook, merges cells A1:A3 vertically and B1:C2 across rows and columns, configures HtmlSaveOptions.MergeEmptyTdType to MergeForcely, and saves the result as MergedCells.html so that merged cells spanning multiple rows appear correctly in the HTML output.
class ExportMergedCellsHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some data
        worksheet.Cells["A1"].PutValue("Header");
        worksheet.Cells["A2"].PutValue("Row 1");
        worksheet.Cells["A3"].PutValue("Row 2");

        // Merge cells vertically (A1:A3) – this is a merged cell spanning multiple rows
        worksheet.Cells.CreateRange("A1", "A3").Merge();

        // Another merged region across rows and columns (B1:C2)
        worksheet.Cells["B1"].PutValue("Block");
        worksheet.Cells.CreateRange("B1", "C2").Merge();

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Force merging of empty TD elements so that merged cells render correctly in HTML
        htmlOptions.MergeEmptyTdType = MergeEmptyTdType.MergeForcely;

        // Save the workbook as HTML
        workbook.Save("MergedCells.html", htmlOptions);
    }
}
