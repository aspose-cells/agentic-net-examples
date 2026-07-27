// Title: Export Excel to HTML with vertically merged cells using Aspose.Cells for .NET
// Description: Shows how to build a workbook, merge cells across several rows, configure HtmlSaveOptions.MergeEmptyTdType to keep the merge intact, and save the sheet as HTML where the merged area renders as a single table cell.
// Keywords: Aspose.Cells | C# HTML export | merged cells | vertical merge | HtmlSaveOptions | MergeEmptyTdType | Excel to HTML | preserve layout | web report | Aspose.Cells .NET
// Common Searches: Aspose.Cells export merged rows to HTML | HtmlSaveOptions MergeEmptyTdType C# | Save Excel with vertical merge as HTML | How to keep merged cells when converting Excel to HTML | C# convert worksheet with merged cells to HTML | Aspose.Cells HTML output merged cells issue
// Developer Intent: Create an HTML version of an Excel worksheet that retains cells merged across multiple rows.
// Use Cases: Generating web‑ready reports where a header spans several rows | Converting legacy Excel templates with vertical merges for online dashboards | Automating newsletter content from Excel files while preserving layout | Building a documentation portal that displays Excel sheets with exact formatting in browsers
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet containing vertically merged cells to HTML, ensuring the merge is retained. | Explain how HtmlSaveOptions.MergeEmptyTdType influences the rendering of merged rows in the generated HTML. | Provide troubleshooting steps when merged cells appear broken after saving Excel as HTML with Aspose.Cells. | Show how to customize the HTML output (styles, table attributes) while preserving merged cells.

using System;
using Aspose.Cells;

// Shows how to build a workbook, merge cells across several rows, configure HtmlSaveOptions.MergeEmptyTdType to keep the merge intact, and save the sheet as HTML where the merged area renders as a single table cell.
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

        // Merge cells A1:A3 (spanning multiple rows)
        // Parameters: startRow, startColumn, totalRows, totalColumns
        worksheet.Cells.Merge(0, 0, 3, 1);

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Ensure empty TD elements are merged in the same way Excel does
            MergeEmptyTdType = MergeEmptyTdType.Default
        };

        // Save the workbook as HTML
        workbook.Save("MergedCells.html", htmlOptions);
    }
}
