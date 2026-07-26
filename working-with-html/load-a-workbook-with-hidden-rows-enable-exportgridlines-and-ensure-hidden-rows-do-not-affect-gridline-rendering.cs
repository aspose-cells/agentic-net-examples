// Title: Aspose.Cells .NET – Export Excel to HTML with Gridlines and Exclude Hidden Rows
// Description: Load a workbook, enable worksheet gridlines, hide rows as needed, and use HtmlSaveOptions with ExportGridLines and HiddenRowDisplayType.Remove to generate HTML where hidden rows are omitted and gridlines render correctly.
// Keywords: Aspose.Cells HTML export | ExportGridLines | HiddenRowDisplayType.Remove | remove hidden rows HTML | gridlines visible Excel to HTML | C# Aspose.Cells example
// Common Searches: Aspose.Cells export gridlines to HTML | How to hide rows in HTML output with Aspose.Cells | HtmlSaveOptions HiddenRowDisplayType example | C# export Excel as HTML without hidden rows
// Developer Intent: Create an HTML representation of an Excel worksheet that shows gridlines exactly as in Excel while stripping out any hidden rows.
// Use Cases: Web preview of spreadsheets where hidden rows must not appear but gridlines are required. | Generating clean HTML tables for dashboards or reports from Excel files. | Automating conversion of Excel workbooks to HTML for documentation while preserving layout integrity.
// AI Prompts: Generate C# code using Aspose.Cells to save an Excel file as HTML with visible gridlines and hidden rows removed. | Explain the effect of HtmlHiddenRowDisplayType.Remove on the HTML output when exporting with Aspose.Cells. | Provide a step‑by‑step tutorial for configuring HtmlSaveOptions to export gridlines and exclude hidden rows.

using System;
using Aspose.Cells;

// Load a workbook, enable worksheet gridlines, hide rows as needed, and use HtmlSaveOptions with ExportGridLines and HiddenRowDisplayType.Remove to generate HTML where hidden rows are omitted and gridlines render correctly.
class Program
{
    static void Main()
    {
        // Load an existing workbook that contains hidden rows
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Make sure gridlines are visible in the worksheet
        worksheet.IsGridlinesVisible = true;

        // (Optional) Hide a specific row to demonstrate the effect
        // This row will be excluded from the HTML output, so it won't break gridlines
        worksheet.Cells.HideRow(2); // hides the third row (0‑based index)

        // Configure HTML save options:
        // - ExportGridLines: export the gridlines as they are visible
        // - HiddenRowDisplayType: Remove hidden rows so they don't affect gridline rendering
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportGridLines = worksheet.IsGridlinesVisible,
            HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove
        };

        // Save the workbook to HTML with the specified options
        workbook.Save("output.html", saveOptions);
    }
}
