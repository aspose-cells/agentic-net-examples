// Title: Export Excel to HTML with Gridlines and Hidden Rows Removed – Aspose.Cells for .NET
// Description: Loads a workbook, hides a specific row, enables worksheet gridlines, sets HtmlSaveOptions to export gridlines and remove hidden rows, then saves the result as an HTML file. Hidden rows are excluded so they do not disrupt the gridline layout.
// Keywords: Aspose.Cells | C# | .NET | HTML export | ExportGridLines | HiddenRowDisplayType.Remove | remove hidden rows | gridlines in HTML | Excel to HTML conversion | Aspose.Cells HtmlSaveOptions
// Common Searches: Aspose.Cells export gridlines to HTML | How to hide rows when exporting Excel to HTML with Aspose | Remove hidden rows from HTML output Aspose.Cells | Export Excel with visible gridlines C# | HtmlSaveOptions HiddenRowDisplayType example
// Developer Intent: Create an HTML representation of an Excel workbook that displays gridlines while automatically omitting any hidden rows, using Aspose.Cells for .NET.
// Use Cases: Generating web‑ready reports where gridlines improve readability but hidden rows must be excluded. | Embedding a spreadsheet view in a portal without showing rows that are concealed in the original file. | Preparing printable HTML output where hidden rows could break the visual grid structure. | Automating conversion of Excel dashboards to HTML while preserving layout consistency.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML, enable ExportGridLines, and remove hidden rows from the output. | Explain the effect of HtmlHiddenRowDisplayType.Remove on gridline rendering in an HTML export. | Provide a step‑by‑step tutorial for hiding rows and ensuring gridlines appear correctly in the generated HTML using Aspose.Cells.

using System;
using Aspose.Cells;

// Loads a workbook, hides a specific row, enables worksheet gridlines, sets HtmlSaveOptions to export gridlines and remove hidden rows, then saves the result as an HTML file. Hidden rows are excluded so they do not disrupt the gridline layout.
class ExportGridLinesWithHiddenRows
{
    static void Main()
    {
        // Load an existing workbook that contains data
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide a row (example: row index 2, i.e., third row)
        worksheet.Cells.HideRow(2);

        // Make sure gridlines are visible in the worksheet
        worksheet.IsGridlinesVisible = true;

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Export the gridlines so they appear in the HTML output
            ExportGridLines = true,
            // Remove hidden rows from the HTML to prevent them from breaking gridline rendering
            HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove
        };

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}
