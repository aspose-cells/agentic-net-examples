// Title: Export Excel to HTML without Hidden Rows using Aspose.Cells for .NET
// Description: Creates a workbook, hides a specific row, sets HtmlSaveOptions.HiddenRowDisplayType to Remove, and saves the file as HTML so the hidden row is omitted from the markup.
// Keywords: Aspose.Cells HTML export | remove hidden rows Aspose | HtmlSaveOptions HiddenRowDisplayType | C# Excel to HTML conversion | skip hidden rows in HTML output | .NET spreadsheet to web page
// Common Searches: Aspose.Cells hide row in HTML export | How to exclude hidden rows when saving Excel as HTML | HtmlSaveOptions Remove hidden rows example | C# export Excel to HTML without hidden rows
// Developer Intent: Generate an HTML file from an Excel workbook while automatically omitting any rows that are hidden in the source worksheet.
// Use Cases: Produce clean web reports that show only visible data rows. | Convert Excel dashboards to HTML for publishing, preserving user‑defined visibility. | Automate documentation pipelines where hidden rows should not appear in the final HTML.
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet to HTML and exclude hidden rows. | Explain how HtmlHiddenRowDisplayType.Remove affects the HTML output and when to use it. | Show how to configure HtmlSaveOptions to omit hidden rows, columns, or both during HTML conversion.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, hides a specific row, sets HtmlSaveOptions.HiddenRowDisplayType to Remove, and saves the file as HTML so the hidden row is omitted from the markup.
    public class ExportExcelToHtmlOmitHiddenRows
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data
            worksheet.Cells["A1"].PutValue("Visible Row 1");
            worksheet.Cells["A2"].PutValue("Hidden Row");
            worksheet.Cells["A3"].PutValue("Visible Row 2");

            // Hide the second row (zero‑based index 1)
            worksheet.Cells.HideRow(1);

            // Set HTML save options to remove hidden rows from the output
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove;

            // Save the workbook as HTML; hidden rows will be omitted
            workbook.Save("output.html", htmlOptions);
        }
    }
}
