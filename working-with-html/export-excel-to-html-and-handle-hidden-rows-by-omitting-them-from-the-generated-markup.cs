// Title: Export Excel to HTML without hidden rows – Aspose.Cells for .NET
// Description: Shows how to create a workbook, hide selected rows, configure HtmlSaveOptions.HiddenRowDisplayType = Remove, and save the file as HTML so hidden rows are omitted from the output markup.
// Keywords: Aspose.Cells | HTML export | HiddenRowDisplayType | Remove hidden rows | C# Excel to HTML | Hide rows Aspose | HtmlSaveOptions | Export without hidden rows | Aspose.Cells .NET | Excel to web
// Common Searches: Aspose.Cells export Excel to HTML without hidden rows | HtmlSaveOptions HiddenRowDisplayType.Remove example C# | How to hide rows and exclude them from HTML export Aspose.Cells | C# generate HTML from workbook omitting hidden rows | Aspose.Cells HTML export hide row settings
// Developer Intent: Create an HTML representation of an Excel workbook that automatically excludes any rows marked as hidden.
// Use Cases: Web‑based reporting where calculation rows are hidden and must not appear in the HTML view. | Generating printable HTML versions of spreadsheets while keeping internal‑only rows invisible to end users. | Providing a clean data feed for web applications, stripping out rows that are hidden in the source workbook.
// AI Prompts: Provide C# code that uses Aspose.Cells to export a workbook to HTML and removes hidden rows via HtmlSaveOptions. | Explain the effect of HtmlHiddenRowDisplayType.Remove and how to apply it when saving Excel as HTML. | Show how to hide specific rows in a worksheet and ensure they are omitted from the generated HTML output.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to create a workbook, hide selected rows, configure HtmlSaveOptions.HiddenRowDisplayType = Remove, and save the file as HTML so hidden rows are omitted from the output markup.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Visible Row 1");
            sheet.Cells["A3"].PutValue("Hidden Row");
            sheet.Cells["A4"].PutValue("Visible Row 2");

            // Hide the third row (index 2)
            sheet.Cells.HideRow(2);

            // Configure HTML save options to remove hidden rows from the output
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.HiddenRowDisplayType = HtmlHiddenRowDisplayType.Remove;

            // Save the workbook as HTML; hidden rows will be omitted
            workbook.Save("ExportedWithoutHiddenRows.html", htmlOptions);

            Console.WriteLine("HTML export completed. Hidden rows have been removed from the markup.");
        }
    }
}
