// Title: How to export an Excel worksheet to HTML with cell text wrapping preserved using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, writes a long string to a cell, enables text wrapping via the Style.IsTextWrapped property, adjusts column width and row height, and saves the worksheet as HTML with the wrap retained. | Demonstrate how to configure HtmlSaveOptions to export only the active worksheet while keeping the cell's wrap settings intact in the generated HTML file.
// Common Searches: Aspose.Cells C# preserve cell wrap when saving as HTML | export Excel to HTML with text wrap enabled using Aspose.Cells | how to keep text wrapping in HTML output from Aspose.Cells workbook | set column width and row height for wrapped text in Aspose.Cells HTML export
// Tags: export worksheet to HTML with text wrap Aspose.Cells | Style.IsTextWrapped property C# | HtmlSaveOptions ExportActiveWorksheetOnly | adjust column width for wrapped text Aspose.Cells | auto fit row height HTML export Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, inserts a long string into cell A1, enables text wrapping via the Style.IsTextWrapped flag, sets column width and row height to showcase wrapping, configures HtmlSaveOptions to export only the active sheet, and saves the result as an HTML file where the cell's wrap setting is retained.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a long text into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("This is a very long piece of text that should wrap inside the cell when the workbook is exported to HTML.");

        // Enable text wrapping for the cell
        Style style = cell.GetStyle();
        style.IsTextWrapped = true;
        cell.SetStyle(style);

        // Adjust column width to make wrapping visible
        sheet.Cells.SetColumnWidth(0, 20); // Column A width

        // Optionally set row height (auto‑fit can also be used)
        sheet.Cells.SetRowHeight(0, 40);

        // Prepare HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Export only the active worksheet (optional)
            ExportActiveWorksheetOnly = true
        };

        // Save the workbook as HTML; wrap settings are preserved in the generated HTML
        workbook.Save("output.html", htmlOptions);
    }
}
