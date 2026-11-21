// Title: Export the active worksheet to HTML with a custom TableCssId using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export only the active sheet to HTML | Assign a TableCssId when saving Excel as HTML with Aspose.Cells | Example of HtmlSaveOptions ExportActiveWorksheetOnly in .NET | Generate HTML from a specific worksheet and set custom CSS ID using Aspose.Cells | How to set custom table identifier in HTML output from Excel with Aspose.Cells
// Tags: HtmlSaveOptions active worksheet export | HTML table CSS identifier Aspose.Cells | active worksheet to HTML C# | custom table CSS ID Aspose.Cells | Aspose.Cells HTML save options usage

using System;
using Aspose.Cells;

// Loads an Excel workbook, optionally activates a worksheet, configures HtmlSaveOptions to export only the active worksheet and assigns a custom TableCssId for the generated HTML table, then saves the result as an HTML file.
class Program
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure the worksheet you want to export is the active one.
        // If you need to set a specific worksheet as active, uncomment the line below:
        // workbook.Worksheets[0].Activate(); // example: activate the first worksheet

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Export only the active worksheet
            ExportActiveWorksheetOnly = true,

            // Assign a CSS ID to the generated HTML table for custom styling
            TableCssId = "myCustomTableId"
        };

        // Save the active worksheet as an HTML file
        workbook.Save("output.html", htmlOptions);
    }
}
