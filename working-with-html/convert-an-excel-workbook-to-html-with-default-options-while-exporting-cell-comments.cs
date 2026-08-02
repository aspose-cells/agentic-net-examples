// Title: Convert Excel to HTML with Cell Comments using Aspose.Cells for .NET (C#)
// Description: Shows how to load an .xlsx file, enable comment export with HtmlSaveOptions.IsExportComments, and save the workbook as an HTML page using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Excel to HTML | HTML export | cell comments | HtmlSaveOptions | IsExportComments | Workbook.Save | convert workbook to HTML
// Common Searches: Aspose.Cells export Excel comments to HTML C# | HtmlSaveOptions default settings example | How to save workbook as HTML with comments .NET | C# convert .xlsx to HTML using Aspose.Cells | Enable comment export in Aspose.Cells HTML output
// Developer Intent: Generate an HTML representation of an Excel workbook that includes any cell comments.
// Use Cases: Create web‑ready spreadsheet previews that retain author notes for documentation. | Produce HTML reports from Excel templates where annotations must be visible to stakeholders. | Automate server‑side conversion of uploaded Excel files to HTML for in‑browser preview, preserving comments.
// AI Prompts: Write C# code that loads an Excel file and saves it as HTML with comments using Aspose.Cells default HtmlSaveOptions. | Explain the purpose of HtmlSaveOptions.IsExportComments and its impact on the generated HTML. | Provide a step‑by‑step tutorial for converting an .xlsx workbook to HTML with comment support, including required NuGet packages.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to load an .xlsx file, enable comment export with HtmlSaveOptions.IsExportComments, and save the workbook as an HTML page using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // OPTIONAL: Add a comment to demonstrate comment export
            // This step can be omitted if the source workbook already contains comments
            Worksheet sheet = workbook.Worksheets[0];
            int commentIndex = sheet.Comments.Add("B2");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Sample comment that will be exported to HTML";

            // Create HTML save options with default settings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable exporting of cell comments (default is false)
            htmlOptions.IsExportComments = true;

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", htmlOptions);
        }
    }
}
