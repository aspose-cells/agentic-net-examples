// Title: Aspose.Cells for .NET – Export Excel to HTML without downlevel‑revealed comments
// Description: Demonstrates how to use Aspose.Cells HtmlSaveOptions.DisableDownlevelRevealedComments in C# to generate clean HTML from a workbook, removing legacy conditional comment syntax.
// Keywords: Aspose.Cells | HtmlSaveOptions | DisableDownlevelRevealedComments | C# HTML export | Excel to HTML | .NET workbook conversion | remove conditional comments
// Common Searches: disable downlevel revealed comments Aspose.Cells | Aspose.Cells export Excel to HTML without IE comments | HtmlSaveOptions.DisableDownlevelRevealedComments example | clean HTML output from Aspose.Cells .NET | remove conditional comments when saving workbook as HTML
// Developer Intent: Export an Excel workbook to HTML while suppressing downlevel‑revealed conditional comment markup.
// Use Cases: Create web‑ready HTML reports from Excel without legacy IE comment syntax. | Generate markup that passes HTML validators for modern browsers. | Produce lightweight HTML files for embedding in web applications.
// AI Prompts: How do I set HtmlSaveOptions.DisableDownlevelRevealedComments in Aspose.Cells C#? | Show a C# code snippet that saves a workbook to HTML without conditional comments. | Explain the impact of disabling downlevel‑revealed comments on the generated HTML.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to use Aspose.Cells HtmlSaveOptions.DisableDownlevelRevealedComments in C# to generate clean HTML from a workbook, removing legacy conditional comment syntax.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty by default)
            Workbook workbook = new Workbook();

            // Add some sample data to demonstrate the export
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            sheet.Cells["B2"].PutValue(12345);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Disable downlevel-revealed conditional comments in the generated HTML
                DisableDownlevelRevealedComments = true
            };

            // Save the workbook as HTML using the configured options
            string outputPath = "ExportedWorkbook.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved to '{outputPath}' with DisableDownlevelRevealedComments = true");
        }
    }
}
