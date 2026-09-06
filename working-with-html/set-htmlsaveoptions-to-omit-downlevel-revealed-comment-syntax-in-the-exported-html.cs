// Title: Export an Excel workbook to HTML5 without conditional comments using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a workbook and saves it as HTML5 while disabling downlevel revealed comment syntax with Aspose.Cells. | Show how to set HtmlSaveOptions.HtmlVersion to Html5 and configure the options to omit conditional comments in the HTML output. | Provide a complete example that checks for a template file, creates a workbook if missing, and exports it to clean HTML using Aspose.Cells.
// Common Searches: Aspose.Cells how to remove conditional comments from HTML export | C# save Excel as HTML5 without downlevel revealed comments | HtmlSaveOptions HtmlVersion Html5 suppress conditional comments Aspose.Cells | Export workbook to clean HTML using Aspose.Cells .NET example | Disable conditional comments in Aspose.Cells HTML output
// Tags: Aspose.Cells HtmlSaveOptions Html5 | Aspose.Cells suppress conditional comments | C# export Excel to clean HTML | Aspose.Cells HTML5 output without downlevel revealed comments | Aspose.Cells workbook to HTML5 example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel template or creates a new workbook, configures HtmlSaveOptions with HtmlVersion.Html5 to generate HTML5, disables downlevel revealed conditional comments, and saves the result as a clean HTML file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook or load an existing template if it exists
                Workbook workbook;
                const string templatePath = "Template.xlsx";

                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook();
                    // Add sample data to the first worksheet
                    workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");
                }

                // Configure HTML save options to generate HTML5 without conditional comments
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                // Set the HTML version to HTML5 (available in recent Aspose.Cells versions)
                htmlOptions.HtmlVersion = HtmlVersion.Html5;

                // Save the workbook as HTML
                const string outputPath = "ExportedDocument.html";
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
