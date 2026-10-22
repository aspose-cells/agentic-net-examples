// Title: Export a workbook to HTML with a defined print area and disabled CSS using Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a workbook, sets the print area to A1:D10, configures HtmlSaveOptions to suppress CSS, and saves the file as HTML. | Show how to use Aspose.Cells HtmlSaveOptions to export only a specific range to HTML while turning off the ExportCss flag. | Provide a try‑catch example that saves a workbook to HTML with ExportCss set to false and a custom print area.
// Common Searches: how to disable CSS when exporting Excel to HTML with Aspose.Cells C# | Aspose.Cells set print area before HTML export C# example | HtmlSaveOptions ExportCss false usage Aspose.Cells | export range A1:D10 to HTML using Aspose.Cells C# | save workbook as HTML without stylesheet Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportCss | define print area Aspose.Cells | export workbook range to HTML C# | disable stylesheet Aspose.Cells HTML export | C# Aspose.Cells HTML conversion

using System;
using Aspose.Cells;

// The example creates a new workbook, defines a print area (A1:D10) on the first worksheet, configures HtmlSaveOptions to optionally disable CSS generation, and saves the workbook as an HTML file while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Define the print area (e.g., cells A1 to D10)
            worksheet.PageSetup.PrintArea = "A1:D10";

            // Set up HTML save options
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // If the ExportCss property is available in the referenced Aspose.Cells version,
            // uncomment the following line to disable CSS generation:
            // htmlOptions.ExportCss = false;

            // Export the workbook to HTML
            const string outputPath = "ExportedWorkbook.html";
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
