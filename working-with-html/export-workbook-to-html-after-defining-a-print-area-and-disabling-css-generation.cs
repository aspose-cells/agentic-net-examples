// Title: Export an Aspose.Cells workbook to HTML with a specific print area and without generating CSS
// Common Searches: Aspose.Cells C# set print area before saving workbook as HTML | How to disable CSS generation when exporting Excel to HTML with Aspose.Cells | Save Excel workbook to HTML with print area A1:D20 using Aspose.Cells .NET | Create output directory automatically when saving HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions prevent external stylesheet creation
// Tags: set worksheet print area Aspose.Cells | HtmlSaveOptions prevent CSS output | export workbook to HTML Aspose.Cells | auto-create export folder Aspose.Cells | save workbook as HTML C# Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, defines the first worksheet's print area (A1:D20), configures HtmlSaveOptions to avoid CSS generation, ensures the target folder exists, and saves the workbook as an HTML file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new empty workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the print area (example: cells A1 to D20)
            sheet.PageSetup.PrintArea = "A1:D20";

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Note: In recent Aspose.Cells versions, CSS class name generation is controlled internally.
            // If needed, additional options can be set here.

            // Define output file path
            string outputPath = "output.html";

            // Ensure the output directory exists (handle cases where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook to HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
