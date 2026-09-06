// Title: Convert an Excel workbook to HTML with column auto‑fit (PresentationPreference.BestFit) and preserve cell comments using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook, sets HtmlSaveOptions.PresentationPreference to BestFit, enables comment export, and saves the result as an HTML file. | Show a .NET console example that validates the source file, configures column auto‑fit and comment preservation, then writes the HTML output with Aspose.Cells. | Explain the steps to configure Aspose.Cells HtmlSaveOptions for preserving Excel comments and applying best‑fit column sizing during HTML conversion.
// Common Searches: Aspose.Cells C# best fit column width when saving workbook as HTML | How to export Excel cell comments to HTML using Aspose.Cells | HtmlSaveOptions PresentationPreference BestFit usage example | Convert .xlsx to .html with comments retained Aspose.Cells | Auto‑fit columns in HTML output from Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference | HTML export preserve cell comments Aspose.Cells | BestFit column auto‑fit Aspose.Cells HTML | C# Excel to HTML conversion using Aspose.Cells | Aspose.Cells auto‑fit columns during HTML conversion

using System;
using System.IO;
using Aspose.Cells;

// The sample program verifies that the input.xlsx file exists, loads it into an Aspose.Cells Workbook, creates HtmlSaveOptions with PresentationPreference set to BestFit and ExportCellComments enabled, then saves the workbook as output.html. It includes exception handling to report any errors during the conversion.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Set up HTML save options with default settings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as an HTML file using the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
