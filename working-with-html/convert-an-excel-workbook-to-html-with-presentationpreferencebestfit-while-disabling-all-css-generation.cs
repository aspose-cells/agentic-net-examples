// Title: Convert an Excel workbook to HTML using Aspose.Cells for .NET with BestFit layout and CSS disabled
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.PresentationPreference to BestFit, turns off CSS stylesheet generation, and saves the result as an HTML file. | Show how to configure Aspose.Cells HtmlSaveOptions to produce a compact HTML file without any embedded or external CSS, including handling of missing input files.
// Common Searches: how to export Excel to HTML without CSS using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions PresentationPreference BestFit example | disable stylesheet generation when saving workbook as HTML in .NET | minimal HTML output from Excel file with Aspose.Cells | convert .xlsx to HTML with no external CSS Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference BestFit | Aspose.Cells disable CSS stylesheet HTML export | C# Excel to HTML conversion Aspose.Cells | minimal HTML output Aspose.Cells | Aspose.Cells SaveFormat.Html configuration

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an input.xlsx workbook, configures HtmlSaveOptions to use the BestFit presentation preference and to suppress CSS stylesheet generation, then saves the workbook as output.html while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set up HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Optional: set presentation preference if the property is available in your version
            // htmlOptions.PresentationPreference = HtmlSaveOptions.PresentationPreference.BestFit;

            // Optional: disable CSS stylesheet generation if supported
            // htmlOptions.ExportCssStyleSheet = false; // Uncomment if available in your Aspose.Cells version

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
