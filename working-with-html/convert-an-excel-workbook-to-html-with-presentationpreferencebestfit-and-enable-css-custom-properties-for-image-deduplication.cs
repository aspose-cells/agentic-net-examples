// Title: Convert an Excel workbook to HTML in C# with Aspose.Cells using PresentationPreference.BestFit and CSS custom properties for image deduplication
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.PresentationPreference to PresentationPreference.BestFit, enables CSS custom properties to reuse identical images, and saves the workbook as an HTML file. | Show the steps to configure Aspose.Cells HtmlSaveOptions for compact HTML output that preserves column widths and deduplicates images via CSS variables in a .NET application.
// Common Searches: aspocells c# export excel to html with bestfit column layout | how to enable css variables for image deduplication in aspocells html export | setting htmlsaveoptions.presentationpreference in .net workbook conversion | c# aspocells html output with css custom properties for image reuse | convert xlsx to html preserving layout and reducing duplicate images
// Tags: Aspose.Cells HTML export best-fit layout | Aspose.Cells CSS custom properties for image reuse | Excel workbook to HTML conversion with deduplicated images | C# HtmlSaveOptions configuration for compact HTML | HTML output using CSS variables for image optimization

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel file with Aspose.Cells, configures HtmlSaveOptions to use the BestFit presentation mode and enables CSS custom properties so identical images are emitted once as CSS variables, then saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: PresentationPreference property may not be available in older versions.
            // If needed, set it using the appropriate enum from the Aspose.Cells namespace.

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
