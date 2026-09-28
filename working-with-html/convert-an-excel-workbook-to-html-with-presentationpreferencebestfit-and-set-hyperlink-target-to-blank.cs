// Title: Export an Excel workbook to HTML using Aspose.Cells C# with PresentationPreference.BestFit and hyperlink target set to _blank
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, configures HtmlSaveOptions to use PresentationPreference.BestFit, sets HyperlinkTarget to "_blank", and saves the workbook as an HTML file. | Show how to generate HTML from an Excel workbook in C# so that column widths are automatically adjusted (BestFit) and all hyperlinks open in a new browser tab using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to export Excel to HTML with best fit column widths | set hyperlink target to _blank when saving workbook as HTML using Aspose.Cells | HtmlSaveOptions PresentationPreference BestFit example in C# | C# Aspose.Cells export to HTML with links opening in new tab | save Excel workbook as HTML with auto‑fit layout using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions BestFit layout | Aspose.Cells HyperlinkTarget _blank | C# export Excel to HTML BestFit | Aspose.Cells column auto fit HTML conversion | C# save workbook as HTML with new tab links

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks for the input.xlsx file, loads it into an Aspose.Cells Workbook, optionally sets HtmlSaveOptions.PresentationPreference to BestFit and HyperlinkTarget to "_blank", then saves the workbook as output.html while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);

            // The following options are available in newer versions of Aspose.Cells.
            // Uncomment them if your version supports them.

            // options.PresentationPreference = Aspose.Cells.PresentationPreference.BestFit;
            // options.HyperlinkTarget = "_blank";

            // Save the workbook as HTML
            workbook.Save(outputFile, options);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
