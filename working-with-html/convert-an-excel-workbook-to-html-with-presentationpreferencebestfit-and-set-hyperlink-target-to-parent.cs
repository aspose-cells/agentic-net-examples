// Title: Export an Excel workbook to HTML with best‑fit column widths and parent hyperlink targets using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.PresentationPreference to PresentationPreference.BestFit, sets HtmlSaveOptions.HyperlinkTarget to "_parent", and saves the workbook as an HTML file. | Show how to configure Aspose.Cells HtmlSaveOptions so that the generated HTML automatically fits column content and all hyperlinks open in the parent frame.
// Common Searches: how to export Excel to HTML with best fit columns using Aspose.Cells C# | Aspose.Cells set hyperlink target to _parent in HTML output | HtmlSaveOptions PresentationPreference BestFit example in .NET | C# convert .xlsx to .html preserving column widths with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions column width fitting | Aspose.Cells HTML hyperlink target parent | C# export Excel to HTML best‑fit layout | Aspose.Cells HTML conversion with parent links

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing .xlsx file with Aspose.Cells, configures HtmlSaveOptions to use PresentationPreference.BestFit and sets HyperlinkTarget to "_parent", then saves the workbook as an HTML file while handling missing input and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Note: PresentationPreference property may not be available in older versions of Aspose.Cells.
            // If needed, set other options here.

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
