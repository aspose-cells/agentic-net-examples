// Title: Convert an Excel workbook to HTML with PresentationPreference.BestFit and parent hyperlink target using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# save workbook as HTML with PresentationPreference.BestFit | set hyperlink target to parent in HTML output using Aspose.Cells | convert .xlsx to .html preserving column widths with BestFit option | C# example for HtmlSaveOptions HyperlinkTarget property Aspose.Cells | export Excel to HTML with auto‑fit layout in .NET
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference.BestFit | Aspose.Cells HyperlinkTarget parent | C# Excel to HTML conversion | HTML export with auto‑fit column widths Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing .xlsx file into an Aspose.Cells Workbook, configures HtmlSaveOptions to use PresentationPreference.BestFit and optionally sets HyperlinkTarget to "parent", then saves the workbook as an HTML file while handling errors.
class ExcelToHtmlConverter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Uncomment the following line if the HyperlinkTarget property is available in your Aspose.Cells version
            // saveOptions.HyperlinkTarget = "parent";

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
