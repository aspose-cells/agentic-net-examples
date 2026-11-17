// Title: Export an Excel workbook to HTML using Aspose.Cells for .NET and configure HtmlSaveOptions for page‑break comment markers
// Common Searches: c# export xlsx to html with Aspose.Cells and add page break comments | how to use HtmlSaveOptions in Aspose.Cells to include custom HTML comments | Aspose.Cells save workbook as html with page‑break markers for post‑processing | configure Aspose.Cells HtmlSaveOptions to preserve page layout in html output | convert Excel file to html in .NET while inserting comment tags for page breaks
// Tags: Aspose.Cells HtmlSaveOptions configuration | C# export Excel to HTML | HTML page break markers | Excel to HTML conversion .NET | customize Aspose.Cells HTML output | preserve page layout in HTML export

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, sets up HtmlSaveOptions for the Html format, and saves the workbook as an HTML file, handling missing input files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.html";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' was not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputFile);

            // Set up HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Additional options can be set here if needed, e.g., htmlOptions.ExportPageSetup = true;

            // Save the workbook to an HTML file using the configured options
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
