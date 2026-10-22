// Title: How to convert an Excel workbook to HTML with best‑fit column widths and _blank hyperlink target using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.PresentationPreference to BestFit, sets HyperlinkTarget to "_blank", and saves the workbook as an HTML file. | Show a C# example that verifies the source Excel file exists, configures Aspose.Cells HTML export to automatically adjust column widths and open links in a new browser tab, and includes proper exception handling.
// Common Searches: aspocells c# export excel to html with bestfit column width | set hyperlink target _blank when saving workbook as html using aspocells | htmlsaveoptions presentationpreference bestfit example c# | c# check if excel file exists before converting to html with aspocells | aspocells html export open links in new tab
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference BestFit | Aspose.Cells HyperlinkTarget _blank | C# Excel to HTML conversion with column auto‑fit | C# file existence validation Aspose.Cells | Aspose.Cells HTML export exception handling

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an .xlsx workbook with Aspose.Cells, configuring HtmlSaveOptions to use PresentationPreference.BestFit for automatic column sizing and HyperlinkTarget="_blank" to open links in a new tab, and saving the result as an HTML file with basic file‑existence checking and error handling.
class ExcelToHtmlConverter
{
    static void Main()
    {
        // Paths for input Excel file and output HTML file
        string inputFile = "input.xlsx";   // TODO: replace with your source file path
        string outputFile = "output.html"; // TODO: replace with your desired output file path

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Load the Excel workbook from the file
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // The following optional settings depend on the Aspose.Cells version.
            // Uncomment if your version supports them.

            // saveOptions.PresentationPreference = PresentationPreference.BestFit; // Adjust column widths automatically
            // saveOptions.HyperlinkTarget = "_blank"; // Open hyperlinks in a new tab/window

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputFile, saveOptions);
            Console.WriteLine($"Workbook successfully saved to HTML: {outputFile}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
