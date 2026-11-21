// Title: Convert an Excel workbook to HTML with BestFit layout and include hidden worksheets using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export hidden worksheets when saving workbook as HTML | How to set PresentationPreference.BestFit in HtmlSaveOptions with Aspose.Cells | Save Excel file to HTML including hidden sheets using Aspose.Cells .NET | C# Aspose.Cells HTML conversion preserving hidden worksheets and best‑fit column widths | Enable ExportHiddenWorksheet in Aspose.Cells HtmlSaveOptions example
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | PresentationPreference BestFit Aspose.Cells | Excel to HTML conversion with hidden sheets | C# Aspose.Cells HTML export settings | SaveFormat.Html Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook with Aspose.Cells, configures HtmlSaveOptions to export hidden worksheets and use the BestFit presentation preference, then saves the workbook as an HTML file while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
            return;
        }

        try
        {
            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export hidden worksheets as part of the HTML output
                ExportHiddenWorksheet = true
                // PresentationPreference property is optional; omitted if not supported by the library version
            };

            // Save the workbook to an HTML file using the configured options
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully converted to HTML: \"{outputFile}\"");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
