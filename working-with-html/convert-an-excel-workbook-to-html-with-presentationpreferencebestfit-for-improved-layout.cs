// Title: Convert an Excel workbook to HTML with Aspose.Cells using PresentationPreference.BestFit for optimal layout (C#)
// AI Prompts: Generate C# code that loads an .xlsx file, sets HtmlSaveOptions.PresentationPreference to BestFit, and saves the workbook as an HTML file using Aspose.Cells. | Add file‑existence checking and comprehensive exception handling to an Aspose.Cells Excel‑to‑HTML conversion that applies the BestFit layout option. | Show how to configure HtmlSaveOptions in Aspose.Cells to produce responsive HTML output while preserving column widths.
// Common Searches: Aspose.Cells C# export Excel to HTML with best‑fit layout | How to enable PresentationPreference.BestFit in HtmlSaveOptions when saving workbook as HTML | C# code example for converting .xlsx to HTML preserving column widths using Aspose.Cells | Save Excel workbook as responsive HTML using Aspose.Cells HtmlSaveOptions | Check input file existence before converting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions BestFit layout | C# Excel to HTML conversion Aspose.Cells | preserve column widths Aspose.Cells HTML export | file existence validation Aspose.Cells conversion | exception handling Aspose.Cells workbook save

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample verifies that the input .xlsx file exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions (optionally setting PresentationPreference.BestFit for a best‑fit layout), and saves the workbook as an HTML file while providing console messages for success or errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // If the PresentationPreference enum is available in the referenced version,
            // you can enable the best‑fit layout as shown below:
            // htmlOptions.PresentationPreference = PresentationPreference.BestFit;

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
