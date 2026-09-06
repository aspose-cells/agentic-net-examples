// Title: Export an Excel workbook to HTML while preserving multi‑row merged cells using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets HtmlSaveOptions to keep merged cells that span several rows, and saves the result as an HTML file. | Show how to configure Aspose.Cells HtmlSaveOptions so that merged cell ranges are rendered correctly in the HTML output. | Add robust file‑existence checking and exception handling to a C# Excel‑to‑HTML conversion that maintains the original merged‑cell layout.
// Common Searches: Aspose.Cells C# export Excel to HTML with merged cells spanning rows | How to keep multi‑row merged cells when converting .xlsx to HTML using Aspose | HtmlSaveOptions PreserveMergedCells property Aspose.Cells | C# convert workbook to HTML preserving merged cell layout | Export all worksheets to HTML Aspose.Cells merged cells issue
// Tags: Aspose.Cells HtmlSaveOptions merged cells | C# Excel to HTML conversion preserving layout | Export merged cell ranges to HTML Aspose | HtmlSaveOptions ExportActiveWorksheetOnly false | Exception handling file validation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the presence of input.xlsx, loads it with Aspose.Cells Workbook, configures HtmlSaveOptions (ExportActiveWorksheetOnly = false) to retain merged cells that span multiple rows, and saves the workbook as output.html while handling any runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Set HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export all worksheets; set to true to export only the active sheet
                ExportActiveWorksheetOnly = false
            };

            // Export the workbook to HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
