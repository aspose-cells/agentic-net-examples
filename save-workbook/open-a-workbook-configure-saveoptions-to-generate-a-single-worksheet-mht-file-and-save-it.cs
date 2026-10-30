// Title: Generate a single‑worksheet HTML file from an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export only the active sheet to HTML | How to save a specific worksheet as HTML with Aspose.Cells .NET | HtmlSaveOptions ExportActiveWorksheetOnly example in C# | Create single‑sheet HTML output from Excel using Aspose.Cells | Save active worksheet as .html using Aspose.Cells SaveFormat.Html
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly | C# save active worksheet as HTML using Aspose.Cells | Aspose.Cells SaveFormat.Html single worksheet export | export active sheet only Aspose.Cells .NET | generate single‑sheet HTML from Excel workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The code loads an existing Excel file, activates the first worksheet, configures HtmlSaveOptions with ExportActiveWorksheetOnly set to true, and saves that worksheet as a single‑worksheet HTML file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the first worksheet is the active one
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Configure HTML save options to export only the active worksheet
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true
            };

            // Save the workbook as a single‑worksheet HTML file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
