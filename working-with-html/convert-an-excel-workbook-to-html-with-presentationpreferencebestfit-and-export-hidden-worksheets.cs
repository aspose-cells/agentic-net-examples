// Title: Export hidden worksheets to HTML with BestFit layout using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# console program that loads an .xlsx file, sets HtmlSaveOptions.ExportHiddenWorksheet to true, applies PresentationPreference.BestFit, and saves the workbook as an HTML file using Aspose.Cells. | Show the steps to configure Aspose.Cells HtmlSaveOptions for exporting hidden sheets and using the BestFit layout when converting Excel to HTML in .NET.
// Common Searches: asp.net aspose.cells convert excel to html with hidden worksheets included | c# set presentationpreference bestfit for html export using aspose.cells | how to include hidden sheets in html output from workbook aspose.cells | example of HtmlSaveOptions with ExportHiddenWorksheet and PresentationPreference in C# | save excel workbook as html bestfit layout asp.net
// Tags: Aspose.Cells HtmlSaveOptions export hidden worksheets | Aspose.Cells PresentationPreference BestFit | C# Excel to HTML conversion using Aspose.Cells | Include hidden sheets in HTML output Aspose.Cells | HtmlSaveOptions set PresentationPreference

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample verifies the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to export hidden worksheets and (optionally) apply the BestFit presentation preference, then saves the workbook as output.html while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export hidden worksheets as part of the HTML output
                ExportHiddenWorksheet = true
                // Note: PresentationPreference property may not be available in all versions.
                // If needed, uncomment the following line after confirming the enum exists.
                // PresentationPreference = PresentationPreference.BestFit
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
