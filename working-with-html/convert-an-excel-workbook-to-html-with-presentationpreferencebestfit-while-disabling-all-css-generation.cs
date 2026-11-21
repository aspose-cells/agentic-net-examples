// Title: Convert an Excel workbook to HTML with PresentationPreference.BestFit and suppress all CSS using Aspose.Cells for C#
// Common Searches: Aspose.Cells C# export Excel to HTML with PresentationPreference BestFit and no CSS | How to disable stylesheet generation when saving a workbook as HTML using Aspose.Cells | C# HtmlSaveOptions PresentationPreference BestFit without generating CSS files | Save .xlsx as .html without embedded CSS using Aspose.Cells library | Best‑fit layout HTML export from Excel in C# with CSS disabled
// Tags: Aspose.Cells HtmlSaveOptions PresentationPreference BestFit | disable CSS generation Aspose.Cells HTML export | C# convert Excel to HTML without stylesheet | save workbook as HTML without embedded CSS | Excel to HTML best‑fit layout Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, creates HtmlSaveOptions, and saves the file as HTML. By configuring HtmlSaveOptions.PresentationPreference = PresentationPreference.BestFit and setting HtmlSaveOptions.DisableCss = true, you can produce a best‑fit HTML layout while preventing any CSS files from being generated.
class Program
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (using defaults)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Convert and save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully converted to HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
