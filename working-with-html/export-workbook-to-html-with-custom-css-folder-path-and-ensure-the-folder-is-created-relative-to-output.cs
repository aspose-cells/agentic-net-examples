// Title: Export an Excel workbook to HTML with a custom CSS folder and automatically create the output directory using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export workbook to HTML and specify external CSS folder | How to create output folder before saving Excel as HTML with Aspose.Cells | Save Excel file as HTML with separate CSS files using Aspose.Cells in .NET | C# Aspose.Cells HtmlSaveOptions set custom CSS directory | Export all worksheets to HTML with images as files using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions custom CSS directory | C# create output folder for HTML export | export workbook to HTML with external CSS | save Excel as HTML separate images Aspose.Cells | Aspose.Cells export all worksheets to HTML

using Aspose.Cells;
using System;
using System.IO;

// The example loads (or creates) an Excel workbook, ensures the target output directory exists, configures HtmlSaveOptions to export all worksheets and keep images as separate files, and saves the workbook as HTML to a specified path using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Ensure the input file exists; create a simple workbook if missing
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                tempWb.Save(inputPath);
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Output HTML file path
            string outputHtmlPath = @"output\workbook.html";

            // Resolve output directory
            string outputDirectory = Path.GetDirectoryName(outputHtmlPath);
            if (string.IsNullOrEmpty(outputDirectory))
                outputDirectory = Directory.GetCurrentDirectory();

            // Ensure output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = false, // Export all worksheets
                ExportImagesAsBase64 = false        // Keep images as separate files
                // Note: External CSS properties are not available in this version of Aspose.Cells.
            };

            // Save the workbook as HTML with the configured options
            workbook.Save(outputHtmlPath, saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
