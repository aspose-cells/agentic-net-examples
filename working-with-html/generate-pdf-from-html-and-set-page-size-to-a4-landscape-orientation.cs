// Title: Convert an HTML file to an A4 landscape PDF using Aspose.Cells for .NET (C#)
// Common Searches: c# convert html to pdf a4 landscape using aspose.cells workbook | how to set page size to A4 and orientation to landscape before saving pdf with aspose.cells | load html file into aspose.cells workbook and export as pdf with custom page setup | asp.net generate pdf from html with aspose.cells and ensure output folder exists | asp.net core html to pdf conversion using aspose.cells pageorientation landscape
// Tags: Aspose.Cells HTML workbook loading | A4 landscape page setup in Aspose.Cells | Export workbook to PDF with Aspose.Cells | Ensure output directory exists C# | HtmlLoadOptions usage Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the HTML input file, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, configures the first worksheet's PageSetup to A4 size and landscape orientation, creates the target directory if needed, and saves the workbook as a PDF, handling any errors that may occur.
class Program
{
    static void Main()
    {
        const string inputPath = "input.html";
        const string outputPath = "output.pdf";

        // Verify that the input HTML file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the HTML file into a workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Configure page setup: A4 size and landscape orientation
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.PaperSize = PaperSizeType.PaperA4;
            pageSetup.Orientation = PageOrientationType.Landscape;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
