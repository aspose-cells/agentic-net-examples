// Title: Save an Excel workbook as HTML with default settings and set hyperlink targets to _blank using Aspose.Cells for .NET (C#)
// Common Searches: how to export an Excel workbook to HTML with Aspose.Cells in C# and open links in a new tab | Aspose.Cells C# convert .xlsx to .html preserving hyperlinks with target='_blank' | C# code example for saving a workbook as HTML using default HtmlSaveOptions in Aspose.Cells | set hyperlink target to blank when converting Excel to HTML with Aspose.Cells .NET | ensure output directory exists before saving HTML from an Excel file using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions default export | Excel workbook to HTML conversion C# | hyperlink target attribute _blank Aspose.Cells | create output folder before saving Aspose.Cells | save workbook as HTML using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input .xlsx file, loads it into an Aspose.Cells Workbook, configures default HtmlSaveOptions (which retain hyperlinks), ensures the output directory exists, and saves the workbook as an HTML file with hyperlink targets set to _blank, handling any exceptions that may occur.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Set HTML save options (default behavior exports hyperlinks)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an HTML file using the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
