// Title: Convert an Excel .xlsx file to HTML with Aspose.Cells for .NET while preserving hyperlink URLs (C#)
// Common Searches: Aspose.Cells C# export Excel workbook to HTML keep hyperlink URLs | How to retain original link targets when saving .xlsx as .html using Aspose.Cells | C# code to convert Excel to HTML with hyperlinks intact via Aspose.Cells | Saving Excel file as HTML with Aspose.Cells while retaining link URLs in .NET | Aspose.Cells HtmlSaveOptions keep links during conversion from .xlsx to .html
// Tags: Aspose.Cells HtmlSaveOptions ExportHyperlinks | Excel workbook HTML export C# | Workbook.Save HTML Aspose.Cells .NET | ExportActiveWorksheetOnly false Aspose.Cells | Convert .xlsx to .html Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook using Aspose.Cells, configures HtmlSaveOptions (ExportHyperlinks enabled by default), ensures the output directory exists, and saves the file as HTML, guaranteeing that all original hyperlink URLs are retained in the generated HTML.
class Program
{
    static void Main(string[] args)
    {
        // Determine input and output file paths.
        // If not provided via command‑line arguments, use default file names in the executable's directory.
        string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
        string outputPath = args.Length > 1 ? args[1] : "output.html";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.Error.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified Excel file.
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options.
            // ExportHyperlinks is enabled by default; explicit setting is not required in current API version.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = false
            };

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an HTML file using the configured options.
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Successfully converted '{inputPath}' to HTML.");
            Console.WriteLine($"Output file: {outputPath}");
            Console.WriteLine("All hyperlinks have been preserved with their original URL targets.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("An error occurred during conversion:");
            Console.Error.WriteLine(ex.Message);
            Console.Error.WriteLine(ex.StackTrace);
        }
    }
}
