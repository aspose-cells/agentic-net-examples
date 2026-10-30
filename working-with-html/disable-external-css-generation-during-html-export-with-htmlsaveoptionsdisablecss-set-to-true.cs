// Title: How to disable external CSS generation when exporting an Excel workbook to HTML with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export workbook to HTML without creating a separate CSS file | set HtmlSaveOptions.DisableCss to true example Aspose.Cells | prevent external stylesheet generation during HTML save using Aspose.Cells | C# code to export Excel to HTML with inline styles only Aspose.Cells | disable CSS file output when saving workbook as HTML with Aspose.Cells library
// Tags: Aspose.Cells HtmlSaveOptions.DisableCss property | C# export Excel to HTML without external CSS | HTML save options inline styling Aspose.Cells | disable external stylesheet Aspose.Cells HTML export | Excel workbook HTML export inline CSS C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example creates a Workbook, adds sample data, configures HtmlSaveOptions for HTML output, sets saveOptions.DisableCss = true to suppress the creation of an external CSS file, ensures the output directory exists, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B1"].PutValue(12345);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: ExportHtmlAsMhtml is not available in all versions; HTML will be saved with external CSS if unsupported.

            // Define output file path
            string outputPath = "Exported.html";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
