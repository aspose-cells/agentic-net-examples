// Title: Export Excel cell comments to HTML with Aspose.Cells in C# using HtmlSaveOptions
// Common Searches: Aspose.Cells C# save workbook as HTML with cell comments visible | How to include worksheet comments when converting Excel to HTML using Aspose.Cells | Enable ExportComments in HtmlSaveOptions for Aspose.Cells .NET | C# generate HTML from Excel that retains cell notes with Aspose.Cells | Aspose.Cells HTML export comments not showing – troubleshooting
// Tags: Aspose.Cells HtmlSaveOptions ExportComments | C# export cell notes to HTML | Excel comment rendering in HTML using Aspose | HTML conversion with worksheet comments Aspose.Cells | save workbook as HTML with comment support

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, adds a comment to cell A1, configures HtmlSaveOptions with ExportComments enabled, ensures the output directory exists, and saves the workbook as an HTML file so the comment appears in the generated HTML.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a sample comment";

            // Configure HTML save options (default options export comments)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Define output file path
            string outputPath = "output.html";

            // Ensure the directory for the output file exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
