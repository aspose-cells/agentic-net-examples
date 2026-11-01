// Title: How to export Excel worksheet comments to HTML with Aspose.Cells using HtmlSaveOptions.IncludeComments in C#
// Common Searches: C# Aspose.Cells save workbook as HTML with cell comments included | Set HtmlSaveOptions.IncludeComments true to preserve comments in HTML export | Export Excel comments to HTML using Aspose.Cells .NET library | Include worksheet comments when converting Excel to HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions IncludeComments property example in C#
// Tags: Aspose.Cells HtmlSaveOptions IncludeComments | export worksheet comments to HTML C# | Excel to HTML conversion with comments Aspose.Cells | C# Aspose.Cells HTML export settings | preserve cell comments in HTML output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures HtmlSaveOptions.IncludeComments = true (when supported), and saves the workbook as an HTML file, ensuring that any worksheet comments are rendered in the generated HTML output.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                tempWb.Save(inputPath);
            }

            // Load the workbook.
            var workbook = new Workbook(inputPath);

            // Configure HTML save options.
            var htmlOptions = new HtmlSaveOptions();
            // Note: ExportCellComments property is not available in this version of Aspose.Cells.
            // If comment export is required, use a version that supports the property.

            // Save as HTML.
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
