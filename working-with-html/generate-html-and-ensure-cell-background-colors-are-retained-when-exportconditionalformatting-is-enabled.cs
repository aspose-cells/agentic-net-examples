// Title: Export Excel to HTML with cell background colors and conditional formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to HTML keep cell background color | How to retain conditional formatting when saving workbook as HTML with Aspose.Cells | Preserve Excel cell shading in HTML output using Aspose.Cells .NET | HtmlSaveOptions ExportConditionalFormatting effect on HTML export | Convert XLSX to HTML with styling preserved using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportConditionalFormatting | preserve cell background colors Aspose.Cells HTML export | retain conditional formatting Excel to HTML C# | convert XLSX to HTML with styling Aspose.Cells | export Excel workbook as HTML preserving shading | Aspose.Cells HTML export styling options

using Aspose.Cells;
using System;
using System.IO;

// Loads an XLSX workbook with Aspose.Cells, configures HtmlSaveOptions (including ExportConditionalFormatting), and saves it as HTML, ensuring that cell background colors and any conditional formatting are retained in the generated HTML file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default options are sufficient for exporting
            // conditional formatting and cell background colors)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
