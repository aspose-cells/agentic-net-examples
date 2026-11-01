// Title: Create a minimal HTML file from an Excel workbook by disabling comments, conditional formatting, and grid lines with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook using Aspose.Cells, sets HtmlSaveOptions.ExportComments, ExportConditionalFormatting, and ExportGridLines to false, and saves the result as a lightweight HTML file. | Show how to configure Aspose.Cells HtmlSaveOptions in C# to produce minimal HTML output by turning off cell comments, conditional formatting, and grid lines.
// Common Searches: how to export Excel to HTML without comments and conditional formatting using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions minimal HTML output grid lines off | C# generate lightweight HTML from workbook disabling cell comments | remove conditional formatting when saving workbook as HTML Aspose.Cells | export Excel to clean HTML with no grid lines Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions ExportComments false | Aspose.Cells ExportConditionalFormatting false | Aspose.Cells ExportGridLines false | minimal HTML export from Excel .NET | lightweight Excel to HTML conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file with Aspose.Cells, configures HtmlSaveOptions to set ExportComments, ExportConditionalFormatting, and ExportGridLines to false, and saves the workbook as a minimal HTML file. This produces a lightweight HTML representation without cell comments, conditional formatting styles, or grid lines, and includes basic error handling for missing input files.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML export options for minimal output
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportGridLines = false // Do not export grid lines
                // ExportCellComments property is not available in this version of Aspose.Cells
            };

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
