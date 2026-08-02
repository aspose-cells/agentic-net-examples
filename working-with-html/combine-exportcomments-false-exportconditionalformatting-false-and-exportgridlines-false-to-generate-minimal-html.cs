// Title: Create Minimal HTML from Excel with Aspose.Cells .NET – Disable Comments, Conditional Formatting & Grid Lines
// Description: Loads an Excel workbook, sets HtmlSaveOptions.IsExportComments, ExportConditionalFormatting and ExportGridLines to false, and saves the file as HTML. The result is a lightweight HTML page that contains only raw cell data, without comments, conditional formatting rules, or grid lines.
// Keywords: Aspose.Cells minimal HTML export | C# export Excel to HTML without comments | disable conditional formatting Aspose.Cells | remove grid lines HTMLSaveOptions | .NET Excel to HTML lightweight | HtmlSaveOptions IsExportComments false | ExportConditionalFormatting false
// Common Searches: Aspose.Cells export Excel to HTML without comments | How to turn off conditional formatting when saving HTML with Aspose.Cells | C# remove grid lines in HTML export using Aspose.Cells | Minimal HTML output from Excel using Aspose.Cells .NET | HtmlSaveOptions example for lightweight HTML
// Developer Intent: Generate an HTML file from an Excel workbook that excludes cell comments, conditional formatting, and grid lines.
// Use Cases: Publish clean data tables on web pages where styling is handled by CSS. | Create compact HTML email attachments that contain only values. | Produce HTML snapshots for automated UI tests without visual noise.
// AI Prompts: Provide C# code using Aspose.Cells to export an Excel workbook to HTML with IsExportComments, ExportConditionalFormatting, and ExportGridLines all set to false. | Explain the impact of each HtmlSaveOptions property on the size and appearance of the generated HTML. | Outline steps to verify that comments, conditional formatting, and grid lines are absent in the saved HTML file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Loads an Excel workbook, sets HtmlSaveOptions.IsExportComments, ExportConditionalFormatting and ExportGridLines to false, and saves the file as HTML. The result is a lightweight HTML page that contains only raw cell data, without comments, conditional formatting rules, or grid lines.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            try
            {
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options for minimal output
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Do not export cell comments
                    IsExportComments = false,
                    // Do not export grid lines
                    ExportGridLines = false
                };

                // Save the workbook as HTML with the specified options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
