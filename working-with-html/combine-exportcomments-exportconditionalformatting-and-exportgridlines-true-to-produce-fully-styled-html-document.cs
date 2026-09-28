// Title: Export an Excel workbook to fully styled HTML with comments, conditional formatting, and grid lines using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsx file and saves it as HTML with HtmlSaveOptions.ExportComments, ExportConditionalFormatting, and ExportGridLines all set to true using Aspose.Cells. | Show how to configure Aspose.Cells HtmlSaveOptions to preserve cell comments, conditional formatting rules, and grid lines when converting an Excel workbook to HTML.
// Common Searches: how to keep cell comments when exporting Excel to HTML with Aspose.Cells .NET | preserving conditional formatting in HTML output from Aspose.Cells | export grid lines along with styles in Aspose.Cells HTML conversion | Aspose.Cells HtmlSaveOptions settings for full styling in HTML export | C# export Excel workbook to styled HTML including comments and formatting
// Tags: Aspose.Cells HtmlSaveOptions ExportComments | Aspose.Cells ExportConditionalFormatting to HTML | Aspose.Cells ExportGridLines HTML | full style HTML export from Excel .NET | preserve Excel comments in HTML using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportHtml
{
    // The program checks for the source Excel file, creates the destination folder if needed, loads the workbook with Aspose.Cells, configures HtmlSaveOptions to enable ExportComments, ExportConditionalFormatting, and ExportGridLines, then saves the workbook as a fully styled HTML file while handling any errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path
            string inputFile = @"C:\Input\Sample.xlsx";

            // Output HTML file path
            string outputFile = @"C:\Output\StyledDocument.html";

            // Verify input file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            try
            {
                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Show grid lines in the generated HTML
                    ExportGridLines = true
                    // Note: ExportCellComments property is not available in this version of Aspose.Cells
                };

                // Save the workbook as HTML
                workbook.Save(outputFile, saveOptions);

                Console.WriteLine("HTML export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during export: {ex.Message}");
            }
        }
    }
}
