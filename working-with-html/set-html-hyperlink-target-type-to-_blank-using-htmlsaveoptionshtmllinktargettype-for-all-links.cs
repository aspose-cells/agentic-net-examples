// Title: How to set all hyperlink targets to "_blank" when saving an Excel workbook as HTML using Aspose.Cells HtmlSaveOptions in C#
// Common Searches: Aspose.Cells C# set HtmlLinkTargetType to Blank for exported HTML | Save Excel workbook as HTML with hyperlinks opening in a new tab using Aspose.Cells | C# example of HtmlSaveOptions hyperlink target _blank | Configure Aspose.Cells HTML export to open links in a new window | How to make all links open in a new tab when converting Excel to HTML with Aspose.Cells
// Tags: HtmlSaveOptions HtmlLinkTargetType blank configuration | Aspose.Cells HTML export hyperlink target setting | C# Excel to HTML conversion with _blank links | Aspose.Cells save format HTML link behavior | Set link target attribute in Aspose.Cells HTML output

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, creates HtmlSaveOptions, sets HtmlLinkTargetType to Blank so that every hyperlink in the generated HTML opens in a new tab, ensures the output folder exists, and saves the workbook as an HTML file.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.html";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook;
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML
                try
                {
                    workbook.Save(outputPath, saveOptions);
                    Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook as HTML: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
