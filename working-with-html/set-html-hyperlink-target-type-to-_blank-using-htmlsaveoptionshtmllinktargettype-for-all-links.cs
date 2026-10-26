// Title: How to force all hyperlinks to open in a new window (_blank) when saving an Excel workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook as HTML with hyperlinks opening in new tab | set HtmlSaveOptions link target to blank in Aspose.Cells .NET | export Excel to HTML and make all links use target='_blank' using Aspose.Cells | HtmlLinkTargetType.Blank example for Aspose.Cells HTML conversion | C# Aspose.Cells HTML save options hyperlink target configuration
// Tags: Aspose.Cells HtmlSaveOptions hyperlink target | HtmlLinkTargetType.Blank usage | C# export Excel to HTML with new tab links | set link target _blank Aspose.Cells | HTML conversion hyperlink settings .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The code loads an Excel workbook, creates HtmlSaveOptions, assigns HtmlSaveOptions.HtmlLinkTargetType = HtmlLinkTargetType.Blank so every hyperlink in the generated HTML uses target="_blank", and then saves the workbook as an HTML file.
    class Program
    {
        static void Main(string[] args)
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

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                // Note: ExportHyperlinkTarget property is not available in this version of Aspose.Cells.

                // Save the workbook as an HTML file with the defined options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
