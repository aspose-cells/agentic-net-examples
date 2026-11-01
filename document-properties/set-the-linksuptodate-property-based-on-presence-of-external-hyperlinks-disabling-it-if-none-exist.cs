// Title: How to toggle the LinksUpToDate property in Aspose.Cells C# after detecting external hyperlinks
// Common Searches: Aspose.Cells C# set LinksUpToDate to false when workbook has no external hyperlinks | detect external http hyperlinks in Excel using Aspose.Cells and control link update behavior | C# iterate worksheets with Aspose.Cells to find external links and disable automatic link updates | check for external hyperlinks in Aspose.Cells workbook before saving | conditionally enable LinksUpToDate property in Aspose.Cells based on hyperlink presence
// Tags: LinksUpToDate toggle Aspose.Cells | external hyperlink detection Aspose.Cells | automatic link update control Excel C# | worksheet hyperlink enumeration Aspose.Cells | conditional workbook settings based on links

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook with Aspose.Cells, scans each worksheet for external hyperlinks (http/https), determines whether any exist, and provides a place to set the LinksUpToDate property accordingly before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Determine if any worksheet contains an external hyperlink
                bool hasExternalHyperlink = false;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Hyperlink link in sheet.Hyperlinks)
                    {
                        // Treat hyperlinks starting with http/https as external
                        if (!string.IsNullOrEmpty(link.Address) &&
                            (link.Address.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                             link.Address.StartsWith("https://", StringComparison.OrdinalIgnoreCase)))
                        {
                            hasExternalHyperlink = true;
                            break; // Exit inner loop
                        }
                    }

                    if (hasExternalHyperlink)
                        break; // Exit outer loop
                }

                // If needed, you can adjust workbook settings here based on hasExternalHyperlink.
                // The specific setting to control link updates may vary between Aspose.Cells versions.

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
